unit MSGraph.Graph.Mail.Attachments.Tests;

interface

uses
  System.SysUtils,
  DUnitX.TestFramework,
  MSGraph.Graph.Http,
  MSGraph.Graph.Http.Transport.Fake,
  MSGraph.Graph.Mail.Attachments,
  MSGraph.Graph.Mail.Interfaces;

type
  [TestFixture]
  TAttachmentUploadTests = class
  strict private
    FFake: IFakeGraphHttpTransport;
    FGraphClient: TGraphHttpClient;
    FMailClient: IMailClient;
    FLogLines: TArray<string>;

    class function MakeBytes(const Size: Integer): TBytes; static;
    class function MakeZeroedBytes(const Size: Integer): TBytes; static;
    class function SessionResponse: string; static;
    class function NextRangeResponse(const NextOffset: Int64): string; static;
    class function ErrorResponse(const Code: string; const Message: string): string; static;

    procedure EnqueueUploadSession;
    procedure EnqueueChunkAccepted(const NextOffset: Int64);
    procedure EnqueueChunkCompleted;
    procedure EnqueueSessionCancelled;

    function AddAttachment(const ContentBytes: TBytes): Boolean;
    function CapturedUploadError(const ContentBytes: TBytes): string;
    procedure AssertRequest(const RequestIndex: Integer; const ExpectedMethod: string;
      const ExpectedUrlSuffix: string);
    procedure AssertContains(const Text: string; const Fragment: string);
    procedure AssertContentRange(const RequestIndex: Integer; const ExpectedRange: string);
    procedure AssertNoSendRequest;
    function LoggedText: string;
  public
    [Setup]
    procedure Setup;
    [TearDown]
    procedure TearDown;

    [Test]
    procedure Upload_SmallAttachment_PostsInlineAttachment;
    [Test]
    procedure Upload_JustBelowThreshold_PostsInlineAttachment;
    [Test]
    procedure Upload_AtThreshold_CreatesUploadSession;
    [Test]
    procedure Upload_EmptyAttachment_RaisesBeforeAnyRequest;
    [Test]
    procedure Upload_AboveMaximumSize_RaisesBeforeAnyRequest;
    [Test]
    procedure Upload_InlineAttachmentTooLarge_FallsBackToUploadSessionOnce;
    [Test]
    procedure Upload_SessionBelowMinimumSize_FallsBackToInlineAttachmentOnce;

    [Test]
    procedure Upload_TenMegabytes_SendsFourChunksWithExactContentRanges;
    [Test]
    procedure Upload_SizeIsExactMultipleOfChunk_SendsNoExtraChunk;
    [Test]
    procedure Upload_CustomChunkSize_SplitsOnConfiguredSize;
    [Test]
    procedure Upload_ServerReportsOtherNextRange_ResumesWhereServerAsks;
    [Test]
    procedure Upload_ServerRepeatsSameRange_RaisesInsteadOfLooping;
    [Test]
    procedure Upload_ChunkRequest_OmitsAuthorizationAndAnchorHeaders;
    [Test]
    procedure Upload_FinalChunk_ReturnsAttachmentIdFromLocationHeader;

    [Test]
    procedure Upload_ChunkFails_CancelsSessionAndReportsByteRange;
    [Test]
    procedure Upload_ConnectionDropsMidUpload_CancelsSessionAndReportsByteRange;
    [Test]
    procedure Upload_ChunkUnauthorized_ReportsExpiredSessionWithoutRetry;
    [Test]
    procedure Upload_SharedMailboxForbidden_KeepsGraphMessageAndAddsHint;
    [Test]
    procedure Upload_OwnMailboxForbidden_ReportsGraphMessageWithoutHint;
    [Test]
    procedure Upload_UploadUrlToken_IsNeverLoggedInFull;
  end;

implementation

uses
  System.Net.URLClient,
  MSGraph.OAuth2.Types,
  MSGraph.Graph.Mail,
  MSGraph.Graph.Mail.Types;

const
  DummyAccessToken = 'unit-test-token';
  SharedMailboxAddress = 'projects@example.com';
  MessageId = 'MSG-1';
  MessageEndpointSuffix = '/me/messages/MSG-1';
  AttachmentsSuffix = '/me/messages/MSG-1/attachments';
  UploadSessionSuffix = '/me/messages/MSG-1/attachments/createUploadSession';
  AttachmentName = 'report.pdf';
  AttachmentContentType = 'application/pdf';
  UploadToken = 'SECRET-UPLOAD-TOKEN';
  UploadUrl = 'https://outlook.office.com/api/v2.0/Users(''ab'')/Messages(''cd'')/' +
    'AttachmentSessions(''ef'')?authtoken=' + UploadToken;
  AttachmentLocation = 'https://outlook.office.com/api/v2.0/Users(''ab'')/Messages(''cd'')/' +
    'Attachments(''AAMkAttachment1'')';
  ExpectedAttachmentId = 'AAMkAttachment1';

  MethodPost = 'POST';
  MethodPut = 'PUT';
  MethodDelete = 'DELETE';
  HeaderAuthorization = 'Authorization';
  HeaderAnchorMailbox = 'X-AnchorMailbox';
  HeaderContentType = 'Content-Type';
  HeaderContentRange = 'Content-Range';
  ExpectedChunkContentType = 'application/octet-stream';

  ChunkSize = 3 * 1024 * 1024;
  TenMegabytes = 10 * 1024 * 1024;
  SmallAttachmentSize = 1024 * 1024;
  BytePatternPeriod = 251;

  AttachmentCreatedResponse = '{"id":"attachment-1"}';
  SwitchesOnce = 'the route switches exactly once';
  ConnectionResetMessage = 'Error receiving data: the connection was reset by the peer';
  AccessDeniedMessage = 'Access is denied. Check credentials and try again.';

procedure TAttachmentUploadTests.Setup;
begin
  FLogLines := nil;
  FFake := TFakeGraphHttpTransport.Create;
  FGraphClient := TGraphHttpClient.Create(DummyAccessToken, FFake,
    procedure(const Level: string; const Message: string)
    begin
      FLogLines := FLogLines + [Message];
    end);
  FMailClient := TMailClient.Create(FGraphClient, True);
end;

procedure TAttachmentUploadTests.TearDown;
begin
  FMailClient := nil;
  FGraphClient := nil;
  FFake := nil;
  FLogLines := nil;
end;

class function TAttachmentUploadTests.MakeBytes(const Size: Integer): TBytes;
begin
  SetLength(Result, Size);

  for var Index := 0 to Size - 1 do
  begin
    Result[Index] := Byte(Index mod BytePatternPeriod);
  end;
end;

class function TAttachmentUploadTests.MakeZeroedBytes(const Size: Integer): TBytes;
begin
  SetLength(Result, Size);
end;

class function TAttachmentUploadTests.SessionResponse: string;
begin
  Result := Format('{"uploadUrl":"%s","expirationDateTime":"2026-08-27T10:00:00Z",' +
                   '"nextExpectedRanges":["0-"]}', [UploadUrl]);
end;

class function TAttachmentUploadTests.NextRangeResponse(const NextOffset: Int64): string;
begin
  Result := Format('{"nextExpectedRanges":["%d-"]}', [NextOffset]);
end;

class function TAttachmentUploadTests.ErrorResponse(const Code: string; const Message: string): string;
begin
  Result := Format('{"error":{"code":"%s","message":"%s"}}', [Code, Message]);
end;

procedure TAttachmentUploadTests.EnqueueUploadSession;
begin
  FFake.EnqueueResponse(201, SessionResponse);
end;

procedure TAttachmentUploadTests.EnqueueChunkAccepted(const NextOffset: Int64);
begin
  FFake.EnqueueResponse(200, NextRangeResponse(NextOffset));
end;

procedure TAttachmentUploadTests.EnqueueChunkCompleted;
begin
  FFake.EnqueueResponseWithHeaders(201, '', [TNetHeader.Create('Location', AttachmentLocation)]);
end;

procedure TAttachmentUploadTests.EnqueueSessionCancelled;
begin
  FFake.EnqueueResponse(204, '');
end;

function TAttachmentUploadTests.AddAttachment(const ContentBytes: TBytes): Boolean;
begin
  Result := FMailClient.AddAttachment(MessageId, AttachmentName, AttachmentContentType, ContentBytes);
end;

function TAttachmentUploadTests.CapturedUploadError(const ContentBytes: TBytes): string;
begin
  Result := '';

  try
    AddAttachment(ContentBytes);
    Assert.Fail('expected the upload to fail');
  except
    on E: EGraphApiException do
      Result := E.Message;
  end;
end;

procedure TAttachmentUploadTests.AssertRequest(const RequestIndex: Integer;
  const ExpectedMethod: string; const ExpectedUrlSuffix: string);
begin
  const Request = FFake.RequestAt(RequestIndex);

  Assert.AreEqual(ExpectedMethod, Request.Method,
    Format('unexpected method on request %d', [RequestIndex]));
  Assert.IsTrue(Request.Url.EndsWith(ExpectedUrlSuffix),
    Format('request %d has unexpected url: %s', [RequestIndex, Request.Url]));
end;

procedure TAttachmentUploadTests.AssertContains(const Text: string; const Fragment: string);
begin
  Assert.IsTrue(Text.Contains(Fragment), Format('"%s" not found in: %s', [Fragment, Text]));
end;

procedure TAttachmentUploadTests.AssertContentRange(const RequestIndex: Integer;
  const ExpectedRange: string);
begin
  Assert.AreEqual(ExpectedRange, FFake.HeaderValue(RequestIndex, HeaderContentRange),
    Format('unexpected content range on request %d', [RequestIndex]));
end;

procedure TAttachmentUploadTests.AssertNoSendRequest;
begin
  for var Index := 0 to FFake.RequestCount - 1 do
  begin
    const Request = FFake.RequestAt(Index);
    Assert.IsFalse(Request.Url.EndsWith('/send'),
      Format('the draft must not be sent, but request %d did: %s', [Index, Request.Url]));
  end;
end;

function TAttachmentUploadTests.LoggedText: string;
begin
  Result := string.Join(sLineBreak, FLogLines);
end;

procedure TAttachmentUploadTests.Upload_SmallAttachment_PostsInlineAttachment;
begin
  FFake.EnqueueResponse(201, AttachmentCreatedResponse);

  Assert.IsTrue(AddAttachment(MakeBytes(SmallAttachmentSize)));

  Assert.AreEqual(1, FFake.RequestCount, 'a small attachment needs a single request');
  AssertRequest(0, MethodPost, AttachmentsSuffix);

  const PostedBody = FFake.RequestAt(0).Body;
  AssertContains(PostedBody, '#microsoft.graph.fileAttachment');
end;

procedure TAttachmentUploadTests.Upload_JustBelowThreshold_PostsInlineAttachment;
begin
  FFake.EnqueueResponse(201, AttachmentCreatedResponse);

  Assert.IsTrue(AddAttachment(MakeZeroedBytes(LargeAttachmentThreshold - 1)));

  Assert.AreEqual(1, FFake.RequestCount);
  AssertRequest(0, MethodPost, AttachmentsSuffix);
end;

procedure TAttachmentUploadTests.Upload_AtThreshold_CreatesUploadSession;
begin
  EnqueueUploadSession;
  EnqueueChunkCompleted;

  Assert.IsTrue(AddAttachment(MakeZeroedBytes(LargeAttachmentThreshold)));

  AssertRequest(0, MethodPost, UploadSessionSuffix);
  AssertRequest(1, MethodPut, UploadUrl);
end;

procedure TAttachmentUploadTests.Upload_EmptyAttachment_RaisesBeforeAnyRequest;
begin
  Assert.WillRaise(
    procedure
    begin
      AddAttachment(nil);
    end,
    EInvalidAttachmentException);

  Assert.AreEqual(0, FFake.RequestCount, 'an empty attachment is rejected without any request');
end;

procedure TAttachmentUploadTests.Upload_AboveMaximumSize_RaisesBeforeAnyRequest;
begin
  Assert.WillRaise(
    procedure
    begin
      AddAttachment(MakeZeroedBytes(MaxAttachmentSize + 1));
    end,
    EInvalidAttachmentException);

  Assert.AreEqual(0, FFake.RequestCount, 'an oversized attachment is rejected without any request');
end;

procedure TAttachmentUploadTests.Upload_InlineAttachmentTooLarge_FallsBackToUploadSessionOnce;
begin
  FFake.EnqueueResponse(413, ErrorResponse('RequestEntityTooLarge', 'Request entity too large'));
  EnqueueUploadSession;
  EnqueueChunkCompleted;

  Assert.IsTrue(AddAttachment(MakeBytes(1024)));

  Assert.AreEqual(3, FFake.RequestCount, SwitchesOnce);
  AssertRequest(0, MethodPost, AttachmentsSuffix);
  AssertRequest(1, MethodPost, UploadSessionSuffix);
  AssertRequest(2, MethodPut, UploadUrl);
end;

procedure TAttachmentUploadTests.Upload_SessionBelowMinimumSize_FallsBackToInlineAttachmentOnce;
begin
  FFake.EnqueueResponse(400, ErrorResponse('ErrorAttachmentSizeShouldNotBeLessThanMinimumSize',
    'Attachment size must be greater than the minimum size'));
  FFake.EnqueueResponse(201, AttachmentCreatedResponse);

  Assert.IsTrue(AddAttachment(MakeZeroedBytes(LargeAttachmentThreshold)));

  Assert.AreEqual(2, FFake.RequestCount, SwitchesOnce);
  AssertRequest(0, MethodPost, UploadSessionSuffix);
  AssertRequest(1, MethodPost, AttachmentsSuffix);
end;

procedure TAttachmentUploadTests.Upload_TenMegabytes_SendsFourChunksWithExactContentRanges;
begin
  EnqueueUploadSession;
  EnqueueChunkAccepted(ChunkSize);
  EnqueueChunkAccepted(2 * ChunkSize);
  EnqueueChunkAccepted(3 * ChunkSize);
  EnqueueChunkCompleted;

  Assert.IsTrue(AddAttachment(MakeBytes(TenMegabytes)));

  Assert.AreEqual(5, FFake.RequestCount, 'one session plus four chunks');
  AssertContentRange(1, 'bytes 0-3145727/10485760');
  AssertContentRange(2, 'bytes 3145728-6291455/10485760');
  AssertContentRange(3, 'bytes 6291456-9437183/10485760');
  AssertContentRange(4, 'bytes 9437184-10485759/10485760');

  Assert.AreEqual(ChunkSize, Length(FFake.RequestAt(1).BodyBytes));
  Assert.AreEqual(TenMegabytes - (3 * ChunkSize), Length(FFake.RequestAt(4).BodyBytes));
  Assert.AreEqual(Integer((3 * ChunkSize) mod 251), Integer(FFake.RequestAt(4).BodyBytes[0]),
    'the last chunk starts at the byte right after the third chunk');
end;

procedure TAttachmentUploadTests.Upload_SizeIsExactMultipleOfChunk_SendsNoExtraChunk;
begin
  EnqueueUploadSession;
  EnqueueChunkAccepted(ChunkSize);
  EnqueueChunkCompleted;

  Assert.IsTrue(AddAttachment(MakeZeroedBytes(2 * ChunkSize)));

  const LastRange = Format('bytes %d-%d/%d', [ChunkSize, (2 * ChunkSize) - 1, 2 * ChunkSize]);

  Assert.AreEqual(3, FFake.RequestCount, 'a size on the chunk boundary needs no empty final chunk');
  AssertContentRange(2, LastRange);
end;

procedure TAttachmentUploadTests.Upload_CustomChunkSize_SplitsOnConfiguredSize;
begin
  const CustomChunkSize = 4 * 1024 * 1024;

  EnqueueUploadSession;
  EnqueueChunkAccepted(CustomChunkSize);
  EnqueueChunkAccepted(2 * CustomChunkSize);
  EnqueueChunkCompleted;

  var Uploader := TAttachmentUploader.Create(FGraphClient, CustomChunkSize);
  try
    Uploader.Upload(MessageEndpointSuffix, AttachmentName, AttachmentContentType,
      MakeZeroedBytes(TenMegabytes));
  finally
    Uploader.Free;
  end;

  const FirstRange = Format('bytes 0-%d/%d', [CustomChunkSize - 1, TenMegabytes]);
  const LastRange = Format('bytes %d-%d/%d',
                           [2 * CustomChunkSize, TenMegabytes - 1, TenMegabytes]);

  Assert.AreEqual(4, FFake.RequestCount, 'one session plus three chunks');
  AssertContentRange(1, FirstRange);
  AssertContentRange(3, LastRange);
end;

procedure TAttachmentUploadTests.Upload_ServerReportsOtherNextRange_ResumesWhereServerAsks;
begin
  const ServerOffset = 1000000;

  EnqueueUploadSession;
  EnqueueChunkAccepted(ServerOffset);
  EnqueueChunkCompleted;

  Assert.IsTrue(AddAttachment(MakeZeroedBytes(LargeAttachmentThreshold)));

  const ResumedRange = Format('bytes %d-%d/%d',
                              [ServerOffset, LargeAttachmentThreshold - 1, LargeAttachmentThreshold]);
  AssertContentRange(2, ResumedRange);
end;

procedure TAttachmentUploadTests.Upload_ServerRepeatsSameRange_RaisesInsteadOfLooping;
begin
  EnqueueUploadSession;
  EnqueueChunkAccepted(0);
  EnqueueSessionCancelled;

  const ErrorMessage = CapturedUploadError(MakeZeroedBytes(LargeAttachmentThreshold));

  AssertContains(ErrorMessage, 'stalled');
  AssertContains(ErrorMessage, AttachmentName);
  Assert.AreEqual(3, FFake.RequestCount, 'the stalled upload stops after cancelling the session');
  AssertRequest(2, MethodDelete, UploadUrl);
  AssertNoSendRequest;
end;

procedure TAttachmentUploadTests.Upload_ChunkRequest_OmitsAuthorizationAndAnchorHeaders;
begin
  FGraphClient.MailboxAddress := SharedMailboxAddress;

  EnqueueUploadSession;
  EnqueueChunkCompleted;

  Assert.IsTrue(AddAttachment(MakeZeroedBytes(LargeAttachmentThreshold)));

  Assert.AreEqual('', FFake.HeaderValue(1, HeaderAuthorization),
    'the upload url is pre-authenticated and must not carry a bearer token');
  Assert.AreEqual('', FFake.HeaderValue(1, HeaderAnchorMailbox),
    'the upload url is not a Graph endpoint');
  Assert.AreEqual(ExpectedChunkContentType, FFake.HeaderValue(1, HeaderContentType));
  Assert.AreEqual(UploadUrl, FFake.RequestAt(1).Url, 'the upload url is used verbatim');
end;

procedure TAttachmentUploadTests.Upload_FinalChunk_ReturnsAttachmentIdFromLocationHeader;
begin
  EnqueueUploadSession;
  EnqueueChunkCompleted;

  var Uploader := TAttachmentUploader.Create(FGraphClient);
  try
    const AttachmentId = Uploader.Upload(MessageEndpointSuffix, AttachmentName,
      AttachmentContentType, MakeZeroedBytes(LargeAttachmentThreshold));

    Assert.AreEqual(ExpectedAttachmentId, AttachmentId);
  finally
    Uploader.Free;
  end;
end;

procedure TAttachmentUploadTests.Upload_ChunkFails_CancelsSessionAndReportsByteRange;
begin
  EnqueueUploadSession;
  EnqueueChunkAccepted(ChunkSize);
  EnqueueChunkAccepted(2 * ChunkSize);
  FFake.EnqueueResponse(500, ErrorResponse('InternalServerError', 'Something went wrong'));
  EnqueueSessionCancelled;

  const ErrorMessage = CapturedUploadError(MakeZeroedBytes(TenMegabytes));

  AssertContains(ErrorMessage, AttachmentName);
  AssertContains(ErrorMessage, '10485760 bytes');
  AssertContains(ErrorMessage, 'bytes 6291456-9437183');
  AssertContains(ErrorMessage, 'HTTP 500');
  AssertContains(ErrorMessage, 'The draft has not been sent.');

  Assert.AreEqual(5, FFake.RequestCount, 'three chunks, then the session is cancelled');
  AssertRequest(4, MethodDelete, UploadUrl);
  AssertNoSendRequest;
end;

procedure TAttachmentUploadTests.Upload_ConnectionDropsMidUpload_CancelsSessionAndReportsByteRange;
begin
  EnqueueUploadSession;
  EnqueueChunkAccepted(ChunkSize);
  FFake.EnqueueTransportFailure(ConnectionResetMessage);
  EnqueueSessionCancelled;

  const ErrorMessage = CapturedUploadError(MakeZeroedBytes(TenMegabytes));

  AssertContains(ErrorMessage, AttachmentName);
  AssertContains(ErrorMessage, '10485760 bytes');
  AssertContains(ErrorMessage, 'bytes 3145728-6291455');
  AssertContains(ErrorMessage, ConnectionResetMessage);
  AssertContains(ErrorMessage, 'The draft has not been sent.');

  Assert.AreEqual(4, FFake.RequestCount, 'two chunks, then the session is cancelled');
  AssertRequest(3, MethodDelete, UploadUrl);
  AssertNoSendRequest;
end;

procedure TAttachmentUploadTests.Upload_ChunkUnauthorized_ReportsExpiredSessionWithoutRetry;
begin
  EnqueueUploadSession;
  FFake.EnqueueResponse(401, ErrorResponse('InvalidAuthenticationToken', 'Access token has expired'));
  EnqueueSessionCancelled;

  const ErrorMessage = CapturedUploadError(MakeZeroedBytes(LargeAttachmentThreshold));

  AssertContains(ErrorMessage, 'the upload session has expired');
  AssertContains(ErrorMessage, 'HTTP 401');

  Assert.AreEqual(3, FFake.RequestCount, 'a single chunk attempt, then the session is cancelled');
  AssertRequest(1, MethodPut, UploadUrl);
  AssertRequest(2, MethodDelete, UploadUrl);
  AssertNoSendRequest;
end;

procedure TAttachmentUploadTests.Upload_SharedMailboxForbidden_KeepsGraphMessageAndAddsHint;
begin
  FGraphClient.MailboxAddress := SharedMailboxAddress;

  FFake.EnqueueResponse(403, ErrorResponse('ErrorAccessDenied', AccessDeniedMessage));

  const ErrorMessage = CapturedUploadError(MakeZeroedBytes(LargeAttachmentThreshold));

  AssertContains(ErrorMessage, AccessDeniedMessage);
  AssertContains(ErrorMessage, 'HTTP 403');
  AssertContains(ErrorMessage, 'shared');
  AssertContains(ErrorMessage, 'application permissions');
  Assert.AreEqual(1, FFake.RequestCount, 'no session was created, so nothing needs cancelling');
  AssertNoSendRequest;
end;

procedure TAttachmentUploadTests.Upload_OwnMailboxForbidden_ReportsGraphMessageWithoutHint;
begin
  FFake.EnqueueResponse(403, ErrorResponse('ErrorAccessDenied', AccessDeniedMessage));

  const ErrorMessage = CapturedUploadError(MakeZeroedBytes(LargeAttachmentThreshold));

  AssertContains(ErrorMessage, AccessDeniedMessage);
  Assert.IsFalse(ErrorMessage.Contains('shared'),
    'the shared mailbox hint belongs only to a request against another mailbox');
  AssertNoSendRequest;
end;

procedure TAttachmentUploadTests.Upload_UploadUrlToken_IsNeverLoggedInFull;
begin
  EnqueueUploadSession;
  EnqueueChunkCompleted;

  Assert.IsTrue(AddAttachment(MakeZeroedBytes(LargeAttachmentThreshold)));

  const Logged = LoggedText;
  Assert.IsFalse(Logged.Contains(UploadToken),
    Format('the upload token must never reach the log: %s', [Logged]));
  AssertContains(Logged, 'authtoken=***');
end;

end.
