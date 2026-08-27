unit MSGraph.Graph.Mail.Attachments;

interface

uses
  System.SysUtils,
  MSGraph.Graph.Http,
  MSGraph.Graph.Http.Types;

const
  LargeAttachmentThreshold = 3 * 1024 * 1024;
  MaxAttachmentSize = 150 * 1024 * 1024;
  DefaultUploadChunkSize = 3 * 1024 * 1024;

type
  TAttachmentUploader = class
  strict private
    FGraphClient: TGraphHttpClient;
    FChunkSize: Integer;

    procedure ValidateAttachment(const FileName: string; const TotalSize: Int64);

    function UploadInline(const MessageEndpoint: string; const FileName: string;
      const ContentType: string; const ContentBytes: TBytes): string;
    function UploadChunked(const MessageEndpoint: string; const FileName: string;
      const ContentType: string; const ContentBytes: TBytes): string;

    function AddInlineAttachment(const MessageEndpoint: string; const FileName: string;
      const ContentType: string; const ContentBytes: TBytes): string;
    function UploadWithNewSession(const MessageEndpoint: string; const FileName: string;
      const ContentType: string; const ContentBytes: TBytes): string;

    function PostInlineAttachment(const MessageEndpoint: string; const FileName: string;
      const ContentType: string; const ContentBytes: TBytes): TGraphHttpResponse;
    function StartUploadSession(const MessageEndpoint: string; const FileName: string;
      const ContentType: string; const TotalSize: Int64): TGraphHttpResponse;
    function RunUploadSession(const SessionResponse: TGraphHttpResponse; const FileName: string;
      const ContentBytes: TBytes): string;

    function EndpointFor(const MessageEndpoint: string; const Suffix: string): string;
    function BuildInlineAttachmentBody(const FileName: string; const ContentType: string;
      const ContentBytes: TBytes): string;
    function BuildUploadSessionBody(const FileName: string; const ContentType: string;
      const TotalSize: Int64): string;

    function UploadChunks(const UploadUrl: string; const FileName: string;
      const ContentBytes: TBytes): string;
    function SendChunkOrFail(const UploadUrl: string; const FileName: string;
      const ContentBytes: TBytes; const Offset, ChunkLength, TotalSize: Int64): TGraphHttpResponse;
    function SendChunk(const UploadUrl: string; const ContentBytes: TBytes;
      const Offset, ChunkLength, TotalSize: Int64): TGraphHttpResponse;
    function BuildContentRange(const RangeStart, RangeEnd, TotalSize: Int64): string;
    function AcceptedOffset(const UploadUrl: string; const FileName: string;
      const Response: TGraphHttpResponse; const CurrentOffset, ExpectedOffset,
      TotalSize: Int64): Int64;
    function ReportedNextOffset(const ResponseContent: string; const FallbackOffset: Int64): Int64;

    function AttachmentIdOrFail(const Response: TGraphHttpResponse; const FileName: string;
      const TotalSize: Int64): string;
    function UploadUrlOrFail(const Response: TGraphHttpResponse; const FileName: string;
      const TotalSize: Int64): string;
    function AttachmentIdFromLocation(const Location: string): string;
    function UnwrapODataKey(const Segment: string): string;

    procedure FailChunkUpload(const UploadUrl: string; const FileName: string;
      const RangeStart, RangeEnd, TotalSize: Int64; const Response: TGraphHttpResponse);
    procedure FailInterruptedUpload(const UploadUrl: string; const FileName: string;
      const RangeStart, RangeEnd, TotalSize: Int64; const Error: Exception);
    procedure FailStalledUpload(const UploadUrl: string; const FileName: string;
      const ResumeOffset, TotalSize: Int64);
    procedure FailUnconfirmedUpload(const UploadUrl: string; const FileName: string;
      const TotalSize: Int64);
    procedure CancelUploadSession(const UploadUrl: string);

    function IsBelowUploadSessionMinimum(const Response: TGraphHttpResponse): Boolean;
    function DescribeSessionFailure(const FileName: string; const TotalSize: Int64;
      const Response: TGraphHttpResponse): string;
    function SessionFailureReason(const Response: TGraphHttpResponse): string;
    function ChunkFailureReason(const Response: TGraphHttpResponse): string;
    function FailureReason(const Response: TGraphHttpResponse): string;
    function ErrorMember(const ResponseContent: string; const MemberName: string): string;
  public
    constructor Create(const GraphClient: TGraphHttpClient;
      const ChunkSize: Integer = DefaultUploadChunkSize);

    function Upload(const MessageEndpoint: string; const FileName: string;
      const ContentType: string; const ContentBytes: TBytes): string;
  end;

implementation

uses
  System.Math,
  System.JSON,
  System.NetEncoding,
  MSGraph.OAuth2.Types,
  MSGraph.Graph.JsonHelper,
  MSGraph.Graph.Mail.Types;

const
  HttpOk = 200;
  HttpCreated = 201;
  HttpUnauthorized = 401;
  HttpForbidden = 403;
  HttpRequestEntityTooLarge = 413;

  UploadSessionEndpointSuffix = '/attachments/createUploadSession';
  AttachmentsEndpointSuffix = '/attachments';
  OdataTypeFileAttachment = '#microsoft.graph.fileAttachment';
  AttachmentTypeFile = 'file';
  HeaderLocation = 'Location';
  LogLevelError = 'ERROR';

  ErrorCodeAttachmentBelowMinimum = 'ErrorAttachmentSizeShouldNotBeLessThanMinimumSize';
  SessionExpiredReason = 'the upload session has expired';
  SharedMailboxRestrictionHint = 'If this is a shared or delegated mailbox: Microsoft Graph does ' +
    'not allow large attachments there while the application uses delegated permissions, so use ' +
    'application permissions instead';

constructor TAttachmentUploader.Create(const GraphClient: TGraphHttpClient; const ChunkSize: Integer);
begin
  inherited Create;

  if not Assigned(GraphClient) then
    raise EGraphApiException.Create('No Graph HTTP client provided.');

  const IsUsableChunkSize = (ChunkSize > 0);
  if not IsUsableChunkSize then
    raise EGraphApiException.CreateFmt('Invalid upload chunk size: %d.', [ChunkSize]);

  FGraphClient := GraphClient;
  FChunkSize := ChunkSize;
end;

function TAttachmentUploader.Upload(const MessageEndpoint: string; const FileName: string;
  const ContentType: string; const ContentBytes: TBytes): string;
begin
  const TotalSize = Int64(Length(ContentBytes));
  ValidateAttachment(FileName, TotalSize);

  const NeedsUploadSession = (TotalSize >= LargeAttachmentThreshold);
  if NeedsUploadSession then
    Result := UploadChunked(MessageEndpoint, FileName, ContentType, ContentBytes)
  else
    Result := UploadInline(MessageEndpoint, FileName, ContentType, ContentBytes);
end;

procedure TAttachmentUploader.ValidateAttachment(const FileName: string; const TotalSize: Int64);
begin
  const IsEmpty = (TotalSize <= 0);
  if IsEmpty then
    raise EInvalidAttachmentException.CreateFmt(
      'Attachment "%s" is empty; there is nothing to upload.', [FileName]);

  const ExceedsMaximum = (TotalSize > MaxAttachmentSize);
  if ExceedsMaximum then
    raise EInvalidAttachmentException.CreateFmt(
      'Attachment "%s" is %d bytes, which exceeds the maximum supported size of %d bytes.',
      [FileName, TotalSize, Int64(MaxAttachmentSize)]);
end;

function TAttachmentUploader.UploadInline(const MessageEndpoint: string; const FileName: string;
  const ContentType: string; const ContentBytes: TBytes): string;
begin
  const TotalSize = Int64(Length(ContentBytes));
  const Response = PostInlineAttachment(MessageEndpoint, FileName, ContentType, ContentBytes);

  const IsTooLargeForMessageBody = (Response.StatusCode = HttpRequestEntityTooLarge);
  if IsTooLargeForMessageBody then
    Result := UploadWithNewSession(MessageEndpoint, FileName, ContentType, ContentBytes)
  else
    Result := AttachmentIdOrFail(Response, FileName, TotalSize);
end;

function TAttachmentUploader.UploadChunked(const MessageEndpoint: string; const FileName: string;
  const ContentType: string; const ContentBytes: TBytes): string;
begin
  const TotalSize = Int64(Length(ContentBytes));
  const SessionResponse = StartUploadSession(MessageEndpoint, FileName, ContentType, TotalSize);

  if IsBelowUploadSessionMinimum(SessionResponse) then
    Result := AddInlineAttachment(MessageEndpoint, FileName, ContentType, ContentBytes)
  else
    Result := RunUploadSession(SessionResponse, FileName, ContentBytes);
end;

function TAttachmentUploader.AddInlineAttachment(const MessageEndpoint: string; const FileName: string;
  const ContentType: string; const ContentBytes: TBytes): string;
begin
  const TotalSize = Int64(Length(ContentBytes));
  const Response = PostInlineAttachment(MessageEndpoint, FileName, ContentType, ContentBytes);

  Result := AttachmentIdOrFail(Response, FileName, TotalSize);
end;

function TAttachmentUploader.UploadWithNewSession(const MessageEndpoint: string; const FileName: string;
  const ContentType: string; const ContentBytes: TBytes): string;
begin
  const TotalSize = Int64(Length(ContentBytes));
  const SessionResponse = StartUploadSession(MessageEndpoint, FileName, ContentType, TotalSize);

  Result := RunUploadSession(SessionResponse, FileName, ContentBytes);
end;

function TAttachmentUploader.PostInlineAttachment(const MessageEndpoint: string; const FileName: string;
  const ContentType: string; const ContentBytes: TBytes): TGraphHttpResponse;
begin
  const Endpoint = EndpointFor(MessageEndpoint, AttachmentsEndpointSuffix);
  const Body = BuildInlineAttachmentBody(FileName, ContentType, ContentBytes);

  Result := FGraphClient.PostRaw(Endpoint, Body);
end;

function TAttachmentUploader.StartUploadSession(const MessageEndpoint: string; const FileName: string;
  const ContentType: string; const TotalSize: Int64): TGraphHttpResponse;
begin
  const Endpoint = EndpointFor(MessageEndpoint, UploadSessionEndpointSuffix);
  const Body = BuildUploadSessionBody(FileName, ContentType, TotalSize);

  Result := FGraphClient.PostRaw(Endpoint, Body);
end;

function TAttachmentUploader.RunUploadSession(const SessionResponse: TGraphHttpResponse;
  const FileName: string; const ContentBytes: TBytes): string;
begin
  const TotalSize = Int64(Length(ContentBytes));
  const UploadUrl = UploadUrlOrFail(SessionResponse, FileName, TotalSize);

  Result := UploadChunks(UploadUrl, FileName, ContentBytes);
end;

function TAttachmentUploader.EndpointFor(const MessageEndpoint: string; const Suffix: string): string;
begin
  Result := Format('%s%s', [MessageEndpoint, Suffix]);
end;

function TAttachmentUploader.BuildInlineAttachmentBody(const FileName: string;
  const ContentType: string; const ContentBytes: TBytes): string;
begin
  var AttachmentObj := TJSONObject.Create;
  try
    AttachmentObj.AddPair('@odata.type', OdataTypeFileAttachment);
    AttachmentObj.AddPair('name', FileName);
    AttachmentObj.AddPair('contentType', ContentType);
    AttachmentObj.AddPair('contentBytes', TNetEncoding.Base64.EncodeBytesToString(ContentBytes));

    Result := AttachmentObj.ToJSON;
  finally
    AttachmentObj.Free;
  end;
end;

function TAttachmentUploader.BuildUploadSessionBody(const FileName: string;
  const ContentType: string; const TotalSize: Int64): string;
begin
  var SessionObj := TJSONObject.Create;
  try
    var AttachmentItem := TJSONObject.Create;
    SessionObj.AddPair('AttachmentItem', AttachmentItem);

    AttachmentItem.AddPair('attachmentType', AttachmentTypeFile);
    AttachmentItem.AddPair('name', FileName);
    AttachmentItem.AddPair('contentType', ContentType);
    AttachmentItem.AddPair('size', TJSONNumber.Create(TotalSize));

    Result := SessionObj.ToJSON;
  finally
    SessionObj.Free;
  end;
end;

function TAttachmentUploader.UploadChunks(const UploadUrl: string; const FileName: string;
  const ContentBytes: TBytes): string;
begin
  Result := '';

  const TotalSize = Int64(Length(ContentBytes));
  var Offset: Int64 := 0;

  while Offset < TotalSize do
  begin
    const RemainingBytes = TotalSize - Offset;
    const ChunkLength = Min(Int64(FChunkSize), RemainingBytes);
    const RangeEnd = Offset + ChunkLength - 1;
    const Response = SendChunkOrFail(UploadUrl, FileName, ContentBytes, Offset, ChunkLength,
                                     TotalSize);

    case Response.StatusCode of
      HttpOk:
        Offset := AcceptedOffset(UploadUrl, FileName, Response, Offset, RangeEnd + 1, TotalSize);
      HttpCreated:
        begin
          const Location = Response.HeaderValue(HeaderLocation);

          Result := AttachmentIdFromLocation(Location);
          Exit;
        end;
    else
      FailChunkUpload(UploadUrl, FileName, Offset, RangeEnd, TotalSize, Response);
    end;
  end;

  FailUnconfirmedUpload(UploadUrl, FileName, TotalSize);
end;

function TAttachmentUploader.SendChunkOrFail(const UploadUrl: string; const FileName: string;
  const ContentBytes: TBytes; const Offset, ChunkLength, TotalSize: Int64): TGraphHttpResponse;
begin
  Result := Default(TGraphHttpResponse);

  const RangeEnd = Offset + ChunkLength - 1;

  try
    Result := SendChunk(UploadUrl, ContentBytes, Offset, ChunkLength, TotalSize);
  except
    on E: Exception do
      FailInterruptedUpload(UploadUrl, FileName, Offset, RangeEnd, TotalSize, E);
  end;
end;

function TAttachmentUploader.SendChunk(const UploadUrl: string; const ContentBytes: TBytes;
  const Offset, ChunkLength, TotalSize: Int64): TGraphHttpResponse;
begin
  const ChunkStart = Integer(Offset);
  const ChunkByteCount = Integer(ChunkLength);
  const Chunk = Copy(ContentBytes, ChunkStart, ChunkByteCount);

  const RangeEnd = Offset + ChunkLength - 1;
  const ContentRange = BuildContentRange(Offset, RangeEnd, TotalSize);

  Result := FGraphClient.PutAbsoluteUrlBytes(UploadUrl, Chunk, ContentRange);
end;

function TAttachmentUploader.BuildContentRange(const RangeStart, RangeEnd, TotalSize: Int64): string;
begin
  Result := Format('bytes %s-%s/%s',
                   [IntToStr(RangeStart), IntToStr(RangeEnd), IntToStr(TotalSize)]);
end;

function TAttachmentUploader.AcceptedOffset(const UploadUrl: string; const FileName: string;
  const Response: TGraphHttpResponse; const CurrentOffset, ExpectedOffset, TotalSize: Int64): Int64;
begin
  Result := ReportedNextOffset(Response.Content, ExpectedOffset);

  const MovesForward = (Result > CurrentOffset);
  if not MovesForward then
    FailStalledUpload(UploadUrl, FileName, Result, TotalSize);
end;

function TAttachmentUploader.ReportedNextOffset(const ResponseContent: string;
  const FallbackOffset: Int64): Int64;
begin
  Result := FallbackOffset;

  var ResponseObj := TJSONObject.ParseJSONValue(ResponseContent) as TJSONObject;
  if not Assigned(ResponseObj) then
    Exit;

  try
    const Ranges = TGraphJson.GetArray(ResponseObj, 'nextExpectedRanges');
    const FirstRange = TGraphJson.ArrayString(Ranges, 0);
    if FirstRange.IsEmpty then
      Exit;

    const RangeParts = FirstRange.Split(['-']);

    Result := StrToInt64Def(RangeParts[0], FallbackOffset);
  finally
    ResponseObj.Free;
  end;
end;

function TAttachmentUploader.AttachmentIdOrFail(const Response: TGraphHttpResponse;
  const FileName: string; const TotalSize: Int64): string;
begin
  if not Response.IsSuccess then
    raise EGraphApiException.CreateFmt(
      'Could not add attachment "%s" (%d bytes) to the draft: HTTP %d - %s.',
      [FileName, TotalSize, Response.StatusCode, FailureReason(Response)]);

  var AttachmentObj := TJSONObject.ParseJSONValue(Response.Content) as TJSONObject;
  if not Assigned(AttachmentObj) then
  begin
    Result := '';
    Exit;
  end;

  try
    Result := TGraphJson.GetString(AttachmentObj, 'id');
  finally
    AttachmentObj.Free;
  end;
end;

function TAttachmentUploader.UploadUrlOrFail(const Response: TGraphHttpResponse;
  const FileName: string; const TotalSize: Int64): string;
begin
  if not Response.IsSuccess then
    raise EGraphApiException.Create(DescribeSessionFailure(FileName, TotalSize, Response));

  var SessionObj := TJSONObject.ParseJSONValue(Response.Content) as TJSONObject;
  if not Assigned(SessionObj) then
    raise EGraphApiException.CreateFmt(
      'The upload session for "%s" did not return a readable response.', [FileName]);

  try
    Result := TGraphJson.GetString(SessionObj, 'uploadUrl');
  finally
    SessionObj.Free;
  end;

  if Result.IsEmpty then
    raise EGraphApiException.CreateFmt(
      'The upload session for "%s" did not return an upload URL.', [FileName]);
end;

function TAttachmentUploader.AttachmentIdFromLocation(const Location: string): string;
begin
  if Location.IsEmpty then
  begin
    Result := '';
    Exit;
  end;

  const Segments = Location.Split(['/']);
  const HasSegments = (Length(Segments) > 0);
  if not HasSegments then
  begin
    Result := '';
    Exit;
  end;

  var LastSegment := Segments[High(Segments)];

  const QueryStart = LastSegment.IndexOf('?');
  const HasQuery = (QueryStart >= 0);
  if HasQuery then
    LastSegment := LastSegment.Substring(0, QueryStart);

  Result := UnwrapODataKey(LastSegment);
end;

function TAttachmentUploader.UnwrapODataKey(const Segment: string): string;
begin
  const KeyStart = Segment.IndexOf('(');
  const KeyEnd = Segment.LastIndexOf(')');

  const HasODataKey = ((KeyStart >= 0) and (KeyEnd > KeyStart));
  if not HasODataKey then
  begin
    Result := Segment;
    Exit;
  end;

  const Key = Segment.Substring(KeyStart + 1, KeyEnd - KeyStart - 1);

  Result := Key.DeQuotedString('''');
end;

procedure TAttachmentUploader.FailChunkUpload(const UploadUrl: string; const FileName: string;
  const RangeStart, RangeEnd, TotalSize: Int64; const Response: TGraphHttpResponse);
begin
  CancelUploadSession(UploadUrl);

  raise EGraphApiException.CreateFmt(
    'Upload of "%s" (%d bytes) failed at bytes %d-%d: HTTP %d - %s. The draft has not been sent.',
    [FileName, TotalSize, RangeStart, RangeEnd, Response.StatusCode, ChunkFailureReason(Response)]);
end;

procedure TAttachmentUploader.FailInterruptedUpload(const UploadUrl: string; const FileName: string;
  const RangeStart, RangeEnd, TotalSize: Int64; const Error: Exception);
begin
  CancelUploadSession(UploadUrl);

  raise EGraphApiException.CreateFmt(
    'Upload of "%s" (%d bytes) was interrupted at bytes %d-%d: %s. The draft has not been sent.',
    [FileName, TotalSize, RangeStart, RangeEnd, Error.Message]);
end;

procedure TAttachmentUploader.FailStalledUpload(const UploadUrl: string; const FileName: string;
  const ResumeOffset, TotalSize: Int64);
begin
  CancelUploadSession(UploadUrl);

  raise EGraphApiException.CreateFmt(
    'Upload of "%s" (%d bytes) stalled: the server asked to resume at byte %d, which is not past ' +
    'the bytes already sent. The draft has not been sent.',
    [FileName, TotalSize, ResumeOffset]);
end;

procedure TAttachmentUploader.FailUnconfirmedUpload(const UploadUrl: string; const FileName: string;
  const TotalSize: Int64);
begin
  CancelUploadSession(UploadUrl);

  raise EGraphApiException.CreateFmt(
    'Upload of "%s" (%d bytes) sent every byte but was never confirmed by the server. ' +
    'The draft has not been sent.', [FileName, TotalSize]);
end;

procedure TAttachmentUploader.CancelUploadSession(const UploadUrl: string);
begin
  try
    FGraphClient.DeleteAbsoluteUrl(UploadUrl);
  except
    on E: Exception do
      FGraphClient.Log(LogLevelError,
        Format('Could not cancel the upload session: %s', [E.Message]));
  end;
end;

function TAttachmentUploader.IsBelowUploadSessionMinimum(const Response: TGraphHttpResponse): Boolean;
begin
  if Response.IsSuccess then
  begin
    Result := False;
    Exit;
  end;

  const ErrorCode = ErrorMember(Response.Content, 'code');

  Result := SameText(ErrorCode, ErrorCodeAttachmentBelowMinimum);
end;

function TAttachmentUploader.DescribeSessionFailure(const FileName: string; const TotalSize: Int64;
  const Response: TGraphHttpResponse): string;
begin
  Result := Format('Could not start an upload session for "%s" (%d bytes): HTTP %d - %s.',
                   [FileName, TotalSize, Response.StatusCode, SessionFailureReason(Response)]);
end;

function TAttachmentUploader.SessionFailureReason(const Response: TGraphHttpResponse): string;
begin
  Result := FailureReason(Response);

  const MayBeSharedMailboxRestriction = ((Response.StatusCode = HttpForbidden) and
                                         FGraphClient.IsSharedMailbox);
  if MayBeSharedMailboxRestriction then
    Result := Format('%s (%s)', [Result, SharedMailboxRestrictionHint]);
end;

function TAttachmentUploader.ChunkFailureReason(const Response: TGraphHttpResponse): string;
begin
  const IsSessionExpired = (Response.StatusCode = HttpUnauthorized);
  if IsSessionExpired then
  begin
    Result := SessionExpiredReason;
    Exit;
  end;

  Result := FailureReason(Response);
end;

function TAttachmentUploader.FailureReason(const Response: TGraphHttpResponse): string;
begin
  Result := ErrorMember(Response.Content, 'message');

  if Result.IsEmpty then
    Result := Response.Content;
end;

function TAttachmentUploader.ErrorMember(const ResponseContent: string;
  const MemberName: string): string;
begin
  Result := '';

  var ResponseObj := TJSONObject.ParseJSONValue(ResponseContent) as TJSONObject;
  if not Assigned(ResponseObj) then
    Exit;

  try
    const ErrorObj = TGraphJson.GetObject(ResponseObj, 'error');
    if not Assigned(ErrorObj) then
      Exit;

    Result := TGraphJson.GetString(ErrorObj, MemberName);
  finally
    ResponseObj.Free;
  end;
end;

end.
