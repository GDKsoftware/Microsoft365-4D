unit MSGraph.Graph.Mail.Tests;

interface

uses
  System.JSON,
  DUnitX.TestFramework,
  MSGraph.Graph.Http,
  MSGraph.Graph.Http.Transport.Fake,
  MSGraph.Graph.Mail.Interfaces,
  MSGraph.Graph.Mail.Types;

type
  [TestFixture]
  TMailClientTests = class
  strict private
    FFake: IFakeGraphHttpTransport;
    FGraphClient: TGraphHttpClient;
    FMailClient: IMailClient;

    class function JsonString(const Obj: TJSONObject; const Name: string): string; static;
    procedure EnqueueCreateDraftResponses;
    class function NumberedHeaders(const Count: Integer): TArray<TMailHeader>; static;
    procedure CreateDraftWithHeaders(const Headers: TArray<TMailHeader>);
    procedure AssertHeaderRejected(const Headers: TArray<TMailHeader>;
      const ExpectedMessage: string);
    function RequestJsonAt(const RequestIndex: Integer): TJSONObject;
    function PostedDraftJson: TJSONObject;
  public
    [Setup]
    procedure Setup;
    [TearDown]
    procedure TearDown;

    [Test]
    procedure CreateDraft_ComposesMessageBody;
    [Test]
    procedure CreateDraft_SharedMailbox_UsesDraftsFolderAndAnchorHeader;
    [Test]
    procedure AddAttachment_SendsBase64FileAttachment;
    [Test]
    procedure UnstubbedRequest_RaisesException;

    [Test]
    procedure CreateDraft_WithCustomHeaders_AddsInternetMessageHeaders;
    [Test]
    procedure CreateDraft_UppercaseXPrefix_IsAccepted;
    [Test]
    procedure CreateDraft_EmptyHeaderArray_OmitsInternetMessageHeaders;
    [Test]
    [TestCase('MissingXPrefix', 'Example-Id;42;Custom mail header "Example-Id" is not allowed: a custom header name must start with "x-".', ';')]
    [TestCase('EmptyName', '   ;42;A custom mail header name must not be empty.', ';')]
    [TestCase('NameWithColon', 'x-example-id:42;value;Custom mail header "x-example-id:42" must contain printable characters only, and no ":".', ';')]
    [TestCase('NameWithTrailingSpace', 'x-example-id ;42;Custom mail header "x-example-id " must contain printable characters only, and no ":".', ';')]
    [TestCase('ValueWithLineBreak', 'x-example-id;42' + #13#10 + 'x-injected: yes;The value of custom mail header "x-example-id" must not contain line breaks or control characters.', ';')]
    [TestCase('ValueWithDeleteCharacter', 'x-example-id;42' + #127 + ';The value of custom mail header "x-example-id" must not contain line breaks or control characters.', ';')]
    [TestCase('ValueWithUnicodeSeparator', 'x-example-id;42' + #$2028 + 'more;The value of custom mail header "x-example-id" must not contain line breaks or control characters.', ';')]
    procedure CreateDraft_InvalidHeader_RaisesReadableException(const Name, Value, ExpectedMessage: string);
    [Test]
    procedure CreateDraft_DuplicateHeaderName_RaisesReadableException;
    [Test]
    procedure CreateDraft_DuplicateAfterOtherHeaders_RaisesReadableException;
    [Test]
    procedure CreateDraft_MoreThanFiveHeaders_RaisesReadableException;
    [Test]
    procedure CreateDraft_FiveHeaders_IsAccepted;
    [Test]
    procedure UpdateDraft_OmitsInternetMessageHeaders;
  end;

implementation

uses
  System.SysUtils,
  MSGraph.Graph.JsonHelper,
  MSGraph.Graph.Mail;

const
  DummyAccessToken = 'unit-test-token';
  SharedMailbox = 'shared@example.com';
  Recipient = 'to@example.com';
  DraftSubject = 'Subject';
  DraftBody = 'Body';
  SignatureResponse = '{}';
  DraftCreatedResponse = '{"id":"AAMkHeaders"}';
  DraftRequestIndex = 1;
  DraftRequestCount = 2;
  RejectedBeforeRequest = 'an invalid header must be rejected before any request';
  InternetMessageHeadersKey = 'internetMessageHeaders';

class function TMailClientTests.JsonString(const Obj: TJSONObject; const Name: string): string;
begin
  Result := '';
  if not Assigned(Obj) then
    Exit;

  const Value = Obj.Values[Name];
  if Assigned(Value) then
    Result := Value.Value;
end;

procedure TMailClientTests.EnqueueCreateDraftResponses;
begin
  FFake.EnqueueResponse(200, SignatureResponse);
  FFake.EnqueueResponse(201, DraftCreatedResponse);
end;

class function TMailClientTests.NumberedHeaders(const Count: Integer): TArray<TMailHeader>;
begin
  SetLength(Result, Count);
  for var Index := 0 to High(Result) do
  begin
    Result[Index] := TMailHeader.Create('x-example-' + Index.ToString, Index.ToString);
  end;
end;

procedure TMailClientTests.CreateDraftWithHeaders(const Headers: TArray<TMailHeader>);
begin
  FMailClient.CreateDraft(DraftSubject, DraftBody, [Recipient], [], [], False, Headers);
end;

procedure TMailClientTests.AssertHeaderRejected(const Headers: TArray<TMailHeader>;
  const ExpectedMessage: string);
begin
  Assert.WillRaiseWithMessage(
    procedure
    begin
      CreateDraftWithHeaders(Headers);
    end,
    EInvalidMailHeaderException,
    ExpectedMessage);

  Assert.AreEqual(0, FFake.RequestCount, RejectedBeforeRequest);
end;

function TMailClientTests.RequestJsonAt(const RequestIndex: Integer): TJSONObject;
begin
  const Posted = FFake.RequestAt(RequestIndex);
  const Parsed = TJSONObject.ParseJSONValue(Posted.Body);
  if Parsed is TJSONObject then
    Result := TJSONObject(Parsed)
  else
  begin
    Parsed.Free;
    Result := nil;
  end;

  Assert.IsNotNull(Result, 'body is not a JSON object: ' + Posted.Body);
end;

function TMailClientTests.PostedDraftJson: TJSONObject;
begin
  Assert.AreEqual(DraftRequestCount, FFake.RequestCount,
    'CreateDraft fetches the signature first, then posts the draft');
  Result := RequestJsonAt(DraftRequestIndex);
end;

procedure TMailClientTests.Setup;
begin
  FFake := TFakeGraphHttpTransport.Create;
  FGraphClient := TGraphHttpClient.Create(DummyAccessToken, FFake);
  FMailClient := TMailClient.Create(FGraphClient, True);
end;

procedure TMailClientTests.TearDown;
begin
  FMailClient := nil;
  FGraphClient := nil;
  FFake := nil;
end;

procedure TMailClientTests.CreateDraft_ComposesMessageBody;
begin
  FFake.EnqueueResponse(200, '{}');
  FFake.EnqueueResponse(201, '{"id":"AAMkTest123"}');

  const Draft = FMailClient.CreateDraft('Quarterly report', '<p>Hello</p>',
    [Recipient], [], [], True);

  Assert.AreEqual('AAMkTest123', Draft.Id, 'id must come from the response');

  const Posted = FFake.RequestAt(DraftRequestIndex);
  Assert.AreEqual('POST', Posted.Method);
  Assert.IsTrue(Posted.Url.EndsWith('/me/messages'), 'unexpected url: ' + Posted.Url);

  const Body = PostedDraftJson;
  try
    Assert.AreEqual('Quarterly report', JsonString(Body, 'subject'));

    const BodyObj = TGraphJson.GetObject(Body, 'body');
    Assert.IsNotNull(BodyObj, 'body object is missing');
    Assert.AreEqual('HTML', JsonString(BodyObj, 'contentType'));
    Assert.AreEqual('<p>Hello</p>', JsonString(BodyObj, 'content'));

    const Recipients = TGraphJson.GetArray(Body, 'toRecipients');
    Assert.IsNotNull(Recipients, 'toRecipients is missing');
    Assert.AreEqual(1, Recipients.Count);

    const Email = TGraphJson.GetObject(TGraphJson.ArrayItem(Recipients, 0), 'emailAddress');
    Assert.AreEqual(Recipient, JsonString(Email, 'address'));

    Assert.IsNull(TGraphJson.GetArray(Body, 'ccRecipients'), 'an empty cc must not appear in the body');
    Assert.IsNull(TGraphJson.GetArray(Body, 'bccRecipients'), 'an empty bcc must not appear in the body');
    Assert.IsNull(TGraphJson.GetArray(Body, InternetMessageHeadersKey),
      'a call without custom headers must not add internetMessageHeaders');
  finally
    Body.Free;
  end;
end;

procedure TMailClientTests.CreateDraft_SharedMailbox_UsesDraftsFolderAndAnchorHeader;
begin
  FGraphClient.MailboxAddress := SharedMailbox;

  FFake.EnqueueResponse(200, '{}');
  FFake.EnqueueResponse(201, '{"id":"AAMkShared"}');

  FMailClient.CreateDraft(DraftSubject, DraftBody, [Recipient], [], [], False);

  const Posted = FFake.RequestAt(DraftRequestIndex);
  Assert.IsTrue(Posted.Url.EndsWith('/users/' + SharedMailbox + '/mailFolders/Drafts/messages'),
    'unexpected url: ' + Posted.Url);
  Assert.AreEqual(SharedMailbox, FFake.HeaderValue(1, 'X-AnchorMailbox'),
    'a shared mailbox must send an anchor header');
end;

procedure TMailClientTests.AddAttachment_SendsBase64FileAttachment;
begin
  FFake.EnqueueResponse(201, '{"id":"attachment-1"}');

  Assert.IsTrue(FMailClient.AddAttachment('MSG-1', 'note.txt', 'text/plain',
    TEncoding.UTF8.GetBytes('Hello')));

  Assert.AreEqual(1, FFake.RequestCount);

  const Posted = FFake.RequestAt(0);
  Assert.AreEqual('POST', Posted.Method);
  Assert.IsTrue(Posted.Url.EndsWith('/me/messages/MSG-1/attachments'), 'unexpected url: ' + Posted.Url);

  const Body = RequestJsonAt(0);
  try
    Assert.AreEqual('#microsoft.graph.fileAttachment', JsonString(Body, '@odata.type'));
    Assert.AreEqual('note.txt', JsonString(Body, 'name'));
    Assert.AreEqual('text/plain', JsonString(Body, 'contentType'));
    Assert.AreEqual('SGVsbG8=', JsonString(Body, 'contentBytes'));
  finally
    Body.Free;
  end;
end;

procedure TMailClientTests.UnstubbedRequest_RaisesException;
begin
  Assert.WillRaise(
    procedure
    begin
      FMailClient.SendDraft('MSG-1');
    end,
    Exception,
    'a request without a stubbed response must not pass silently');
end;

procedure TMailClientTests.CreateDraft_WithCustomHeaders_AddsInternetMessageHeaders;
begin
  EnqueueCreateDraftResponses;

  CreateDraftWithHeaders([TMailHeader.Create('x-example-id', '42'),
    TMailHeader.Create('x-example-source', 'integration')]);

  const Body = PostedDraftJson;
  try
    const Headers = TGraphJson.GetArray(Body, InternetMessageHeadersKey);
    Assert.IsNotNull(Headers, 'internetMessageHeaders is missing');
    Assert.AreEqual(2, Headers.Count);

    const First = TGraphJson.ArrayItem(Headers, 0);
    Assert.IsTrue(JsonString(First, 'name') = 'x-example-id', 'unexpected first header name');
    Assert.IsTrue(JsonString(First, 'value') = '42', 'unexpected first header value');

    const Second = TGraphJson.ArrayItem(Headers, 1);
    Assert.IsTrue(JsonString(Second, 'name') = 'x-example-source', 'unexpected second header name');
    Assert.IsTrue(JsonString(Second, 'value') = 'integration', 'unexpected second header value');
  finally
    Body.Free;
  end;
end;

procedure TMailClientTests.CreateDraft_UppercaseXPrefix_IsAccepted;
begin
  EnqueueCreateDraftResponses;

  CreateDraftWithHeaders([TMailHeader.Create('X-Example-Id', '42')]);

  const Body = PostedDraftJson;
  try
    const Headers = TGraphJson.GetArray(Body, InternetMessageHeadersKey);
    Assert.IsNotNull(Headers, 'the prefix check must be case insensitive');
    Assert.AreEqual(1, Headers.Count);

    const PostedName = JsonString(TGraphJson.ArrayItem(Headers, 0), 'name');
    Assert.IsTrue(PostedName = 'X-Example-Id',
      'the name must be sent exactly as supplied, but was: ' + PostedName);
  finally
    Body.Free;
  end;
end;

procedure TMailClientTests.CreateDraft_EmptyHeaderArray_OmitsInternetMessageHeaders;
begin
  EnqueueCreateDraftResponses;

  CreateDraftWithHeaders([]);

  const Body = PostedDraftJson;
  try
    Assert.IsNull(TGraphJson.GetArray(Body, InternetMessageHeadersKey),
      'an empty header array must not appear in the body');
  finally
    Body.Free;
  end;
end;

procedure TMailClientTests.CreateDraft_InvalidHeader_RaisesReadableException(const Name, Value,
  ExpectedMessage: string);
begin
  AssertHeaderRejected([TMailHeader.Create(Name, Value)], ExpectedMessage);
end;

procedure TMailClientTests.CreateDraft_DuplicateHeaderName_RaisesReadableException;
begin
  AssertHeaderRejected([TMailHeader.Create('x-example-id', '1'), TMailHeader.Create('X-Example-Id', '2')],
    'Custom mail header "x-example-id" is specified more than once.');
end;

procedure TMailClientTests.CreateDraft_DuplicateAfterOtherHeaders_RaisesReadableException;
begin
  AssertHeaderRejected([TMailHeader.Create('x-example-a', '1'), TMailHeader.Create('x-example-b', '2'),
    TMailHeader.Create('x-example-b', '3')],
    'Custom mail header "x-example-b" is specified more than once.');
end;

procedure TMailClientTests.CreateDraft_MoreThanFiveHeaders_RaisesReadableException;
begin
  AssertHeaderRejected(NumberedHeaders(6),
    'A message accepts at most 5 custom mail headers, but 6 were supplied.');
end;

procedure TMailClientTests.CreateDraft_FiveHeaders_IsAccepted;
begin
  EnqueueCreateDraftResponses;

  CreateDraftWithHeaders(NumberedHeaders(5));

  const Body = PostedDraftJson;
  try
    const Headers = TGraphJson.GetArray(Body, InternetMessageHeadersKey);
    Assert.IsNotNull(Headers, 'five headers is the documented maximum and must be accepted');
    Assert.AreEqual(5, Headers.Count);
  finally
    Body.Free;
  end;
end;

procedure TMailClientTests.UpdateDraft_OmitsInternetMessageHeaders;
begin
  FFake.EnqueueResponse(200, '{"id":"MSG-1"}');

  FMailClient.UpdateDraft('MSG-1', DraftSubject, DraftBody, [Recipient], [], [], False);

  Assert.AreEqual('PATCH', FFake.LastRequest.Method);

  const Body = RequestJsonAt(0);
  try
    Assert.IsNull(TGraphJson.GetArray(Body, InternetMessageHeadersKey),
      'Graph accepts internetMessageHeaders only when creating a message, not on a patch');
  finally
    Body.Free;
  end;
end;

initialization
  TDUnitX.RegisterTestFixture(TMailClientTests);

end.
