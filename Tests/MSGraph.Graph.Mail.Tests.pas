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
    function CaptureHeaderError(const Headers: TArray<TMailHeader>): string;
    function PostedMessageBody(const RequestIndex: Integer): TJSONObject;
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
    procedure CreateDraft_HeaderWithoutXPrefix_RaisesReadableException;
    [Test]
    procedure CreateDraft_HeaderWithEmptyName_RaisesReadableException;
    [Test]
    procedure CreateDraft_HeaderValueWithLineBreak_RaisesReadableException;
    [Test]
    procedure CreateDraft_DuplicateHeaderName_RaisesReadableException;
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

class function TMailClientTests.JsonString(const Obj: TJSONObject; const Name: string): string;
begin
  Result := '';
  if not Assigned(Obj) then
    Exit;

  var Value := Obj.Values[Name];
  if Assigned(Value) then
    Result := Value.Value;
end;

function TMailClientTests.CaptureHeaderError(const Headers: TArray<TMailHeader>): string;
begin
  Result := '';
  try
    FMailClient.CreateDraft('Subject', 'Body', [Recipient], [], [], False, Headers);
  except
    on E: EInvalidMailHeaderException do
      Result := E.Message;
  end;
end;

function TMailClientTests.PostedMessageBody(const RequestIndex: Integer): TJSONObject;
begin
  const Posted = FFake.RequestAt(RequestIndex);
  Result := TJSONObject.ParseJSONValue(Posted.Body) as TJSONObject;
  Assert.IsNotNull(Result, 'body is not valid JSON: ' + Posted.Body);
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

  var Draft := FMailClient.CreateDraft('Quarterly report', '<p>Hello</p>',
    ['to@example.com'], [], [], True);

  Assert.AreEqual('AAMkTest123', Draft.Id, 'id must come from the response');
  Assert.AreEqual(2, FFake.RequestCount, 'CreateDraft fetches the signature first, then posts the draft');

  var Posted := FFake.RequestAt(1);
  Assert.AreEqual('POST', Posted.Method);
  Assert.IsTrue(Posted.Url.EndsWith('/me/messages'), 'unexpected url: ' + Posted.Url);

  var Body := TJSONObject.ParseJSONValue(Posted.Body) as TJSONObject;
  Assert.IsNotNull(Body, 'body is not valid JSON: ' + Posted.Body);
  try
    Assert.AreEqual('Quarterly report', JsonString(Body, 'subject'));

    var BodyObj := TGraphJson.GetObject(Body, 'body');
    Assert.IsNotNull(BodyObj, 'body object is missing');
    Assert.AreEqual('HTML', JsonString(BodyObj, 'contentType'));
    Assert.AreEqual('<p>Hello</p>', JsonString(BodyObj, 'content'));

    var Recipients := TGraphJson.GetArray(Body, 'toRecipients');
    Assert.IsNotNull(Recipients, 'toRecipients is missing');
    Assert.AreEqual(1, Recipients.Count);

    var Email := TGraphJson.GetObject(TGraphJson.ArrayItem(Recipients, 0), 'emailAddress');
    Assert.AreEqual('to@example.com', JsonString(Email, 'address'));

    Assert.IsNull(TGraphJson.GetArray(Body, 'ccRecipients'), 'an empty cc must not appear in the body');
    Assert.IsNull(TGraphJson.GetArray(Body, 'bccRecipients'), 'an empty bcc must not appear in the body');
    Assert.IsNull(TGraphJson.GetArray(Body, 'internetMessageHeaders'),
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

  FMailClient.CreateDraft('Subject', 'Body', ['to@example.com'], [], [], False);

  var Posted := FFake.RequestAt(1);
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

  var Posted := FFake.RequestAt(0);
  Assert.AreEqual('POST', Posted.Method);
  Assert.IsTrue(Posted.Url.EndsWith('/me/messages/MSG-1/attachments'), 'unexpected url: ' + Posted.Url);

  var Body := TJSONObject.ParseJSONValue(Posted.Body) as TJSONObject;
  Assert.IsNotNull(Body, 'body is not valid JSON: ' + Posted.Body);
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
  FFake.EnqueueResponse(200, '{}');
  FFake.EnqueueResponse(201, '{"id":"AAMkHeaders"}');

  FMailClient.CreateDraft('Subject', 'Body', [Recipient], [], [], False,
    [TMailHeader.Create('x-example-id', '42'), TMailHeader.Create('x-example-source', 'integration')]);

  var Body := PostedMessageBody(1);
  try
    var Headers := TGraphJson.GetArray(Body, 'internetMessageHeaders');
    Assert.IsNotNull(Headers, 'internetMessageHeaders is missing');
    Assert.AreEqual(2, Headers.Count);

    var First := TGraphJson.ArrayItem(Headers, 0);
    Assert.AreEqual('x-example-id', JsonString(First, 'name'));
    Assert.AreEqual('42', JsonString(First, 'value'));

    var Second := TGraphJson.ArrayItem(Headers, 1);
    Assert.AreEqual('x-example-source', JsonString(Second, 'name'));
    Assert.AreEqual('integration', JsonString(Second, 'value'));
  finally
    Body.Free;
  end;
end;

procedure TMailClientTests.CreateDraft_UppercaseXPrefix_IsAccepted;
begin
  FFake.EnqueueResponse(200, '{}');
  FFake.EnqueueResponse(201, '{"id":"AAMkHeaders"}');

  FMailClient.CreateDraft('Subject', 'Body', [Recipient], [], [], False,
    [TMailHeader.Create('X-Example-Id', '42')]);

  var Body := PostedMessageBody(1);
  try
    var Headers := TGraphJson.GetArray(Body, 'internetMessageHeaders');
    Assert.IsNotNull(Headers, 'the prefix check must be case insensitive');
    Assert.AreEqual('X-Example-Id', JsonString(TGraphJson.ArrayItem(Headers, 0), 'name'),
      'the name must be sent exactly as supplied');
  finally
    Body.Free;
  end;
end;

procedure TMailClientTests.CreateDraft_EmptyHeaderArray_OmitsInternetMessageHeaders;
begin
  FFake.EnqueueResponse(200, '{}');
  FFake.EnqueueResponse(201, '{"id":"AAMkHeaders"}');

  FMailClient.CreateDraft('Subject', 'Body', [Recipient], [], [], False, []);

  var Body := PostedMessageBody(1);
  try
    Assert.IsNull(TGraphJson.GetArray(Body, 'internetMessageHeaders'),
      'an empty header array must not appear in the body');
  finally
    Body.Free;
  end;
end;

procedure TMailClientTests.CreateDraft_HeaderWithoutXPrefix_RaisesReadableException;
begin
  const Message = CaptureHeaderError([TMailHeader.Create('Example-Id', '42')]);

  Assert.IsTrue(Message.Contains('Example-Id'), 'the message must name the header: ' + Message);
  Assert.IsTrue(Message.Contains('"x-"'), 'the message must state the rule: ' + Message);
  Assert.AreEqual(0, FFake.RequestCount, 'an invalid header must be rejected before any request');
end;

procedure TMailClientTests.CreateDraft_HeaderWithEmptyName_RaisesReadableException;
begin
  const Message = CaptureHeaderError([TMailHeader.Create('   ', '42')]);

  Assert.IsTrue(Message.Contains('must not be empty'), 'unexpected message: ' + Message);
  Assert.AreEqual(0, FFake.RequestCount, 'an invalid header must be rejected before any request');
end;

procedure TMailClientTests.CreateDraft_HeaderValueWithLineBreak_RaisesReadableException;
begin
  const Message = CaptureHeaderError([TMailHeader.Create('x-example-id', '42'#13#10'x-injected: yes')]);

  Assert.IsTrue(Message.Contains('x-example-id'), 'the message must name the header: ' + Message);
  Assert.IsTrue(Message.Contains('control characters'), 'unexpected message: ' + Message);
  Assert.AreEqual(0, FFake.RequestCount, 'an invalid header must be rejected before any request');
end;

procedure TMailClientTests.CreateDraft_DuplicateHeaderName_RaisesReadableException;
begin
  const Message = CaptureHeaderError([TMailHeader.Create('x-example-id', '1'),
    TMailHeader.Create('X-Example-Id', '2')]);

  Assert.IsTrue(Message.Contains('more than once'), 'unexpected message: ' + Message);
  Assert.AreEqual(0, FFake.RequestCount, 'an invalid header must be rejected before any request');
end;

procedure TMailClientTests.CreateDraft_MoreThanFiveHeaders_RaisesReadableException;
begin
  var Headers: TArray<TMailHeader>;
  SetLength(Headers, 6);
  for var Index := 0 to High(Headers) do
    Headers[Index] := TMailHeader.Create('x-example-' + Index.ToString, Index.ToString);

  const Message = CaptureHeaderError(Headers);

  Assert.IsTrue(Message.Contains('at most 5'), 'unexpected message: ' + Message);
  Assert.AreEqual(0, FFake.RequestCount, 'too many headers must be rejected before any request');
end;

procedure TMailClientTests.CreateDraft_FiveHeaders_IsAccepted;
begin
  FFake.EnqueueResponse(200, '{}');
  FFake.EnqueueResponse(201, '{"id":"AAMkHeaders"}');

  var Headers: TArray<TMailHeader>;
  SetLength(Headers, 5);
  for var Index := 0 to High(Headers) do
    Headers[Index] := TMailHeader.Create('x-example-' + Index.ToString, Index.ToString);

  FMailClient.CreateDraft('Subject', 'Body', [Recipient], [], [], False, Headers);

  var Body := PostedMessageBody(1);
  try
    var Posted := TGraphJson.GetArray(Body, 'internetMessageHeaders');
    Assert.IsNotNull(Posted, 'five headers is the documented maximum and must be accepted');
    Assert.AreEqual(5, Posted.Count);
  finally
    Body.Free;
  end;
end;

procedure TMailClientTests.UpdateDraft_OmitsInternetMessageHeaders;
begin
  FFake.EnqueueResponse(200, '{"id":"MSG-1"}');

  FMailClient.UpdateDraft('MSG-1', 'Subject', 'Body', [Recipient], [], [], False);

  Assert.AreEqual('PATCH', FFake.LastRequest.Method);

  var Body := PostedMessageBody(0);
  try
    Assert.IsNull(TGraphJson.GetArray(Body, 'internetMessageHeaders'),
      'Graph accepts internetMessageHeaders only when creating a message, not on a patch');
  finally
    Body.Free;
  end;
end;

initialization
  TDUnitX.RegisterTestFixture(TMailClientTests);

end.
