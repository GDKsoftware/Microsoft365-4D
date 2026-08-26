unit MSGraph.Graph.Mail.Tests;

interface

uses
  System.JSON,
  DUnitX.TestFramework,
  MSGraph.Graph.Http,
  MSGraph.Graph.Http.Transport.Fake,
  MSGraph.Graph.Mail.Interfaces;

type
  [TestFixture]
  TMailClientTests = class
  strict private
    FFake: IFakeGraphHttpTransport;
    FGraphClient: TGraphHttpClient;
    FMailClient: IMailClient;

    class function JsonString(const Obj: TJSONObject; const Name: string): string; static;
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
  end;

implementation

uses
  System.SysUtils,
  MSGraph.Graph.JsonHelper,
  MSGraph.Graph.Mail,
  MSGraph.Graph.Mail.Types;

const
  DummyAccessToken = 'unit-test-token';
  SharedMailbox = 'shared@example.com';

class function TMailClientTests.JsonString(const Obj: TJSONObject; const Name: string): string;
begin
  Result := '';
  if not Assigned(Obj) then
    Exit;

  var Value := Obj.Values[Name];
  if Assigned(Value) then
    Result := Value.Value;
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

initialization
  TDUnitX.RegisterTestFixture(TMailClientTests);

end.
