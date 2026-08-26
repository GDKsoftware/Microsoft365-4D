unit MSGraph.Graph.Http.Tests;

interface

uses
  DUnitX.TestFramework,
  MSGraph.Graph.Http,
  MSGraph.Graph.Http.Transport.Fake;

type
  [TestFixture]
  TGraphHttpClientTests = class
  strict private
    FFake: IFakeGraphHttpTransport;
    FClient: TGraphHttpClient;

    procedure ExecuteAndDiscard;
  public
    [Setup]
    procedure Setup;
    [TearDown]
    procedure TearDown;

    [Test]
    procedure Get_SendsAuthorizationAndContentType;
    [Test]
    procedure Get_SharedMailbox_AddsAnchorHeader;
    [Test]
    procedure Delete_SendsOnlyAuthorizationHeader;
    [Test]
    procedure Patch_SendsBodyUnchanged;
    [Test]
    procedure Post_WithoutBody_SendsEmptyBody;
    [Test]
    procedure GetWithHeaders_AddsExtraHeaderAndClearsItAfterwards;
    [Test]
    procedure GetRawBytes_ReturnsResponseBytes;
    [Test]
    procedure GetRawBytes_NonSuccess_RaisesGraphApiException;
    [Test]
    procedure EmptyAccessToken_RaisesBeforeReachingTransport;
    [Test]
    procedure NoTransport_RaisesGraphApiException;
  end;

implementation

uses
  System.SysUtils,
  System.JSON,
  MSGraph.OAuth2.Types,

  MSGraph.Graph.Http.Interfaces;

const
  DummyAccessToken = 'unit-test-token';
  ExpectedAuthorization = 'Bearer ' + DummyAccessToken;
  SharedMailbox = 'shared@example.com';

procedure TGraphHttpClientTests.Setup;
begin
  FFake := TFakeGraphHttpTransport.Create;
  FClient := TGraphHttpClient.Create(DummyAccessToken, FFake);
end;

procedure TGraphHttpClientTests.TearDown;
begin
  FClient.Free;
  FClient := nil;
  FFake := nil;
end;

procedure TGraphHttpClientTests.ExecuteAndDiscard;
begin
  var Response := FClient.Get('/me');
  Response.Free;
end;

procedure TGraphHttpClientTests.Get_SendsAuthorizationAndContentType;
begin
  FFake.EnqueueResponse(200, '{}');

  ExecuteAndDiscard;

  Assert.AreEqual(1, FFake.RequestCount);
  Assert.AreEqual('GET', FFake.LastRequest.Method);
  Assert.AreEqual('https://graph.microsoft.com/v1.0/me', FFake.LastRequest.Url);
  Assert.AreEqual(2, Length(FFake.LastRequest.Headers), 'GET sends Authorization and Content-Type');
  Assert.AreEqual(ExpectedAuthorization, FFake.HeaderValue(0, 'Authorization'));
  Assert.AreEqual('application/json', FFake.HeaderValue(0, 'Content-Type'));
end;

procedure TGraphHttpClientTests.Get_SharedMailbox_AddsAnchorHeader;
begin
  FClient.MailboxAddress := SharedMailbox;
  FFake.EnqueueResponse(200, '{}');

  ExecuteAndDiscard;

  Assert.AreEqual(3, Length(FFake.LastRequest.Headers), 'a shared mailbox adds a third header');
  Assert.AreEqual(SharedMailbox, FFake.HeaderValue(0, 'X-AnchorMailbox'));
end;

procedure TGraphHttpClientTests.Delete_SendsOnlyAuthorizationHeader;
begin
  FClient.MailboxAddress := SharedMailbox;
  FFake.EnqueueResponse(204, '');

  var Response := FClient.Delete('/me/messages/MSG-1');
  Response.Free;

  Assert.AreEqual('DELETE', FFake.LastRequest.Method);
  Assert.AreEqual(1, Length(FFake.LastRequest.Headers),
    'DELETE has always sent Authorization only, even for a shared mailbox');
  Assert.AreEqual(ExpectedAuthorization, FFake.HeaderValue(0, 'Authorization'));
  Assert.AreEqual('', FFake.HeaderValue(0, 'Content-Type'));
  Assert.AreEqual('', FFake.HeaderValue(0, 'X-AnchorMailbox'));
end;

procedure TGraphHttpClientTests.Patch_SendsBodyUnchanged;
begin
  FFake.EnqueueResponse(200, '{}');

  const Payload = '{"isRead":true}';
  var Response := FClient.Patch('/me/messages/MSG-1', Payload);
  Response.Free;

  Assert.AreEqual('PATCH', FFake.LastRequest.Method);
  Assert.AreEqual(Payload, FFake.LastRequest.Body);
  Assert.AreEqual('application/json', FFake.HeaderValue(0, 'Content-Type'));
end;

procedure TGraphHttpClientTests.Post_WithoutBody_SendsEmptyBody;
begin
  FFake.EnqueueResponse(202, '');

  var Response := FClient.Post('/me/messages/MSG-1/send');
  Response.Free;

  Assert.AreEqual('POST', FFake.LastRequest.Method);
  Assert.AreEqual('', FFake.LastRequest.Body, 'a POST without a body sends no content');
end;

procedure TGraphHttpClientTests.GetWithHeaders_AddsExtraHeaderAndClearsItAfterwards;
begin
  FFake.EnqueueResponse(200, '{}');
  FFake.EnqueueResponse(200, '{}');

  var Response := FClient.GetWithHeaders('/me/messages', '$top=5',
    ['Prefer: outlook.body-content-type="text"']);
  Response.Free;

  Assert.AreEqual('https://graph.microsoft.com/v1.0/me/messages?$top=5', FFake.RequestAt(0).Url);
  Assert.AreEqual('outlook.body-content-type="text"', FFake.HeaderValue(0, 'Prefer'));

  ExecuteAndDiscard;

  Assert.AreEqual('', FFake.HeaderValue(1, 'Prefer'),
    'the extra header must not leak into the next request');
end;

procedure TGraphHttpClientTests.GetRawBytes_ReturnsResponseBytes;
begin
  const Payload = 'raw mime content';
  FFake.EnqueueBinaryResponse(200, TEncoding.UTF8.GetBytes(Payload));

  const Received = FClient.GetRawBytes('/me/messages/MSG-1/$value');

  Assert.AreEqual(Payload, TEncoding.UTF8.GetString(Received));
  Assert.AreEqual(1, Length(FFake.LastRequest.Headers),
    'the raw variant sends Authorization only');
  Assert.AreEqual(ExpectedAuthorization, FFake.HeaderValue(0, 'Authorization'));
end;

procedure TGraphHttpClientTests.GetRawBytes_NonSuccess_RaisesGraphApiException;
begin
  FFake.EnqueueBinaryResponse(404, nil);

  Assert.WillRaise(
    procedure
    begin
      FClient.GetRawBytes('/me/messages/MSG-1/$value');
    end,
    EGraphApiException);
end;

procedure TGraphHttpClientTests.EmptyAccessToken_RaisesBeforeReachingTransport;
begin
  FClient.SetAccessToken('');

  Assert.WillRaise(
    procedure
    begin
      ExecuteAndDiscard;
    end,
    EGraphApiException);

  Assert.AreEqual(0, FFake.RequestCount,
    'without an access token no request may reach the transport');
end;

procedure TGraphHttpClientTests.NoTransport_RaisesGraphApiException;
begin
  Assert.WillRaise(
    procedure
    begin
      TGraphHttpClient.Create(DummyAccessToken, IGraphHttpTransport(nil)).Free;
    end,
    EGraphApiException);
end;

initialization
  TDUnitX.RegisterTestFixture(TGraphHttpClientTests);

end.
