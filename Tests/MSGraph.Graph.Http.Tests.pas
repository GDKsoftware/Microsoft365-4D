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

    [Test]
    procedure PostRaw_ReturnsStatusCodeAndUnparsedBody;
    [Test]
    procedure PutAbsoluteUrlBytes_SendsBytesWithContentRangeAndWithoutAuthorization;
    [Test]
    procedure DeleteAbsoluteUrl_SendsWithoutAuthorization;
    [Test]
    procedure ResponseHeaderValue_IgnoresCaseAndReturnsEmptyWhenAbsent;
  end;

implementation

uses
  System.SysUtils,
  System.JSON,
  System.Net.URLClient,
  MSGraph.OAuth2.Types,

  MSGraph.Graph.Http.Types,
  MSGraph.Graph.Http.Interfaces;

const
  DummyAccessToken = 'unit-test-token';
  ExpectedAuthorization = 'Bearer ' + DummyAccessToken;
  SharedMailbox = 'shared@example.com';
  UploadUrl = 'https://outlook.office.com/api/v2.0/AttachmentSessions?authtoken=SECRET';
  UploadContentRange = 'bytes 0-2/3';
  HeaderContentType = 'Content-Type';
  HeaderContentRange = 'Content-Range';
  HeaderAuthorization = 'Authorization';
  HeaderAnchorMailbox = 'X-AnchorMailbox';

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

procedure TGraphHttpClientTests.PostRaw_ReturnsStatusCodeAndUnparsedBody;
begin
  const Payload = '{"error":{"code":"ErrorCode","message":"Something went wrong"}}';
  FFake.EnqueueResponse(413, Payload);

  const Response = FClient.PostRaw('/me/messages/MSG-1/attachments', '{}');

  Assert.AreEqual(413, Response.StatusCode, 'the raw variant keeps the status code');
  Assert.AreEqual(Payload, Response.Content, 'the raw variant keeps the untouched body');
  Assert.IsFalse(Response.IsSuccess);
end;

procedure TGraphHttpClientTests.PutAbsoluteUrlBytes_SendsBytesWithContentRangeAndWithoutAuthorization;
begin
  FClient.MailboxAddress := SharedMailbox;
  FFake.EnqueueResponse(200, '{}');

  const Payload: TBytes = [1, 2, 3];
  FClient.PutAbsoluteUrlBytes(UploadUrl, Payload, UploadContentRange);

  Assert.AreEqual('PUT', FFake.LastRequest.Method);
  Assert.AreEqual(UploadUrl, FFake.LastRequest.Url, 'the upload url is used verbatim');
  Assert.AreEqual(3, Length(FFake.LastRequest.BodyBytes));
  Assert.AreEqual(2, Length(FFake.LastRequest.Headers),
    'an upload sends Content-Type and Content-Range only');
  Assert.AreEqual('application/octet-stream', FFake.HeaderValue(0, HeaderContentType));
  Assert.AreEqual(UploadContentRange, FFake.HeaderValue(0, HeaderContentRange));
  Assert.AreEqual('', FFake.HeaderValue(0, HeaderAuthorization),
    'the upload url is pre-authenticated');
  Assert.AreEqual('', FFake.HeaderValue(0, HeaderAnchorMailbox),
    'the upload url is not a Graph endpoint');
end;

procedure TGraphHttpClientTests.DeleteAbsoluteUrl_SendsWithoutAuthorization;
begin
  FFake.EnqueueResponse(204, '');

  FClient.DeleteAbsoluteUrl(UploadUrl);

  Assert.AreEqual('DELETE', FFake.LastRequest.Method);
  Assert.AreEqual(UploadUrl, FFake.LastRequest.Url);
  Assert.AreEqual(0, Length(FFake.LastRequest.Headers),
    'cancelling an upload session needs no headers at all');
end;

procedure TGraphHttpClientTests.ResponseHeaderValue_IgnoresCaseAndReturnsEmptyWhenAbsent;
begin
  var Response := Default(TGraphHttpResponse);
  Response.Headers := [TNetHeader.Create('Location', 'https://example.com/attachment')];

  Assert.AreEqual('https://example.com/attachment', Response.HeaderValue('location'));
  Assert.AreEqual('', Response.HeaderValue(HeaderContentRange));
end;

initialization
  TDUnitX.RegisterTestFixture(TGraphHttpClientTests);

end.
