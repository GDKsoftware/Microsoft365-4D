unit MSGraph.Graph.Http;

interface

uses
  System.SysUtils,
  System.JSON,
  System.Net.URLClient,
  MSGraph.OAuth2.Types,
  MSGraph.Graph.Http.Types,
  MSGraph.Graph.Http.Interfaces;

type
  TUserProfile = record
    Mail: string;
    UserPrincipalName: string;
    DisplayName: string;
  end;

  TGraphHttpClient = class
  strict private
    FAccessToken: string;
    FLogProc: TLogProc;
    FExtraHeaders: TArray<TNetHeader>;
    FMailboxAddress: string;
    FTransport: IGraphHttpTransport;

    function BuildUrl(const Endpoint: string; const QueryParams: string = ''): string;
    function BuildHeaders: TArray<TNetHeader>;
    function BuildAuthorizationHeader: TArray<TNetHeader>;
    function BuildUploadHeaders(const ContentRange: string): TArray<TNetHeader>;
    function ExecuteRequest(const Method: string; const Url: string; const Body: string = ''): TJSONObject;
    function ParseResponse(const StatusCode: Integer; const ResponseText: string): TJSONObject;
    function ParseErrorResponse(const StatusCode: Integer; const ResponseText: string): TJSONObject;
    procedure ValidateAccessToken;
    procedure LogRequest(const Method: string; const Url: string);

    const
      GraphBaseUrl = 'https://graph.microsoft.com/v1.0';
      HeaderAuthorization = 'Authorization';
      HeaderContentType = 'Content-Type';
      HeaderContentRange = 'Content-Range';
      BearerPrefix = 'Bearer ';
      ContentTypeJson = 'application/json';
      ContentTypeOctetStream = 'application/octet-stream';
      MethodGet = 'GET';
      MethodPost = 'POST';
      MethodPatch = 'PATCH';
      MethodPut = 'PUT';
      MethodDelete = 'DELETE';
      LogDebug = 'DEBUG';
      LogError = 'ERROR';
      LogFormatGraphHttpError = 'Graph API HTTP %d - %s';
  public
    constructor Create(const AccessToken: string; const LogProc: TLogProc = nil); overload;
    constructor Create(const AccessToken: string; const Transport: IGraphHttpTransport;
      const LogProc: TLogProc = nil); overload;

    function Get(const Endpoint: string; const QueryParams: string = ''): TJSONObject;
    function GetRawBytes(const Endpoint: string): TBytes;
    function Post(const Endpoint: string; const Body: string = ''): TJSONObject;
    function PostRaw(const Endpoint: string; const Body: string): TGraphHttpResponse;
    function Patch(const Endpoint: string; const Body: string): TJSONObject;
    function Delete(const Endpoint: string): TJSONObject;

    function GetWithHeaders(const Endpoint: string; const QueryParams: string;
      const ExtraHeaders: TArray<string>): TJSONObject;
    function GetAbsoluteUrl(const FullUrl: string): TJSONObject;
    function GetAbsoluteUrlWithHeaders(const FullUrl: string;
      const ExtraHeaders: TArray<string>): TJSONObject;

    function PutAbsoluteUrlBytes(const FullUrl: string; const Bytes: TBytes;
      const ContentRange: string): TGraphHttpResponse;
    function DeleteAbsoluteUrl(const FullUrl: string): TGraphHttpResponse;

    function GetUserPrefix: string;
    function IsSharedMailbox: Boolean;

    procedure SetAccessToken(const Value: string);
    function GetAccessToken: string;

    function GetUserProfile: TUserProfile;

    procedure Log(const Level: string; const Message: string);

    property MailboxAddress: string read FMailboxAddress write FMailboxAddress;
  end;

implementation

uses
  System.NetEncoding,
  MSGraph.Graph.JsonHelper,
  MSGraph.Graph.Http.Redaction,
  MSGraph.Graph.Http.Transport;

constructor TGraphHttpClient.Create(const AccessToken: string; const LogProc: TLogProc);
begin
  inherited Create;
  FAccessToken := AccessToken;
  FLogProc := LogProc;
  FTransport := TNetHttpTransport.Create;
end;

constructor TGraphHttpClient.Create(const AccessToken: string; const Transport: IGraphHttpTransport;
  const LogProc: TLogProc);
begin
  inherited Create;

  if not Assigned(Transport) then
    raise EGraphApiException.Create('No HTTP transport provided.');

  FAccessToken := AccessToken;
  FLogProc := LogProc;
  FTransport := Transport;
end;

procedure TGraphHttpClient.Log(const Level: string; const Message: string);
begin
  if Assigned(FLogProc) then
    FLogProc(Level, Message);
end;

procedure TGraphHttpClient.LogRequest(const Method: string; const Url: string);
begin
  Log(LogDebug, Format('%s %s', [Method, TGraphUrlRedactor.Redact(Url)]));
end;

function TGraphHttpClient.BuildUrl(const Endpoint: string; const QueryParams: string): string;
begin
  Result := GraphBaseUrl + Endpoint;
  if not QueryParams.IsEmpty then
    Result := Result + '?' + QueryParams;
end;

function TGraphHttpClient.BuildHeaders: TArray<TNetHeader>;
begin
  var BaseCount := 2;
  if IsSharedMailbox then
    Inc(BaseCount);
  const ExtraCount = Length(FExtraHeaders);
  SetLength(Result, BaseCount + ExtraCount);
  Result[0] := TNetHeader.Create(HeaderAuthorization, BearerPrefix + FAccessToken);
  Result[1] := TNetHeader.Create(HeaderContentType, ContentTypeJson);
  if IsSharedMailbox then
    Result[2] := TNetHeader.Create('X-AnchorMailbox', FMailboxAddress);

  for var Index := 0 to ExtraCount - 1 do
  begin
    Result[BaseCount + Index] := FExtraHeaders[Index];
  end;
end;

function TGraphHttpClient.BuildAuthorizationHeader: TArray<TNetHeader>;
begin
  SetLength(Result, 1);
  Result[0] := TNetHeader.Create(HeaderAuthorization, BearerPrefix + FAccessToken);
end;

function TGraphHttpClient.BuildUploadHeaders(const ContentRange: string): TArray<TNetHeader>;
begin
  SetLength(Result, 2);
  Result[0] := TNetHeader.Create(HeaderContentType, ContentTypeOctetStream);
  Result[1] := TNetHeader.Create(HeaderContentRange, ContentRange);
end;

procedure TGraphHttpClient.ValidateAccessToken;
begin
  if FAccessToken.Trim.IsEmpty then
    raise EGraphApiException.Create('No access token provided. Please authenticate first.');
end;

function TGraphHttpClient.ParseErrorResponse(const StatusCode: Integer; const ResponseText: string): TJSONObject;
begin
  var DefaultError := Format('HTTP %d: %s', [StatusCode, ResponseText]);

  var ErrorObj := TJSONObject.ParseJSONValue(ResponseText) as TJSONObject;
  if not Assigned(ErrorObj) then
  begin
    Result := TJSONObject.Create;
    Result.AddPair('error', DefaultError);
    Exit;
  end;

  try
    var GraphError := ErrorObj.FindValue('error') as TJSONObject;
    if not Assigned(GraphError) then
    begin
      Result := TJSONObject.Create;
      Result.AddPair('error', DefaultError);
      Exit;
    end;

    var MessageValue := GraphError.FindValue('message');
    var ErrorMessage := '';
    if Assigned(MessageValue) then
      ErrorMessage := MessageValue.Value;

    Result := TJSONObject.Create;
    Result.AddPair('error', Format('Microsoft Graph API error (HTTP %d): %s', [StatusCode, ErrorMessage]));
  finally
    ErrorObj.Free;
  end;
end;

function TGraphHttpClient.ParseResponse(const StatusCode: Integer; const ResponseText: string): TJSONObject;
begin
  const IsSuccess = (StatusCode >= 200) and (StatusCode < 300);

  if not IsSuccess then
  begin
    Log(LogError, Format(LogFormatGraphHttpError, [StatusCode, ResponseText]));
    Result := ParseErrorResponse(StatusCode, ResponseText);
    Exit;
  end;

  if ResponseText.Trim.IsEmpty then
  begin
    Result := TJSONObject.Create;
    Result.AddPair('success', TJSONBool.Create(True));
    Exit;
  end;

  var ParsedValue := TJSONObject.ParseJSONValue(ResponseText);
  if ParsedValue is TJSONObject then
    Result := TJSONObject(ParsedValue)
  else
  begin
    ParsedValue.Free;
    Result := TJSONObject.Create;
    Result.AddPair('raw', ResponseText);
  end;
end;

function TGraphHttpClient.ExecuteRequest(const Method: string; const Url: string; const Body: string): TJSONObject;
begin
  ValidateAccessToken;

  var Request := Default(TGraphHttpRequest);
  Request.Method := Method;
  Request.Url := Url;
  Request.Body := Body;

  if Method = MethodDelete then
    Request.Headers := BuildAuthorizationHeader
  else
    Request.Headers := BuildHeaders;

  var Response := FTransport.Execute(Request);
  Result := ParseResponse(Response.StatusCode, Response.Content);
end;

function TGraphHttpClient.Get(const Endpoint: string; const QueryParams: string): TJSONObject;
begin
  var Url := BuildUrl(Endpoint, QueryParams);
  LogRequest(MethodGet, Url);
  Result := ExecuteRequest(MethodGet, Url);
end;

function TGraphHttpClient.GetRawBytes(const Endpoint: string): TBytes;
begin
  ValidateAccessToken;

  var Url := BuildUrl(Endpoint);
  Log(LogDebug, Format('%s %s (raw)', [MethodGet, TGraphUrlRedactor.Redact(Url)]));

  var Request := Default(TGraphHttpRequest);
  Request.Method := MethodGet;
  Request.Url := Url;
  Request.Headers := BuildAuthorizationHeader;

  var Response := FTransport.ExecuteBinary(Request);
  if not Response.IsSuccess then
    raise EGraphApiException.Create(Format('HTTP %d fetching raw content', [Response.StatusCode]));

  Result := Response.ContentBytes;
end;

function TGraphHttpClient.GetWithHeaders(const Endpoint: string; const QueryParams: string;
  const ExtraHeaders: TArray<string>): TJSONObject;
begin
  var ParsedHeaders: TArray<TNetHeader>;
  SetLength(ParsedHeaders, Length(ExtraHeaders));

  for var Index := 0 to High(ExtraHeaders) do
  begin
    var Parts := ExtraHeaders[Index].Split([': '], 2);
    if Length(Parts) = 2 then
      ParsedHeaders[Index] := TNetHeader.Create(Parts[0], Parts[1]);
  end;

  FExtraHeaders := ParsedHeaders;
  try
    var Url := BuildUrl(Endpoint, QueryParams);
    LogRequest(MethodGet, Url);
    Result := ExecuteRequest(MethodGet, Url);
  finally
    FExtraHeaders := nil;
  end;
end;

function TGraphHttpClient.GetAbsoluteUrl(const FullUrl: string): TJSONObject;
begin
  LogRequest(MethodGet, FullUrl);
  Result := ExecuteRequest(MethodGet, FullUrl);
end;

function TGraphHttpClient.GetAbsoluteUrlWithHeaders(const FullUrl: string;
  const ExtraHeaders: TArray<string>): TJSONObject;
begin
  var ParsedHeaders: TArray<TNetHeader>;
  SetLength(ParsedHeaders, Length(ExtraHeaders));

  for var Index := 0 to High(ExtraHeaders) do
  begin
    var Parts := ExtraHeaders[Index].Split([': '], 2);
    if Length(Parts) = 2 then
      ParsedHeaders[Index] := TNetHeader.Create(Parts[0], Parts[1]);
  end;

  FExtraHeaders := ParsedHeaders;
  try
    LogRequest(MethodGet, FullUrl);
    Result := ExecuteRequest(MethodGet, FullUrl);
  finally
    FExtraHeaders := nil;
  end;
end;

function TGraphHttpClient.Post(const Endpoint: string; const Body: string): TJSONObject;
begin
  var Url := BuildUrl(Endpoint);
  LogRequest(MethodPost, Url);
  Result := ExecuteRequest(MethodPost, Url, Body);
end;

function TGraphHttpClient.PostRaw(const Endpoint: string; const Body: string): TGraphHttpResponse;
begin
  ValidateAccessToken;

  var Url := BuildUrl(Endpoint);
  LogRequest(MethodPost, Url);

  var Request := Default(TGraphHttpRequest);
  Request.Method  := MethodPost;
  Request.Url     := Url;
  Request.Body    := Body;
  Request.Headers := BuildHeaders;

  Result := FTransport.Execute(Request);

  if not Result.IsSuccess then
    Log(LogError, Format(LogFormatGraphHttpError, [Result.StatusCode, Result.Content]));
end;

function TGraphHttpClient.Patch(const Endpoint: string; const Body: string): TJSONObject;
begin
  var Url := BuildUrl(Endpoint);
  LogRequest(MethodPatch, Url);
  Result := ExecuteRequest(MethodPatch, Url, Body);
end;

function TGraphHttpClient.Delete(const Endpoint: string): TJSONObject;
begin
  var Url := BuildUrl(Endpoint);
  LogRequest(MethodDelete, Url);
  Result := ExecuteRequest(MethodDelete, Url);
end;

function TGraphHttpClient.PutAbsoluteUrlBytes(const FullUrl: string; const Bytes: TBytes;
  const ContentRange: string): TGraphHttpResponse;
begin
  Log(LogDebug, Format('%s %s [%s]', [MethodPut, TGraphUrlRedactor.Redact(FullUrl), ContentRange]));

  var Request := Default(TGraphHttpRequest);
  Request.Method    := MethodPut;
  Request.Url       := FullUrl;
  Request.BodyBytes := Bytes;
  Request.Headers   := BuildUploadHeaders(ContentRange);

  Result := FTransport.Execute(Request);

  if not Result.IsSuccess then
    Log(LogError, Format('Upload HTTP %d - %s', [Result.StatusCode, Result.Content]));
end;

function TGraphHttpClient.DeleteAbsoluteUrl(const FullUrl: string): TGraphHttpResponse;
begin
  LogRequest(MethodDelete, FullUrl);

  var Request := Default(TGraphHttpRequest);
  Request.Method := MethodDelete;
  Request.Url    := FullUrl;

  Result := FTransport.Execute(Request);
end;

function TGraphHttpClient.GetUserPrefix: string;
begin
  if FMailboxAddress.Trim.IsEmpty then
    Result := '/me'
  else
    Result := '/users/' + TNetEncoding.URL.Encode(FMailboxAddress);
end;

function TGraphHttpClient.IsSharedMailbox: Boolean;
begin
  Result := not FMailboxAddress.Trim.IsEmpty;
end;

procedure TGraphHttpClient.SetAccessToken(const Value: string);
begin
  FAccessToken := Value;
end;

function TGraphHttpClient.GetAccessToken: string;
begin
  Result := FAccessToken;
end;

function TGraphHttpClient.GetUserProfile: TUserProfile;
begin
  Result := Default(TUserProfile);
  var Response := Get(GetUserPrefix, '$select=mail,userPrincipalName,displayName');
  try
    if TGraphJson.HasError(Response) then
      Exit;
    Result.Mail := TGraphJson.GetString(Response, 'mail');
    Result.UserPrincipalName := TGraphJson.GetString(Response, 'userPrincipalName');
    Result.DisplayName := TGraphJson.GetString(Response, 'displayName');
  finally
    Response.Free;
  end;
end;

end.
