unit MSGraph.Graph.Http;

interface

uses
  System.JSON,
  System.Net.URLClient,
  MSGraph.OAuth2.Types;

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

    function BuildUrl(const Endpoint: string; const QueryParams: string = ''): string;
    function BuildHeaders: TArray<TNetHeader>;
    function ExecuteRequest(const Method: string; const Url: string; const Body: string = ''): TJSONObject;
    function ParseResponse(const StatusCode: Integer; const ResponseText: string): TJSONObject;
    function ParseErrorResponse(const StatusCode: Integer; const ResponseText: string): TJSONObject;
    procedure ValidateAccessToken;
    procedure Log(const Level: string; const Message: string);

    const
      GraphBaseUrl = 'https://graph.microsoft.com/v1.0';
      HeaderAuthorization = 'Authorization';
      HeaderContentType = 'Content-Type';
      BearerPrefix = 'Bearer ';
      ContentTypeJson = 'application/json';
      MethodGet = 'GET';
      MethodPost = 'POST';
      MethodPatch = 'PATCH';
      MethodDelete = 'DELETE';
      LogDebug = 'DEBUG';
      LogError = 'ERROR';
  public
    constructor Create(const AccessToken: string; const LogProc: TLogProc = nil);

    function Get(const Endpoint: string; const QueryParams: string = ''): TJSONObject;
    function Post(const Endpoint: string; const Body: string = ''): TJSONObject;
    function Patch(const Endpoint: string; const Body: string): TJSONObject;
    function Delete(const Endpoint: string): TJSONObject;

    function GetWithHeaders(const Endpoint: string; const QueryParams: string;
      const ExtraHeaders: TArray<string>): TJSONObject;

    function GetUserPrefix: string;
    function IsSharedMailbox: Boolean;

    procedure SetAccessToken(const Value: string);
    function GetAccessToken: string;

    function GetUserProfile: TUserProfile;

    property MailboxAddress: string read FMailboxAddress write FMailboxAddress;
  end;

implementation

uses
  System.SysUtils,
  System.Classes,
  System.NetEncoding,
  System.Net.HttpClient,
  MSGraph.Graph.JsonHelper;

constructor TGraphHttpClient.Create(const AccessToken: string; const LogProc: TLogProc);
begin
  inherited Create;
  FAccessToken := AccessToken;
  FLogProc := LogProc;
end;

procedure TGraphHttpClient.Log(const Level: string; const Message: string);
begin
  if Assigned(FLogProc) then
    FLogProc(Level, Message);
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
    Result[BaseCount + Index] := FExtraHeaders[Index];
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
    Log(LogError, Format('Graph API HTTP %d - %s', [StatusCode, ResponseText]));
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

  var HttpClient := THTTPClient.Create;
  try
    var Headers := BuildHeaders;
    var Response: IHTTPResponse;

    if Method = MethodGet then
      Response := HttpClient.Get(Url, nil, Headers)
    else if Method = MethodPost then
    begin
      var Content: TStringStream := nil;
      if not Body.IsEmpty then
        Content := TStringStream.Create(Body, TEncoding.UTF8);
      try
        Response := HttpClient.Post(Url, Content, nil, Headers);
      finally
        Content.Free;
      end;
    end
    else if Method = MethodPatch then
    begin
      var Content := TStringStream.Create(Body, TEncoding.UTF8);
      try
        Response := HttpClient.Patch(Url, Content, nil, Headers);
      finally
        Content.Free;
      end;
    end
    else if Method = MethodDelete then
    begin
      HttpClient.CustomHeaders[HeaderAuthorization] := BearerPrefix + FAccessToken;
      Response := HttpClient.Delete(Url);
    end
    else
      raise EGraphApiException.Create('Unsupported HTTP method: ' + Method);

    Result := ParseResponse(Response.StatusCode, Response.ContentAsString(TEncoding.UTF8));
  finally
    HttpClient.Free;
  end;
end;

function TGraphHttpClient.Get(const Endpoint: string; const QueryParams: string): TJSONObject;
begin
  var Url := BuildUrl(Endpoint, QueryParams);
  Log(LogDebug, MethodGet + ' ' + Url);
  Result := ExecuteRequest(MethodGet, Url);
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
    Log(LogDebug, MethodGet + ' ' + Url);
    Result := ExecuteRequest(MethodGet, Url);
  finally
    FExtraHeaders := nil;
  end;
end;

function TGraphHttpClient.Post(const Endpoint: string; const Body: string): TJSONObject;
begin
  var Url := BuildUrl(Endpoint);
  Log(LogDebug, MethodPost + ' ' + Url);
  Result := ExecuteRequest(MethodPost, Url, Body);
end;

function TGraphHttpClient.Patch(const Endpoint: string; const Body: string): TJSONObject;
begin
  var Url := BuildUrl(Endpoint);
  Log(LogDebug, MethodPatch + ' ' + Url);
  Result := ExecuteRequest(MethodPatch, Url, Body);
end;

function TGraphHttpClient.Delete(const Endpoint: string): TJSONObject;
begin
  var Url := BuildUrl(Endpoint);
  Log(LogDebug, MethodDelete + ' ' + Url);
  Result := ExecuteRequest(MethodDelete, Url);
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
