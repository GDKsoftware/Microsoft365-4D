unit MSGraph.OAuth2.Client;

interface

uses
  System.Classes,
  MSGraph.OAuth2.Types;

type
  TOAuth2Client = class
  strict private
    FConfig: TOAuth2Config;
    FLogProc: TLogProc;

    function BuildScopeString(const Scopes: TArray<string>): string;
    function ParseTokenResponse(const JsonResponse: string): TOAuth2TokenResponse;
    function PostTokenRequest(const PostData: TStringList): TOAuth2TokenResponse;
    procedure Log(const Level: string; const Message: string);

    const
      ContentTypeFormUrlEncoded = 'application/x-www-form-urlencoded';
      LogInfo = 'INFO';
      LogDebug = 'DEBUG';
  public
    constructor Create(const Config: TOAuth2Config; const LogProc: TLogProc = nil);

    function GenerateAuthorizationUrl(const PKCESession: TPKCESession): string;

    function ExchangeCodeForToken(
      const Code: string;
      const CodeVerifier: string): TOAuth2TokenResponse;

    function RefreshAccessToken(const RefreshToken: string): TOAuth2TokenResponse;

    property Config: TOAuth2Config read FConfig;
  end;

implementation

uses
  System.SysUtils,
  System.JSON,
  System.Net.HttpClient,
  System.Net.URLClient,
  System.NetConsts,
  System.NetEncoding,
  System.DateUtils;

constructor TOAuth2Client.Create(const Config: TOAuth2Config; const LogProc: TLogProc);
begin
  inherited Create;
  FConfig := Config;
  FLogProc := LogProc;
end;

procedure TOAuth2Client.Log(const Level: string; const Message: string);
begin
  if Assigned(FLogProc) then
    FLogProc(Level, Message);
end;

function TOAuth2Client.BuildScopeString(const Scopes: TArray<string>): string;
begin
  Result := '';
  for var Scope in Scopes do
  begin
    if not Result.IsEmpty then
      Result := Result + ' ';
    Result := Result + Scope;
  end;
end;

function TOAuth2Client.ParseTokenResponse(const JsonResponse: string): TOAuth2TokenResponse;
begin
  var JsonObj := TJSONObject.ParseJSONValue(JsonResponse) as TJSONObject;
  if not Assigned(JsonObj) then
    raise EOAuth2Exception.Create('Invalid JSON response from token endpoint');

  try
    Result.AccessToken := JsonObj.GetValue<string>('access_token', '');
    Result.RefreshToken := JsonObj.GetValue<string>('refresh_token', '');
    Result.ExpiresIn := JsonObj.GetValue<Integer>('expires_in', 0);
    Result.TokenType := JsonObj.GetValue<string>('token_type', 'Bearer');
    Result.Scope := JsonObj.GetValue<string>('scope', '');
    Result.ExpiresAt := IncSecond(Now, Result.ExpiresIn);

    if Result.AccessToken.IsEmpty then
      raise EOAuth2Exception.Create('Access token not found in response');
  finally
    JsonObj.Free;
  end;
end;

function TOAuth2Client.PostTokenRequest(const PostData: TStringList): TOAuth2TokenResponse;
begin
  var HttpClient := THTTPClient.Create;
  try
    HttpClient.ContentType := ContentTypeFormUrlEncoded;
    var Response := HttpClient.Post(FConfig.TokenUrl, PostData);
    if Response.StatusCode <> 200 then
    begin
      raise EOAuth2Exception.CreateFmt('Token request failed with status %d: %s',
        [Response.StatusCode, Response.ContentAsString]);
    end;

    Result := ParseTokenResponse(Response.ContentAsString);
  finally
    HttpClient.Free;
  end;
end;

function TOAuth2Client.GenerateAuthorizationUrl(const PKCESession: TPKCESession): string;
begin
  var Scope := BuildScopeString(FConfig.Scopes);

  Result := FConfig.AuthUrl +
    '?response_type=code' +
    '&client_id=' + TNetEncoding.URL.Encode(FConfig.ClientId) +
    '&redirect_uri=' + TNetEncoding.URL.Encode(FConfig.RedirectUri) +
    '&scope=' + TNetEncoding.URL.Encode(Scope) +
    '&state=' + TNetEncoding.URL.Encode(PKCESession.State) +
    '&code_challenge=' + TNetEncoding.URL.Encode(PKCESession.CodeChallenge) +
    '&code_challenge_method=S256';

  if not FConfig.ExtraParams.Trim.IsEmpty then
    Result := Result + '&' + FConfig.ExtraParams;

  Log(LogDebug, 'Authorization URL generated');
end;

function TOAuth2Client.ExchangeCodeForToken(
  const Code: string;
  const CodeVerifier: string): TOAuth2TokenResponse;
begin
  Log(LogInfo, 'Exchanging authorization code for tokens');

  var PostData := TStringList.Create;
  try
    PostData.Add('grant_type=authorization_code');
    PostData.Add('code=' + Code);
    PostData.Add('redirect_uri=' + FConfig.RedirectUri);
    PostData.Add('client_id=' + FConfig.ClientId);
    PostData.Add('client_secret=' + FConfig.ClientSecret);
    PostData.Add('code_verifier=' + CodeVerifier);

    Result := PostTokenRequest(PostData);
    Log(LogInfo, 'Token exchange successful, expires in ' + IntToStr(Result.ExpiresIn) + 's');
  finally
    PostData.Free;
  end;
end;

function TOAuth2Client.RefreshAccessToken(const RefreshToken: string): TOAuth2TokenResponse;
begin
  Log(LogInfo, 'Refreshing access token');

  var PostData := TStringList.Create;
  try
    PostData.Add('grant_type=refresh_token');
    PostData.Add('refresh_token=' + RefreshToken);
    PostData.Add('client_id=' + FConfig.ClientId);
    PostData.Add('client_secret=' + FConfig.ClientSecret);

    Result := PostTokenRequest(PostData);
    Log(LogInfo, 'Token refresh successful, expires in ' + IntToStr(Result.ExpiresIn) + 's');
  finally
    PostData.Free;
  end;
end;

end.
