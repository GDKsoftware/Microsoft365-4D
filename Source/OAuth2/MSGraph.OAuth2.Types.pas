unit MSGraph.OAuth2.Types;

interface

uses
  System.SysUtils,
  System.DateUtils;

type
  TLogProc = reference to procedure(const Level: string; const Message: string);

  EMSGraphException = class(Exception);
  EOAuth2Exception = class(EMSGraphException);
  EGraphApiException = class(EMSGraphException);
  ETokenStoreException = class(EMSGraphException);

  TOAuth2Config = record
    ClientId: string;
    ClientSecret: string;
    TenantId: string;
    RedirectUri: string;
    Port: Integer;
    Scopes: TArray<string>;
    ExtraParams: string;
    function AuthUrl: string;
    function TokenUrl: string;
  end;

  TOAuth2TokenResponse = record
    AccessToken: string;
    RefreshToken: string;
    ExpiresIn: Integer;
    TokenType: string;
    Scope: string;
    ExpiresAt: TDateTime;
    function IsExpired: Boolean;
    function IsExpiringSoon(const MarginSeconds: Integer = 300): Boolean;
  end;

  TPKCESession = record
    State: string;
    CodeVerifier: string;
    CodeChallenge: string;
    ExpiresAt: TDateTime;
  end;

implementation

const
  MicrosoftLoginBaseUrl = 'https://login.microsoftonline.com/';
  OAuth2PathAuthorize = '/oauth2/v2.0/authorize';
  OAuth2PathToken = '/oauth2/v2.0/token';

function TOAuth2Config.AuthUrl: string;
begin
  Result := MicrosoftLoginBaseUrl + TenantId + OAuth2PathAuthorize;
end;

function TOAuth2Config.TokenUrl: string;
begin
  Result := MicrosoftLoginBaseUrl + TenantId + OAuth2PathToken;
end;

function TOAuth2TokenResponse.IsExpired: Boolean;
begin
  Result := (Now >= ExpiresAt);
end;

function TOAuth2TokenResponse.IsExpiringSoon(const MarginSeconds: Integer): Boolean;
begin
  Result := (Now >= IncSecond(ExpiresAt, -MarginSeconds));
end;

end.
