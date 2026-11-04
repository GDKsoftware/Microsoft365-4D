unit MSGraph.OAuth2.PKCE;

interface

uses
  System.SysUtils,
  MSGraph.OAuth2.Types,
  MSGraph.OAuth2.PKCE.Random.Interfaces;

type
  TOAuth2PKCE = class
  strict private
    const
      CodeVerifierLength = 128;
      StateTimeoutMinutes = 5;
  public
    class function Generate: TPKCESession; overload;
    class function Generate(const Random: IOAuth2PKCERandom): TPKCESession; overload;
  end;

implementation

uses
  MSGraph.OAuth2.PKCE.Random;

class function TOAuth2PKCE.Generate: TPKCESession;
begin
  Result := Generate(TOAuth2PKCERandom.Create);
end;

class function TOAuth2PKCE.Generate(const Random: IOAuth2PKCERandom): TPKCESession;
begin
  Result.CodeVerifier := Random.GenerateSecureRandomString(CodeVerifierLength);
  Result.CodeChallenge := Random.GenerateCodeChallenge(Result.CodeVerifier);
  Result.State := Random.GenerateNonce;
  Result.ExpiresAt := Now + (StateTimeoutMinutes / (24 * 60));
end;

end.
