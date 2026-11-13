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

    class function CreateRandom: IOAuth2PKCERandom; static;

  public
    class function Generate: TPKCESession; overload;
    class function Generate(const Random: IOAuth2PKCERandom): TPKCESession; overload;
    class function GenerateCodeChallenge(const CodeVerifier: string): string; static;
  end;

implementation

uses
  System.Hash,
  System.NetEncoding,
  {$IFDEF MSWINDOWS}
  MSGraph.OAuth2.PKCE.Random.Windows;
  {$ELSE}
  MSGraph.OAuth2.PKCE.Random.Posix;
  {$ENDIF}

class function TOAuth2PKCE.Generate: TPKCESession;
begin
  const PlatformRandom = CreateRandom;

  Result := Generate(PlatformRandom);
end;

class function TOAuth2PKCE.CreateRandom: IOAuth2PKCERandom;
begin
  {$IFDEF MSWINDOWS}
  Result := TOAuth2PKCERandomWindows.Create;
  {$ELSE}
  Result := TOAuth2PKCERandomPosix.Create;
  {$ENDIF}
end;

class function TOAuth2PKCE.Generate(const Random: IOAuth2PKCERandom): TPKCESession;
begin
  Result.CodeVerifier := Random.GenerateSecureRandomString(CodeVerifierLength);
  Result.CodeChallenge := GenerateCodeChallenge(Result.CodeVerifier);
  Result.State := Random.GenerateNonce;
  Result.ExpiresAt := Now + (StateTimeoutMinutes / (24 * 60));
end;

class function TOAuth2PKCE.GenerateCodeChallenge(const CodeVerifier: string): string;
begin
  const Hash = THashSHA2.GetHashBytes(CodeVerifier);
  Result := TNetEncoding.Base64URL.EncodeBytesToString(Hash);
end;

end.
