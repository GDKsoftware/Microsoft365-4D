unit MSGraph.OAuth2.PKCE;

interface

uses
  System.SysUtils,
  MSGraph.OAuth2.Types;

type
  TOAuth2PKCE = class
  strict private
    const
      CodeVerifierLength = 128;
      StateTimeoutMinutes = 5;

    class function Base64UrlEncode(const Data: TBytes): string;
    class function GenerateSecureRandomString(const Length: Integer): string;
    class function GenerateCodeChallenge(const CodeVerifier: string): string;
    class function GenerateNonce: string;
  public
    class function Generate: TPKCESession;
  end;

implementation

uses
  System.NetEncoding,
  System.Hash;

class function TOAuth2PKCE.Base64UrlEncode(const Data: TBytes): string;
begin
  Result := TNetEncoding.Base64.EncodeBytesToString(Data);
  Result := Result.Replace('+', '-').Replace('/', '_').Replace('=', '');
end;

class function TOAuth2PKCE.GenerateSecureRandomString(const Length: Integer): string;
const
  Base64UrlChars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-_';
begin
  var RandomBytes: TBytes;
  SetLength(RandomBytes, Length);

  for var Index := 0 to Length - 1 do
    RandomBytes[Index] := Random(256);

  Result := '';
  for var Index := 0 to Length - 1 do
    Result := Result + Base64UrlChars[(RandomBytes[Index] mod 64) + 1];
end;

class function TOAuth2PKCE.GenerateCodeChallenge(const CodeVerifier: string): string;
begin
  var Hash := THashSHA2.GetHashBytes(CodeVerifier);
  Result := Base64UrlEncode(Hash);
end;

class function TOAuth2PKCE.GenerateNonce: string;
var
  Guid: TGUID;
begin
  CreateGUID(Guid);
  Result := GUIDToString(Guid).Replace('{', '').Replace('}', '').Replace('-', '');
end;

class function TOAuth2PKCE.Generate: TPKCESession;
begin
  Randomize;
  Result.CodeVerifier := GenerateSecureRandomString(CodeVerifierLength);
  Result.CodeChallenge := GenerateCodeChallenge(Result.CodeVerifier);
  Result.State := GenerateNonce;
  Result.ExpiresAt := Now + (StateTimeoutMinutes / (24 * 60));
end;

end.
