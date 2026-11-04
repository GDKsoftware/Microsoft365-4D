unit MSGraph.OAuth2.PKCE.Random;

interface

uses
  MSGraph.OAuth2.PKCE.Random.Interfaces,
  System.SysUtils,
  System.NetEncoding,
  System.Hash;

type
  TOAuth2PKCERandom = class(TInterfacedObject, IOAuth2PKCERandom)
  public
    function Base64UrlEncode(const Data: TBytes): string;
    function GenerateSecureRandomString(const Length: Integer): string;
    function GenerateCodeChallenge(const CodeVerifier: string): string;
    function GenerateNonce: string;
  end;

implementation

{ TOAuth2PKCERandom }

function TOAuth2PKCERandom.Base64UrlEncode(const Data: TBytes): string;
begin
  Result := TNetEncoding.Base64String.EncodeBytesToString(Data);
  Result := Result.Replace('+', '-').Replace('/', '_').Replace('=', '');
end;

function TOAuth2PKCERandom.GenerateCodeChallenge(const CodeVerifier: string): string;
begin
  var Hash := THashSHA2.GetHashBytes(CodeVerifier);
  Result := Base64UrlEncode(Hash);
end;

function TOAuth2PKCERandom.GenerateNonce: string;
var
  Guid: TGUID;
begin
  CreateGUID(Guid);
  Result := GUIDToString(Guid).Replace('{', '').Replace('}', '').Replace('-', '');
end;

function TOAuth2PKCERandom.GenerateSecureRandomString(const Length: Integer): string;
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

end.
