unit MSGraph.OAuth2.PKCE.Random;

interface

uses
  MSGraph.OAuth2.PKCE.Random.Interfaces,
  System.SysUtils,
  System.NetEncoding,
  System.Hash;

type
  TOAuth2PKCERandom = class(TInterfacedObject, IOAuth2PKCERandom)
  protected
    function GenerateRandomBytes(const Count: Integer): TBytes; virtual;

  public
    function GenerateSecureRandomString(const Length: Integer): string;
    function GenerateCodeChallenge(const CodeVerifier: string): string;
    function GenerateNonce: string;
  end;

implementation

{ TOAuth2PKCERandom }

function TOAuth2PKCERandom.GenerateCodeChallenge(const CodeVerifier: string): string;
begin
  const Hash = THashSHA2.GetHashBytes(CodeVerifier);
  Result := TNetEncoding.Base64URL.EncodeBytesToString(Hash);
end;

function TOAuth2PKCERandom.GenerateNonce: string;
begin
  const DigitsOnlyFormat = 'N';

  const Guid = TGUID.NewGuid;

  Result := Guid.ToString(DigitsOnlyFormat);
end;

function TOAuth2PKCERandom.GenerateSecureRandomString(const Length: Integer): string;
begin
  const BytesPerGroup      = 3;
  const CharactersPerGroup = 4;

  const ByteCount   = (Length * BytesPerGroup + CharactersPerGroup - 1) div CharactersPerGroup;
  const RandomBytes = GenerateRandomBytes(ByteCount);
  const Encoded     = TNetEncoding.Base64URL.EncodeBytesToString(RandomBytes);

  Result := Encoded.Substring(0, Length);
end;

function TOAuth2PKCERandom.GenerateRandomBytes(const Count: Integer): TBytes;
begin
  SetLength(Result, Count);

  for var Index := 0 to Count - 1 do
    Result[Index] := Random(256);
end;

end.
