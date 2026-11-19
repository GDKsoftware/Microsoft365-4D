unit MSGraph.OAuth2.PKCE.Random;

interface

uses
  MSGraph.OAuth2.PKCE.Random.Interfaces,
  System.SysUtils,
  System.NetEncoding,
  System.Hash;

type
  TOAuth2PKCERandom = class abstract(TInterfacedObject, IOAuth2PKCERandom)
  strict private
    function GenerateRandomBytes(const Count: Integer): TBytes;

  protected
    procedure FillRandomBytes(const Buffer: TBytes); virtual; abstract;

  public
    function GenerateSecureRandomString(const CharacterCount: Integer): string;
    function GenerateNonce: string;
  end;

implementation

{ TOAuth2PKCERandom }

function TOAuth2PKCERandom.GenerateNonce: string;
begin
  const NonceByteCount = 16;

  const RandomBytes = GenerateRandomBytes(NonceByteCount);

  Result := THash.DigestAsString(RandomBytes);
end;

function TOAuth2PKCERandom.GenerateSecureRandomString(const CharacterCount: Integer): string;
begin
  const BytesPerGroup      = 3;
  const CharactersPerGroup = 4;

  const ByteCount   = (CharacterCount * BytesPerGroup + CharactersPerGroup - 1) div CharactersPerGroup;
  const RandomBytes = GenerateRandomBytes(ByteCount);
  const Encoded     = TNetEncoding.Base64URL.EncodeBytesToString(RandomBytes);

  Result := Encoded.Substring(0, CharacterCount);
end;

function TOAuth2PKCERandom.GenerateRandomBytes(const Count: Integer): TBytes;
begin
  SetLength(Result, Count);

  const IsEmpty = (Count = 0);
  if IsEmpty then
    Exit;

  FillRandomBytes(Result);
end;

end.
