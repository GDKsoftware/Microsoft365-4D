unit MSGraph.OAuth2.PKCE.Random.Stub;

interface

uses
  System.SysUtils,
  MSGraph.OAuth2.PKCE.Random;

type
  TOAuth2PKCERandomStub = class(TOAuth2PKCERandom)
  protected
    function GenerateRandomBytes(const Count: Integer): TBytes; override;
  end;

implementation

const
  ByteValueCount = 256;

function TOAuth2PKCERandomStub.GenerateRandomBytes(const Count: Integer): TBytes;
begin
  SetLength(Result, Count);

  for var Index := 0 to Count - 1 do
  begin
    Result[Index] := Index mod ByteValueCount;
  end;
end;

end.
