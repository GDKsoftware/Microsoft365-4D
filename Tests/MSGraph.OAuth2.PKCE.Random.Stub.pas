unit MSGraph.OAuth2.PKCE.Random.Stub;

interface

uses
  System.SysUtils,
  MSGraph.OAuth2.PKCE.Random;

type
  TOAuth2PKCERandomStub = class(TOAuth2PKCERandom)
  protected
    procedure FillRandomBytes(const Buffer: TBytes); override;
  end;

implementation

const
  ByteValueCount = 256;

procedure TOAuth2PKCERandomStub.FillRandomBytes(const Buffer: TBytes);
begin
  for var Index := 0 to High(Buffer) do
  begin
    Buffer[Index] := Index mod ByteValueCount;
  end;
end;

end.
