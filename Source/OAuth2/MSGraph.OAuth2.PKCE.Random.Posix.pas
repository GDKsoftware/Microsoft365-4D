unit MSGraph.OAuth2.PKCE.Random.Posix;

interface

{$IFDEF POSIX}

uses
  System.SysUtils,
  MSGraph.OAuth2.PKCE.Random;

type
  TOAuth2PKCERandomPosix = class(TOAuth2PKCERandom)
  strict private
    procedure FillRandomBytes(const Buffer: TBytes);

  protected
    function GenerateRandomBytes(const Count: Integer): TBytes; override;
  end;

{$ENDIF}

implementation

{$IFDEF POSIX}

uses
  System.Classes,
  MSGraph.OAuth2.Types;

const
  RandomDevicePath = '/dev/urandom';

function TOAuth2PKCERandomPosix.GenerateRandomBytes(const Count: Integer): TBytes;
begin
  SetLength(Result, Count);

  const IsEmpty = (Count = 0);
  if IsEmpty then
    Exit;

  FillRandomBytes(Result);
end;

procedure TOAuth2PKCERandomPosix.FillRandomBytes(const Buffer: TBytes);
begin
  const BufferSize = Length(Buffer);
  try
    const Device = TFileStream.Create(RandomDevicePath, fmOpenRead or fmShareDenyNone);
    try
      Device.ReadBuffer(Buffer[0], BufferSize);
    finally
      Device.Free;
    end;
  except
    on Error: EStreamError do
      raise EOAuth2RandomException.CreateFmt('Random bytes could not be read from %s: %s', [RandomDevicePath, Error.Message]);
  end;
end;

{$ENDIF}

end.
