unit MSGraph.OAuth2.PKCE.Random.Windows;

interface

{$IFDEF MSWINDOWS}

uses
  System.SysUtils,
  MSGraph.OAuth2.PKCE.Random;

type
  TOAuth2PKCERandomWindows = class(TOAuth2PKCERandom)
  strict private
    procedure GuardStatus(const Status: Integer);

  protected
    function GenerateRandomBytes(const Count: Integer): TBytes; override;
  end;

{$ENDIF}

implementation

{$IFDEF MSWINDOWS}

uses
  MSGraph.OAuth2.Types;

const
  BCryptLibrary = 'bcrypt.dll';
  BCryptUseSystemPreferredRng = $00000002;
  StatusSuccess = 0;

function BCryptGenRandom(
  const Algorithm: THandle;
  const Buffer: PByte;
  const BufferSize: Cardinal;
  const Flags: Cardinal): Integer; stdcall; external BCryptLibrary;

{ TOAuth2PKCERandomWindows }

function TOAuth2PKCERandomWindows.GenerateRandomBytes(const Count: Integer): TBytes;
begin
  SetLength(Result, Count);

  const IsEmpty = (Count = 0);
  if IsEmpty then
    Exit;

  const Status = BCryptGenRandom(0, @Result[0], Count, BCryptUseSystemPreferredRng);
  GuardStatus(Status);
end;

procedure TOAuth2PKCERandomWindows.GuardStatus(const Status: Integer);
begin
  const Succeeded = (Status = StatusSuccess);
  if Succeeded then
    Exit;

  raise EOAuth2RandomException.CreateFmt('BCryptGenRandom failed with status 0x%.8x', [Status]);
end;

{$ENDIF}

end.
