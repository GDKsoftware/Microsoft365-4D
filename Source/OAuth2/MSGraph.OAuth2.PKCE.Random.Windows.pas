unit MSGraph.OAuth2.PKCE.Random.Windows;

interface

{$IFDEF MSWINDOWS}

uses
  System.SysUtils,
  MSGraph.OAuth2.PKCE.Random;

type
  TOAuth2PKCERandomWindows = class(TOAuth2PKCERandom)
  strict private
    type
      TBCryptGenRandom = function(const Algorithm: THandle;
                                  const Buffer: PByte;
                                  const BufferSize: Cardinal;
                                  const Flags: Cardinal): Integer; stdcall;

    procedure FillRandomBytes(const Buffer: TBytes);
    function LoadBCryptLibrary: HMODULE;
    procedure GuardLibraryLoaded(const Module: HMODULE; const LibraryName: string);
    function GetBCryptGenRandom(const Module: HMODULE): TBCryptGenRandom;
    procedure GuardFunctionExported(const Address: Pointer);
    procedure GuardStatus(const Status: Integer);

  protected
    function GenerateRandomBytes(const Count: Integer): TBytes; override;
    function BCryptLibraryName: string; virtual;
  end;

{$ENDIF}

implementation

{$IFDEF MSWINDOWS}

uses
  Winapi.Windows,
  MSGraph.OAuth2.Types;

const
  BCryptLibrary               = 'bcrypt.dll';
  BCryptGenRandomName         = 'BCryptGenRandom';
  BCryptUseSystemPreferredRng = $00000002;
  LoadLibrarySearchSystem32   = $00000800;
  StatusSuccess               = 0;

function TOAuth2PKCERandomWindows.GenerateRandomBytes(const Count: Integer): TBytes;
begin
  SetLength(Result, Count);

  const IsEmpty = (Count = 0);
  if IsEmpty then
    Exit;

  FillRandomBytes(Result);
end;

procedure TOAuth2PKCERandomWindows.FillRandomBytes(const Buffer: TBytes);
begin
  const BufferSize = Length(Buffer);
  const Module     = LoadBCryptLibrary;
  try
    const GenerateRandom: TBCryptGenRandom = GetBCryptGenRandom(Module);

    const Status = GenerateRandom(0, @Buffer[0], BufferSize, BCryptUseSystemPreferredRng);
    GuardStatus(Status);
  finally
    FreeLibrary(Module);
  end;
end;

function TOAuth2PKCERandomWindows.LoadBCryptLibrary: HMODULE;
begin
  const LibraryName = BCryptLibraryName;

  Result := LoadLibraryEx(PChar(LibraryName), 0, LoadLibrarySearchSystem32);
  GuardLibraryLoaded(Result, LibraryName);
end;

function TOAuth2PKCERandomWindows.BCryptLibraryName: string;
begin
  Result := BCryptLibrary;
end;

procedure TOAuth2PKCERandomWindows.GuardLibraryLoaded(const Module: HMODULE; const LibraryName: string);
begin
  const IsLoaded = (Module <> 0);
  if IsLoaded then
    Exit;

  const ErrorCode = GetLastError;
  raise EOAuth2RandomException.CreateFmt('%s could not be loaded from the system directory (error %d)', [LibraryName, ErrorCode]);
end;

function TOAuth2PKCERandomWindows.GetBCryptGenRandom(const Module: HMODULE): TBCryptGenRandom;
begin
  const Address = GetProcAddress(Module, PChar(BCryptGenRandomName));
  GuardFunctionExported(Address);

  Result := TBCryptGenRandom(Address);
end;

procedure TOAuth2PKCERandomWindows.GuardFunctionExported(const Address: Pointer);
begin
  const IsExported = Assigned(Address);
  if IsExported then
    Exit;

  const ErrorCode = GetLastError;
  raise EOAuth2RandomException.CreateFmt('%s does not export %s (error %d)', [BCryptLibraryName, BCryptGenRandomName, ErrorCode]);
end;

procedure TOAuth2PKCERandomWindows.GuardStatus(const Status: Integer);
begin
  const Succeeded = (Status = StatusSuccess);
  if Succeeded then
    Exit;

  raise EOAuth2RandomException.CreateFmt('%s failed with status 0x%.8x', [BCryptGenRandomName, Status]);
end;

{$ENDIF}

end.
