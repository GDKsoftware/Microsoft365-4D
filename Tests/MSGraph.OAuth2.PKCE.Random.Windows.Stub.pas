unit MSGraph.OAuth2.PKCE.Random.Windows.Stub;

interface

uses
  MSGraph.OAuth2.PKCE.Random.Windows;

type
  TOAuth2PKCERandomWindowsStub = class(TOAuth2PKCERandomWindows)
  strict private
    FLibraryName: string;

  protected
    function BCryptLibraryName: string; override;

  public
    constructor Create(const LibraryName: string);
  end;

implementation

constructor TOAuth2PKCERandomWindowsStub.Create(const LibraryName: string);
begin
  inherited Create;
  FLibraryName := LibraryName;
end;

function TOAuth2PKCERandomWindowsStub.BCryptLibraryName: string;
begin
  Result := FLibraryName;
end;

end.
