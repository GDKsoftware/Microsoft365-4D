unit MSGraph.OAuth2.PKCE.Characters;

interface

uses
  System.SysUtils;

type
  TPKCECharacters = class sealed
  public
    const
      UrlSafe = ['A'..'Z', 'a'..'z', '0'..'9', '-', '_'];

    class function ContainsOnly(const Value: string; const Allowed: TSysCharSet): Boolean; static;
  end;

implementation

{ TPKCECharacters }

class function TPKCECharacters.ContainsOnly(const Value: string; const Allowed: TSysCharSet): Boolean;
begin
  for var Character in Value do
  begin
    if not CharInSet(Character, Allowed) then
      Exit(False);
  end;

  Result := True;
end;

end.
