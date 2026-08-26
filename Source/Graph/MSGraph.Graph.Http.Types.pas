unit MSGraph.Graph.Http.Types;

interface

uses
  System.SysUtils,
  System.Net.URLClient;

type
  TGraphHttpRequest = record
  public
    Method: string;
    Url: string;
    Body: string;
    BodyBytes: TBytes;
    Headers: TArray<TNetHeader>;
  end;

  TGraphHttpResponse = record
  public
    StatusCode: Integer;
    Content: string;
    ContentBytes: TBytes;
    Headers: TArray<TNetHeader>;

    function IsSuccess: Boolean;
    function HeaderValue(const Name: string): string;
  end;

implementation

function TGraphHttpResponse.IsSuccess: Boolean;
begin
  Result := (StatusCode >= 200) and (StatusCode < 300);
end;

function TGraphHttpResponse.HeaderValue(const Name: string): string;
begin
  Result := '';

  for var Header in Headers do
  begin
    if SameText(Header.Name, Name) then
      Exit(Header.Value);
  end;
end;

end.
