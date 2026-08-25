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
    Headers: TArray<TNetHeader>;
  end;

  TGraphHttpResponse = record
  public
    StatusCode: Integer;
    Content: string;
    ContentBytes: TBytes;

    function IsSuccess: Boolean;
  end;

implementation

function TGraphHttpResponse.IsSuccess: Boolean;
begin
  Result := (StatusCode >= 200) and (StatusCode < 300);
end;

end.
