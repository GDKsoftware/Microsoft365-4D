unit MSGraph.Graph.Http.Interfaces;

interface

uses
  MSGraph.Graph.Http.Types;

type
  IGraphHttpTransport = interface
    ['{529F25D2-F41C-4F0D-BFA1-EF799BC954D0}']
    function Execute(const Request: TGraphHttpRequest): TGraphHttpResponse;
    function ExecuteBinary(const Request: TGraphHttpRequest): TGraphHttpResponse;
  end;

implementation

end.
