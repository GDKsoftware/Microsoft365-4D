unit MSGraph.Graph.People.Interfaces;

interface

uses
  MSGraph.Graph.People.Types;

type
  IPeopleClient = interface
    ['{E8F9A0B1-C2D3-4E5F-6A7B-8C9D0E1F2A3B}']
    function GetRelevantPeople(const Top: Integer = 20): TArray<TPerson>;
    function SearchPeople(const Query: string; const Top: Integer = 10): TArray<TPerson>;
  end;

implementation

end.
