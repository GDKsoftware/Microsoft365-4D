unit MSGraph.Graph.SharePoint.Interfaces;

interface

uses
  MSGraph.Graph.SharePoint.Types;

type
  ISharePointClient = interface
    ['{E5F6A7B8-C9D0-4E5F-2A3B-4C5D6E7F8A9B}']
    function ListSites(const Query: string; const Top: Integer = 25): TArray<TSite>;
    function GetSite(const SiteId: string): TSite;
    function ListDriveItems(const SiteId: string; const FolderId: string;
      const Top: Integer = 50): TArray<TDriveItem>;
    function SearchDriveItems(const SiteId: string; const Query: string;
      const Top: Integer = 25): TArray<TDriveItem>;
    function GetDriveItemContent(const SiteId: string; const ItemId: string): TDriveItem;
  end;

implementation

end.
