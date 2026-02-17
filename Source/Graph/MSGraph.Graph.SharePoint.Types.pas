unit MSGraph.Graph.SharePoint.Types;

interface

type
  TSite = record
    Id: string;
    Name: string;
    DisplayName: string;
    WebUrl: string;
    Description: string;
    CreatedDateTime: string;
    LastModifiedDateTime: string;
  end;

  TDriveItem = record
    Id: string;
    Name: string;
    Size: Int64;
    WebUrl: string;
    ItemType: string;
    MimeType: string;
    ChildCount: Integer;
    ParentPath: string;
    DownloadUrl: string;
    CreatedDateTime: string;
    LastModifiedDateTime: string;
  end;

implementation

end.
