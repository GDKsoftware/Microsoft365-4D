unit MSGraph.Graph.Drive.Interfaces;

interface

uses
  System.SysUtils,
  MSGraph.Graph.Drive.Types;

type
  /// <summary>
  /// File and folder operations on one drive (OneDrive or a SharePoint document
  /// library). Paths are relative to the drive root, use <c>/</c> as separator
  /// and are not URL-encoded; an empty path is the drive root.
  /// </summary>
  IGraphDriveClient = interface
    ['{6E2B9C14-8F3A-4D57-A1E0-5C7D9B2F4A68}']
    /// GET {drive}/root:/{path}
    function GetItem(const Path: string): TGraphDriveItem;
    /// GET {drive}/root:/{path}:/children - the first page.
    function ListChildren(const Path: string): TGraphDrivePage;
    /// GET {nextLink} - the next page of a listing or a delta query.
    function ListNext(const NextLink: string): TGraphDrivePage;
    /// GET {downloadUrl} without an Authorization header.
    function DownloadContent(const DownloadUrl: string): TBytes;
    /// PUT {drive}/root:/{path}:/content
    function Upload(const Request: TGraphDriveUploadRequest; const Content: TBytes): TGraphDriveItem;
    /// POST {drive}/root:/{parent}:/children - fails when the name exists.
    function CreateFolder(const ParentPath, Name: string): TGraphDriveItem;
    /// PATCH {drive}/root:/{path} with a new parent and name.
    function Move(const Path, NewParentId, NewName: string): TGraphDriveItem;
    /// DELETE {drive}/root:/{path} - moves the item to the recycle bin.
    procedure Delete(const Path: string);
    /// GET {drive}/root/delta, or the given next or delta link.
    function Delta(const Link: string): TGraphDrivePage;
    /// GET {drive}/root/delta?token=latest - a delta link that starts now.
    function LatestDeltaLink: string;
  end;

implementation

end.
