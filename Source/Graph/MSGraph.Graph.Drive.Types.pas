unit MSGraph.Graph.Drive.Types;

interface

const
  GraphDriveDefaultBaseUrl    = 'https://graph.microsoft.com/v1.0';
  GraphMyDrivePath            = '/me/drive';
  GraphDrivePathFormat        = '/drives/%s';
  GraphDriveRootAddress       = '/root';
  GraphDriveItemAddressFormat = '/items/%s';

type
  {$SCOPEDENUMS ON}
  TGraphDriveItemKind = (Unknown, FileItem, FolderItem);
  TGraphDriveWriteMode = (Replace, FailIfExists, IfMatch);
  {$SCOPEDENUMS OFF}

  /// <summary>
  /// Where a drive client points: the base URL of Microsoft Graph and the drive
  /// segment (<c>/me/drive</c> for the signed-in user's OneDrive, or
  /// <c>/drives/{drive-id}</c> for a SharePoint document library). The root
  /// address is the item that paths are relative to: <c>/root</c>, or
  /// <c>/items/{id}</c> for a folder in someone else's drive, such as a shared
  /// folder.
  /// </summary>
  TGraphDriveConfig = record
    AccessToken: string;
    BaseUrl: string;
    DrivePath: string;
    RootAddress: string;

    class function ForMyDrive(const AccessToken: string): TGraphDriveConfig; static;
    class function ForDrive(const AccessToken, DriveId: string): TGraphDriveConfig; static;
    class function ForFolder(const AccessToken, DriveId, FolderId: string): TGraphDriveConfig; static;
  end;

  /// <summary>
  /// A file or folder in a drive. <c>ETag</c> changes with every change to the
  /// item and is the value to send in <c>If-Match</c>; <c>DownloadUrl</c> is a
  /// short-lived, pre-authenticated link (files only); <c>Raw</c> holds the full
  /// item JSON.
  /// </summary>
  TGraphDriveItem = record
    Id: string;
    Name: string;
    Kind: TGraphDriveItemKind;
    ETag: string;
    CTag: string;
    Size: Int64;
    LastModified: string;
    ParentId: string;
    IsDeleted: Boolean;
    DownloadUrl: string;
    Raw: string;
  end;

  /// <summary>
  /// One page of a listing or a delta query. While <c>NextLink</c> is set more
  /// pages follow; a delta query ends with a <c>DeltaLink</c> to ask for the
  /// changes since.
  /// </summary>
  TGraphDrivePage = record
    Items: TArray<TGraphDriveItem>;
    NextLink: string;
    DeltaLink: string;

    function HasMore: Boolean;
  end;

  TGraphDriveUploadRequest = record
    Path: string;
    Mode: TGraphDriveWriteMode;
    /// The eTag the item must still have; required when <c>Mode</c> is IfMatch.
    ETag: string;

    class function ForPath(const Path: string; const Mode: TGraphDriveWriteMode): TGraphDriveUploadRequest; static;
  end;

implementation

uses
  System.SysUtils,
  MSGraph.Graph.Drive.Exceptions,
  MSGraph.Graph.Drive.Path;

const
  SMissingDriveId = 'A folder address needs the id of the drive';
  SMissingFolderId = 'A folder address needs the id of the folder';
  SUnsafeFolderId = 'A folder address needs a real folder id, not "." or ".."';
  CurrentSegment = '.';
  ParentSegment = '..';

class function TGraphDriveConfig.ForMyDrive(const AccessToken: string): TGraphDriveConfig;
begin
  Result.AccessToken := AccessToken;
  Result.BaseUrl     := GraphDriveDefaultBaseUrl;
  Result.DrivePath   := GraphMyDrivePath;
  Result.RootAddress := GraphDriveRootAddress;
end;

class function TGraphDriveConfig.ForDrive(const AccessToken, DriveId: string): TGraphDriveConfig;
begin
  Result.AccessToken := AccessToken;
  Result.BaseUrl     := GraphDriveDefaultBaseUrl;
  Result.DrivePath   := Format(GraphDrivePathFormat, [DriveId]);
  Result.RootAddress := GraphDriveRootAddress;
end;

class function TGraphDriveConfig.ForFolder(const AccessToken, DriveId, FolderId: string): TGraphDriveConfig;
begin
  if DriveId.Trim.IsEmpty then
    raise EGraphDriveValidationError.Create(0, '', SMissingDriveId);

  if FolderId.Trim.IsEmpty then
    raise EGraphDriveValidationError.Create(0, '', SMissingFolderId);

  const IsDotSegment = (FolderId = CurrentSegment) or (FolderId = ParentSegment);
  if IsDotSegment then
    raise EGraphDriveValidationError.Create(0, '', SUnsafeFolderId);

  const EncodedFolderId = TGraphDrivePath.EncodeSegment(FolderId);

  Result.AccessToken := AccessToken;
  Result.BaseUrl     := GraphDriveDefaultBaseUrl;
  Result.DrivePath   := Format(GraphDrivePathFormat, [DriveId]);
  Result.RootAddress := Format(GraphDriveItemAddressFormat, [EncodedFolderId]);
end;

function TGraphDrivePage.HasMore: Boolean;
begin
  Result := not NextLink.IsEmpty;
end;

class function TGraphDriveUploadRequest.ForPath(const Path: string; const Mode: TGraphDriveWriteMode): TGraphDriveUploadRequest;
begin
  Result      := Default(TGraphDriveUploadRequest);
  Result.Path := Path;
  Result.Mode := Mode;
end;

end.
