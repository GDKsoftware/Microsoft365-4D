unit MSGraph.Graph.Drive.Tests;

interface

uses
  DUnitX.TestFramework,
  MSGraph.Graph.Http.Transport.Fake,
  MSGraph.Graph.Drive.Interfaces;

type
  [TestFixture]
  TGraphDriveClientTests = class
  strict private
    FFake: IFakeGraphHttpTransport;
    FDrive: IGraphDriveClient;

    class function FileJson(const Id, Name, ETag: string): string; static;
    class function FolderJson(const Id, Name: string): string; static;
  public
    [Setup]
    procedure Setup;

    [Test]
    procedure GetItem_NestedPath_RequestsEncodedAddressWithBearerToken;
    [Test]
    procedure GetItem_File_MapsItemFields;
    [Test]
    procedure GetItem_EmptyPath_RequestsTheRoot;
    [Test]
    procedure GetItem_ParentReference_MapsParentDriveId;
    [Test]
    procedure GetItem_FolderShortcut_MapsTargetAndKind;
    [Test]
    procedure GetItem_RemoteFile_IsNotAFolderShortcut;
    [Test]
    procedure GetItem_PlainFolder_IsNoShortcut;
    [Test]
    procedure GetItem_FolderShortcut_KeepsTheRawJson;
    [Test]
    procedure GetItem_RemoteItemWithoutParentReference_LeavesRemoteDriveIdEmpty;
    [Test]
    procedure GetItem_RemoteFolderWithoutId_IsNoShortcut;
    [Test]
    procedure GetItem_OwnFacetAndRemoteFacet_OwnFacetWins;
    [Test]
    procedure ForDrive_UsesTheDrivesSegment;
    [Test]
    procedure ForMyDrive_UrlsStayRootRelative;
    [Test]
    procedure ForFolder_BuildsItemUrls;
    [Test]
    procedure ForFolder_Delta_UsesTheFolder;
    [Test]
    procedure ForFolder_EmptyDriveId_RaisesValidationError;
    [Test]
    procedure ForFolder_EmptyFolderId_RaisesValidationError;
    [Test]
    [TestCase('Current', '.')]
    [TestCase('Parent', '..')]
    procedure ForFolder_DotFolderId_RaisesValidationError(const FolderId: string);
    [Test]
    procedure ForFolder_FolderIdWithSpecialCharacters_EncodesItAsOneSegment;
    [Test]
    procedure EmptyRootAddress_FallsBackToTheDriveRoot;
    [Test]
    procedure ListChildren_Folder_MapsItemsAndNextLink;
    [Test]
    procedure ListNext_GraphLink_FollowsTheLink;
    [Test]
    procedure ListNext_ForeignLink_RaisesWithoutSendingTheToken;
    [Test]
    procedure DownloadContent_SendsNoAuthorizationHeader;
    [Test]
    procedure Upload_Replace_PutsBytesWithReplaceBehavior;
    [Test]
    procedure Upload_FailIfExists_UsesFailBehavior;
    [Test]
    procedure Upload_IfMatch_SendsTheETag;
    [Test]
    procedure Upload_IfMatchWithoutETag_RaisesWithoutRequest;
    [Test]
    procedure CreateFolder_PostsNameFolderFacetAndFailBehavior;
    [Test]
    procedure Move_PatchesParentIdAndName;
    [Test]
    procedure Delete_SendsDeleteToTheItemAddress;
    [Test]
    procedure Delta_EmptyLink_StartsAtTheRootAndMapsDeletions;
    [Test]
    procedure Delta_ForeignLink_RaisesWithoutRequest;
    [Test]
    procedure LatestDeltaLink_AsksForTheLatestToken;

    [Test]
    [TestCase('Unauthorized', '401,EGraphDriveAuthenticationError')]
    [TestCase('Forbidden', '403,EGraphDriveAccessDeniedError')]
    [TestCase('NotFound', '404,EGraphDriveNotFoundError')]
    [TestCase('Conflict', '409,EGraphDriveConflictError')]
    [TestCase('Gone', '410,EGraphDriveResyncRequiredError')]
    [TestCase('PreconditionFailed', '412,EGraphDrivePreconditionFailedError')]
    [TestCase('ServerError', '500,EGraphDriveServerError')]
    [TestCase('Unavailable', '503,EGraphDriveRateLimitError')]
    procedure ErrorStatus_RaisesMatchingError(const StatusCode: Integer; const ExpectedClass: string);

    [Test]
    procedure TooManyRequests_ReadsRetryAfter;
    [Test]
    procedure TransportFailure_RaisesNetworkError;
    [Test]
    procedure NonJsonResponse_RaisesMappingError;
  end;

implementation

uses
  System.SysUtils,
  System.JSON,
  System.Net.URLClient,
  MSGraph.Graph.Http.Types,
  MSGraph.Graph.Drive.Types,
  MSGraph.Graph.Drive.Exceptions,
  MSGraph.Graph.Drive;

const
  Token = 'drive-test-token';
  MyDriveUrl = 'https://graph.microsoft.com/v1.0/me/drive';
  DrivesUrl = 'https://graph.microsoft.com/v1.0/drives/D1';
  FolderUrl = DrivesUrl + '/items/F1';
  ETagValue = '"{1A2B},3"';
  FileFormat = '{"id":"%s","name":"%s","eTag":"%s","cTag":"c-%s","size":12,' +
    '"lastModifiedDateTime":"2026-10-14T10:15:30Z","parentReference":{"id":"parent-1","driveId":"d1"},' +
    '"file":{"mimeType":"text/markdown"},"@microsoft.graph.downloadUrl":"https://files.example/%s"}';
  FolderFormat = '{"id":"%s","name":"%s","eTag":"e-%s","folder":{"childCount":1},"parentReference":{"id":"root-1"}}';
  ShortcutJson = '{"id":"s1","name":"Gedeeld","eTag":"e-s1","parentReference":{"id":"root-1","driveId":"mine"},' +
    '"remoteItem":{"id":"r1","parentReference":{"driveId":"owner"},"folder":{"childCount":3}}}';
  RemoteFileJson = '{"id":"s2","name":"Los.md","eTag":"e-s2","parentReference":{"id":"root-1","driveId":"mine"},' +
    '"remoteItem":{"id":"r2","parentReference":{"driveId":"owner"},"file":{"mimeType":"text/markdown"}}}';
  RemoteWithoutParentJson = '{"id":"s3","name":"Gedeeld","eTag":"e-s3","remoteItem":{"id":"r3","folder":{}}}';
  RemoteFolderWithoutIdJson = '{"id":"s4","name":"Gedeeld","eTag":"e-s4","remoteItem":{"parentReference":{"driveId":"owner"},"folder":{}}}';
  OwnFileWithRemoteFolderJson = '{"id":"s5","name":"Los.md","eTag":"e-s5","file":{},"remoteItem":{"id":"r5","folder":{}}}';
  ErrorBody = '{"error":{"code":"someCode","message":"Something went wrong"}}';

class function TGraphDriveClientTests.FileJson(const Id, Name, ETag: string): string;
begin
  const EscapedETag = ETag.Replace('"', '\"');
  Result := Format(FileFormat, [Id, Name, EscapedETag, Id, Id]);
end;

class function TGraphDriveClientTests.FolderJson(const Id, Name: string): string;
begin
  Result := Format(FolderFormat, [Id, Name, Id]);
end;

procedure TGraphDriveClientTests.Setup;
begin
  FFake := TFakeGraphHttpTransport.Create;
  FDrive := TGraphDriveClient.Create(TGraphDriveConfig.ForMyDrive(Token), FFake);
end;

procedure TGraphDriveClientTests.GetItem_NestedPath_RequestsEncodedAddressWithBearerToken;
begin
  FFake.EnqueueResponse(200, FileJson('f1', '2026-10-14 Idee #1.md', ETagValue));

  FDrive.GetItem('Notes/2026-10-14 Idee #1.md');

  Assert.AreEqual('GET', FFake.LastRequest.Method);
  Assert.AreEqual(MyDriveUrl + '/root:/Notes/2026-10-14%20Idee%20%231.md:', FFake.LastRequest.Url);
  Assert.AreEqual('Bearer ' + Token, FFake.HeaderValue(0, 'Authorization'));
end;

procedure TGraphDriveClientTests.GetItem_File_MapsItemFields;
begin
  FFake.EnqueueResponse(200, FileJson('f1', 'Today.md', ETagValue));

  const Item = FDrive.GetItem('Notes/Today.md');

  Assert.AreEqual('f1', Item.Id);
  Assert.AreEqual('Today.md', Item.Name);
  Assert.AreEqual(Ord(TGraphDriveItemKind.FileItem), Ord(Item.Kind));
  Assert.AreEqual(ETagValue, Item.ETag);
  Assert.AreEqual('c-f1', Item.CTag);
  Assert.AreEqual(Int64(12), Item.Size);
  Assert.AreEqual('2026-10-14T10:15:30Z', Item.LastModified);
  Assert.AreEqual('parent-1', Item.ParentId);
  Assert.AreEqual('https://files.example/f1', Item.DownloadUrl);
  Assert.IsFalse(Item.IsDeleted);
  Assert.IsTrue(Item.Raw.Contains('"mimeType"'));
end;

procedure TGraphDriveClientTests.GetItem_EmptyPath_RequestsTheRoot;
begin
  FFake.EnqueueResponse(200, FolderJson('root-1', 'root'));

  const Item = FDrive.GetItem('');

  Assert.AreEqual(MyDriveUrl + '/root', FFake.LastRequest.Url);
  Assert.AreEqual(Ord(TGraphDriveItemKind.FolderItem), Ord(Item.Kind));
end;

procedure TGraphDriveClientTests.GetItem_ParentReference_MapsParentDriveId;
begin
  FFake.EnqueueResponse(200, FileJson('f1', 'Today.md', ETagValue));

  const Item = FDrive.GetItem('Notes/Today.md');

  Assert.AreEqual('d1', Item.ParentDriveId);
  Assert.AreEqual('', Item.RemoteDriveId);
  Assert.AreEqual('', Item.RemoteItemId);
end;

procedure TGraphDriveClientTests.GetItem_FolderShortcut_MapsTargetAndKind;
begin
  FFake.EnqueueResponse(200, ShortcutJson);

  const Item = FDrive.GetItem('Gedeeld');

  Assert.AreEqual('mine', Item.ParentDriveId);
  Assert.AreEqual('owner', Item.RemoteDriveId);
  Assert.AreEqual('r1', Item.RemoteItemId);
  Assert.AreEqual(Ord(TGraphDriveItemKind.FolderItem), Ord(Item.Kind));
  Assert.IsTrue(Item.IsFolderShortcut);
end;

procedure TGraphDriveClientTests.GetItem_RemoteFile_IsNotAFolderShortcut;
begin
  FFake.EnqueueResponse(200, RemoteFileJson);

  const Item = FDrive.GetItem('Los.md');

  Assert.AreEqual('r2', Item.RemoteItemId);
  Assert.AreEqual(Ord(TGraphDriveItemKind.FileItem), Ord(Item.Kind));
  Assert.IsFalse(Item.IsFolderShortcut);
end;

procedure TGraphDriveClientTests.GetItem_PlainFolder_IsNoShortcut;
begin
  FFake.EnqueueResponse(200, FolderJson('d1', 'Notes'));

  const Item = FDrive.GetItem('Notes');

  Assert.IsFalse(Item.IsFolderShortcut);
  Assert.AreEqual('', Item.ParentDriveId);
end;

procedure TGraphDriveClientTests.GetItem_FolderShortcut_KeepsTheRawJson;
begin
  FFake.EnqueueResponse(200, ShortcutJson);

  const Item = FDrive.GetItem('Gedeeld');

  const HasRemoteItem = Item.Raw.Contains('"remoteItem"');
  const HasChildCount = Item.Raw.Contains('"childCount":3');
  Assert.IsTrue(HasRemoteItem);
  Assert.IsTrue(HasChildCount);
end;

procedure TGraphDriveClientTests.GetItem_RemoteItemWithoutParentReference_LeavesRemoteDriveIdEmpty;
begin
  FFake.EnqueueResponse(200, RemoteWithoutParentJson);

  const Item = FDrive.GetItem('Gedeeld');

  Assert.AreEqual('', Item.RemoteDriveId);
  Assert.AreEqual('r3', Item.RemoteItemId);
end;

procedure TGraphDriveClientTests.GetItem_RemoteFolderWithoutId_IsNoShortcut;
begin
  FFake.EnqueueResponse(200, RemoteFolderWithoutIdJson);

  const Item = FDrive.GetItem('Gedeeld');

  Assert.AreEqual('owner', Item.RemoteDriveId);
  Assert.IsFalse(Item.IsFolderShortcut);
end;

procedure TGraphDriveClientTests.GetItem_OwnFacetAndRemoteFacet_OwnFacetWins;
begin
  FFake.EnqueueResponse(200, OwnFileWithRemoteFolderJson);

  const Item = FDrive.GetItem('Los.md');

  Assert.AreEqual(Ord(TGraphDriveItemKind.FileItem), Ord(Item.Kind));
end;

procedure TGraphDriveClientTests.ForDrive_UsesTheDrivesSegment;
begin
  const Drive: IGraphDriveClient = TGraphDriveClient.Create(TGraphDriveConfig.ForDrive(Token, 'b!lib-42'), FFake);
  FFake.EnqueueResponse(200, FolderJson('root-1', 'root'));

  Drive.GetItem('');

  Assert.AreEqual('https://graph.microsoft.com/v1.0/drives/b!lib-42/root', FFake.LastRequest.Url);
end;

procedure TGraphDriveClientTests.ForMyDrive_UrlsStayRootRelative;
begin
  FFake.EnqueueResponse(200, FolderJson('d1', 'Notes'));
  FFake.EnqueueResponse(200, '{"value":[]}');
  FFake.EnqueueResponse(201, FileJson('f1', 'a.md', ETagValue));
  FFake.EnqueueResponse(201, FolderJson('d2', 'New'));
  FFake.EnqueueResponse(204, '');

  FDrive.GetItem('Notes');
  Assert.AreEqual(MyDriveUrl + '/root:/Notes:', FFake.LastRequest.Url);
  FDrive.ListChildren('Notes');
  Assert.AreEqual(MyDriveUrl + '/root:/Notes:/children', FFake.LastRequest.Url);
  FDrive.Upload(TGraphDriveUploadRequest.ForPath('Notes/a.md', TGraphDriveWriteMode.Replace), TEncoding.UTF8.GetBytes('x'));
  Assert.AreEqual(MyDriveUrl + '/root:/Notes/a.md:/content?@microsoft.graph.conflictBehavior=replace', FFake.LastRequest.Url);
  FDrive.CreateFolder('Notes', 'New');
  Assert.AreEqual(MyDriveUrl + '/root:/Notes:/children', FFake.LastRequest.Url);
  FDrive.Delete('Notes/a.md');
  Assert.AreEqual(MyDriveUrl + '/root:/Notes/a.md:', FFake.LastRequest.Url);
end;

procedure TGraphDriveClientTests.ForFolder_BuildsItemUrls;
begin
  const Drive: IGraphDriveClient = TGraphDriveClient.Create(TGraphDriveConfig.ForFolder(Token, 'D1', 'F1'), FFake);
  FFake.EnqueueResponse(200, FolderJson('d1', 'Notes'));
  FFake.EnqueueResponse(200, '{"value":[]}');
  FFake.EnqueueResponse(201, FileJson('f1', 'a.md', ETagValue));
  FFake.EnqueueResponse(201, FolderJson('d2', 'New'));
  FFake.EnqueueResponse(204, '');

  Drive.GetItem('Notes');
  Assert.AreEqual(FolderUrl + ':/Notes:', FFake.LastRequest.Url);
  Drive.ListChildren('');
  Assert.AreEqual(FolderUrl + '/children', FFake.LastRequest.Url);
  Drive.Upload(TGraphDriveUploadRequest.ForPath('Notes/a.md', TGraphDriveWriteMode.Replace), TEncoding.UTF8.GetBytes('x'));
  Assert.AreEqual(FolderUrl + ':/Notes/a.md:/content?@microsoft.graph.conflictBehavior=replace', FFake.LastRequest.Url);
  Drive.CreateFolder('Notes', 'New');
  Assert.AreEqual(FolderUrl + ':/Notes:/children', FFake.LastRequest.Url);
  Drive.Delete('Notes/a.md');
  Assert.AreEqual(FolderUrl + ':/Notes/a.md:', FFake.LastRequest.Url);
end;

procedure TGraphDriveClientTests.ForFolder_Delta_UsesTheFolder;
begin
  const Drive: IGraphDriveClient = TGraphDriveClient.Create(TGraphDriveConfig.ForFolder(Token, 'D1', 'F1'), FFake);
  FFake.EnqueueResponse(200, '{"value":[],"@odata.deltaLink":"' + FolderUrl + '/delta?token=t1"}');
  FFake.EnqueueResponse(200, '{"value":[],"@odata.deltaLink":"' + FolderUrl + '/delta?token=t2"}');

  Drive.Delta('');
  Assert.AreEqual(FolderUrl + '/delta', FFake.LastRequest.Url);
  Drive.LatestDeltaLink;
  Assert.AreEqual(FolderUrl + '/delta?token=latest', FFake.LastRequest.Url);
end;

procedure TGraphDriveClientTests.ForFolder_EmptyDriveId_RaisesValidationError;
begin
  Assert.WillRaiseWithMessage(
    procedure
    begin
      TGraphDriveConfig.ForFolder(Token, '', 'F1');
    end,
    EGraphDriveValidationError,
    'A folder address needs the id of the drive');
end;

procedure TGraphDriveClientTests.ForFolder_EmptyFolderId_RaisesValidationError;
begin
  Assert.WillRaiseWithMessage(
    procedure
    begin
      TGraphDriveConfig.ForFolder(Token, 'D1', '');
    end,
    EGraphDriveValidationError,
    'A folder address needs the id of the folder');
end;

procedure TGraphDriveClientTests.ForFolder_DotFolderId_RaisesValidationError(const FolderId: string);
begin
  Assert.WillRaiseWithMessage(
    procedure
    begin
      TGraphDriveConfig.ForFolder(Token, 'D1', FolderId);
    end,
    EGraphDriveValidationError,
    'A folder address needs a real folder id, not "." or ".."');
end;

procedure TGraphDriveClientTests.ForFolder_FolderIdWithSpecialCharacters_EncodesItAsOneSegment;
begin
  const Drive: IGraphDriveClient = TGraphDriveClient.Create(TGraphDriveConfig.ForFolder(Token, 'D1', 'B7F0!12#3/x'), FFake);
  FFake.EnqueueResponse(200, FolderJson('d1', 'Notes'));

  Drive.GetItem('Notes');

  Assert.AreEqual(DrivesUrl + '/items/B7F0%2112%233%2Fx:/Notes:', FFake.LastRequest.Url);
end;

procedure TGraphDriveClientTests.EmptyRootAddress_FallsBackToTheDriveRoot;
begin
  var Config := Default(TGraphDriveConfig);
  Config.AccessToken := Token;
  Config.BaseUrl     := GraphDriveDefaultBaseUrl;
  Config.DrivePath   := GraphMyDrivePath;
  const Drive: IGraphDriveClient = TGraphDriveClient.Create(Config, FFake);
  FFake.EnqueueResponse(200, FolderJson('root-1', 'root'));
  FFake.EnqueueResponse(200, '{"value":[],"@odata.deltaLink":"' + MyDriveUrl + '/root/delta?token=t1"}');

  Drive.GetItem('');
  Assert.AreEqual(MyDriveUrl + '/root', FFake.LastRequest.Url);
  Drive.Delta('');
  Assert.AreEqual(MyDriveUrl + '/root/delta', FFake.LastRequest.Url);
end;

procedure TGraphDriveClientTests.ListChildren_Folder_MapsItemsAndNextLink;
begin
  const Body = Format('{"value":[%s,%s],"@odata.nextLink":"%s/root:/Notes:/children?$skiptoken=abc"}',
    [FolderJson('d1', 'AI'), FileJson('f1', 'Today.md', ETagValue), MyDriveUrl]);
  FFake.EnqueueResponse(200, Body);

  const Page = FDrive.ListChildren('Notes');

  Assert.AreEqual(MyDriveUrl + '/root:/Notes:/children', FFake.LastRequest.Url);
  Assert.AreEqual(2, Integer(Length(Page.Items)));
  Assert.AreEqual('AI', Page.Items[0].Name);
  Assert.AreEqual(Ord(TGraphDriveItemKind.FolderItem), Ord(Page.Items[0].Kind));
  Assert.AreEqual('Today.md', Page.Items[1].Name);
  Assert.IsTrue(Page.HasMore);
  Assert.AreEqual(MyDriveUrl + '/root:/Notes:/children?$skiptoken=abc', Page.NextLink);
end;

procedure TGraphDriveClientTests.ListNext_GraphLink_FollowsTheLink;
begin
  const Link = MyDriveUrl + '/root:/Notes:/children?$skiptoken=abc';
  FFake.EnqueueResponse(200, '{"value":[]}');

  const Page = FDrive.ListNext(Link);

  Assert.AreEqual(Link, FFake.LastRequest.Url);
  Assert.IsFalse(Page.HasMore);
end;

procedure TGraphDriveClientTests.ListNext_ForeignLink_RaisesWithoutSendingTheToken;
begin
  Assert.WillRaise(
    procedure
    begin
      FDrive.ListNext('https://graph.microsoft.com.evil.example/v1.0/me/drive/root/children');
    end,
    EGraphDriveValidationError);

  Assert.AreEqual(0, FFake.RequestCount);
end;

procedure TGraphDriveClientTests.DownloadContent_SendsNoAuthorizationHeader;
begin
  FFake.EnqueueBinaryResponse(200, TEncoding.UTF8.GetBytes('# Today'));

  const Bytes = FDrive.DownloadContent('https://files.example/f1');

  Assert.AreEqual('# Today', TEncoding.UTF8.GetString(Bytes));
  Assert.AreEqual('https://files.example/f1', FFake.LastRequest.Url);
  Assert.AreEqual('', FFake.HeaderValue(0, 'Authorization'));
end;

procedure TGraphDriveClientTests.Upload_Replace_PutsBytesWithReplaceBehavior;
begin
  FFake.EnqueueResponse(201, FileJson('f1', 'Today.md', ETagValue));

  const Request = TGraphDriveUploadRequest.ForPath('Notes/Today.md', TGraphDriveWriteMode.Replace);
  const Item = FDrive.Upload(Request, TEncoding.UTF8.GetBytes('# Today'));

  Assert.AreEqual(ETagValue, Item.ETag);
  Assert.AreEqual('PUT', FFake.LastRequest.Method);
  Assert.AreEqual(MyDriveUrl + '/root:/Notes/Today.md:/content?@microsoft.graph.conflictBehavior=replace', FFake.LastRequest.Url);
  Assert.AreEqual('# Today', TEncoding.UTF8.GetString(FFake.LastRequest.BodyBytes));
  Assert.AreEqual('application/octet-stream', FFake.HeaderValue(0, 'Content-Type'));
  Assert.AreEqual('', FFake.HeaderValue(0, 'If-Match'));
end;

procedure TGraphDriveClientTests.Upload_FailIfExists_UsesFailBehavior;
begin
  FFake.EnqueueResponse(201, FileJson('f1', 'Today.md', ETagValue));

  const Request = TGraphDriveUploadRequest.ForPath('Notes/Today.md', TGraphDriveWriteMode.FailIfExists);
  FDrive.Upload(Request, TEncoding.UTF8.GetBytes('# Today'));

  Assert.AreEqual(MyDriveUrl + '/root:/Notes/Today.md:/content?@microsoft.graph.conflictBehavior=fail', FFake.LastRequest.Url);
end;

procedure TGraphDriveClientTests.Upload_IfMatch_SendsTheETag;
begin
  FFake.EnqueueResponse(200, FileJson('f1', 'Today.md', '"{1A2B},4"'));

  var Request := TGraphDriveUploadRequest.ForPath('Notes/Today.md', TGraphDriveWriteMode.IfMatch);
  Request.ETag := ETagValue;
  FDrive.Upload(Request, TEncoding.UTF8.GetBytes('# Today'));

  Assert.AreEqual(MyDriveUrl + '/root:/Notes/Today.md:/content', FFake.LastRequest.Url);
  Assert.AreEqual(ETagValue, FFake.HeaderValue(0, 'If-Match'));
end;

procedure TGraphDriveClientTests.Upload_IfMatchWithoutETag_RaisesWithoutRequest;
begin
  Assert.WillRaise(
    procedure
    begin
      const Request = TGraphDriveUploadRequest.ForPath('Notes/Today.md', TGraphDriveWriteMode.IfMatch);
      FDrive.Upload(Request, TEncoding.UTF8.GetBytes('# Today'));
    end,
    EGraphDriveValidationError);

  Assert.AreEqual(0, FFake.RequestCount);
end;

procedure TGraphDriveClientTests.CreateFolder_PostsNameFolderFacetAndFailBehavior;
begin
  FFake.EnqueueResponse(201, FolderJson('d2', 'AI'));

  const Item = FDrive.CreateFolder('Notes', 'AI');

  Assert.AreEqual('d2', Item.Id);
  Assert.AreEqual('POST', FFake.LastRequest.Method);
  Assert.AreEqual(MyDriveUrl + '/root:/Notes:/children', FFake.LastRequest.Url);

  const Body = TJSONObject.ParseJSONValue(FFake.LastRequest.Body) as TJSONObject;
  try
    Assert.AreEqual('AI', Body.GetValue<string>('name'));
    Assert.IsTrue(Body.GetValue('folder') is TJSONObject);
    Assert.AreEqual('fail', (Body.GetValue('@microsoft.graph.conflictBehavior') as TJSONString).Value);
  finally
    Body.Free;
  end;
end;

procedure TGraphDriveClientTests.Move_PatchesParentIdAndName;
begin
  FFake.EnqueueResponse(200, FileJson('f1', 'Renamed.md', ETagValue));

  FDrive.Move('Notes/Today.md', 'd2', 'Renamed.md');

  Assert.AreEqual('PATCH', FFake.LastRequest.Method);
  Assert.AreEqual(MyDriveUrl + '/root:/Notes/Today.md:', FFake.LastRequest.Url);

  const Body = TJSONObject.ParseJSONValue(FFake.LastRequest.Body) as TJSONObject;
  try
    Assert.AreEqual('Renamed.md', Body.GetValue<string>('name'));
    Assert.AreEqual('d2', (Body.GetValue('parentReference') as TJSONObject).GetValue<string>('id'));
  finally
    Body.Free;
  end;
end;

procedure TGraphDriveClientTests.Delete_SendsDeleteToTheItemAddress;
begin
  FFake.EnqueueResponse(204, '');

  FDrive.Delete('Notes/Today.md');

  Assert.AreEqual('DELETE', FFake.LastRequest.Method);
  Assert.AreEqual(MyDriveUrl + '/root:/Notes/Today.md:', FFake.LastRequest.Url);
  Assert.AreEqual('Bearer ' + Token, FFake.HeaderValue(0, 'Authorization'));
end;

procedure TGraphDriveClientTests.Delta_EmptyLink_StartsAtTheRootAndMapsDeletions;
begin
  const Body = Format('{"value":[{"id":"f9","name":"Old.md","deleted":{"state":"deleted"},"parentReference":{"id":"d1"}}],' +
    '"@odata.deltaLink":"%s/root/delta?token=t2"}', [MyDriveUrl]);
  FFake.EnqueueResponse(200, Body);

  const Page = FDrive.Delta('');

  Assert.AreEqual(MyDriveUrl + '/root/delta', FFake.LastRequest.Url);
  Assert.AreEqual(1, Integer(Length(Page.Items)));
  Assert.IsTrue(Page.Items[0].IsDeleted);
  Assert.IsFalse(Page.HasMore);
  Assert.AreEqual(MyDriveUrl + '/root/delta?token=t2', Page.DeltaLink);
end;

procedure TGraphDriveClientTests.Delta_ForeignLink_RaisesWithoutRequest;
begin
  Assert.WillRaise(
    procedure
    begin
      FDrive.Delta('http://graph.microsoft.com/v1.0/me/drive/root/delta?token=t2');
    end,
    EGraphDriveValidationError);

  Assert.AreEqual(0, FFake.RequestCount);
end;

procedure TGraphDriveClientTests.LatestDeltaLink_AsksForTheLatestToken;
begin
  FFake.EnqueueResponse(200, Format('{"value":[],"@odata.deltaLink":"%s/root/delta?token=now"}', [MyDriveUrl]));

  const Link = FDrive.LatestDeltaLink;

  Assert.AreEqual(MyDriveUrl + '/root/delta?token=latest', FFake.LastRequest.Url);
  Assert.AreEqual(MyDriveUrl + '/root/delta?token=now', Link);
end;

procedure TGraphDriveClientTests.ErrorStatus_RaisesMatchingError(const StatusCode: Integer; const ExpectedClass: string);
begin
  FFake.EnqueueResponse(StatusCode, ErrorBody);

  var IsRaised := False;
  try
    FDrive.GetItem('Notes/Today.md');
  except
    on E: EGraphDriveError do
    begin
      Assert.AreEqual(ExpectedClass, E.ClassName);
      Assert.AreEqual(StatusCode, E.StatusCode);
      Assert.AreEqual('someCode', E.ErrorCode);
      Assert.IsFalse(E.Message.Contains(Token));
      IsRaised := True;
    end;
  end;

  Assert.IsTrue(IsRaised);
end;

procedure TGraphDriveClientTests.TooManyRequests_ReadsRetryAfter;
begin
  FFake.EnqueueResponseWithHeaders(429, ErrorBody, [TNetHeader.Create('Retry-After', '12')]);

  var RetryAfter := 0;
  try
    FDrive.ListChildren('Notes');
  except
    on E: EGraphDriveRateLimitError do
      RetryAfter := E.RetryAfterSeconds;
  end;

  Assert.AreEqual(12, RetryAfter);
end;

procedure TGraphDriveClientTests.TransportFailure_RaisesNetworkError;
begin
  FFake.EnqueueTransportFailure('connection refused');

  Assert.WillRaise(
    procedure
    begin
      FDrive.GetItem('Notes/Today.md');
    end,
    EGraphDriveNetworkError);
end;

procedure TGraphDriveClientTests.NonJsonResponse_RaisesMappingError;
begin
  FFake.EnqueueResponse(200, '<html>maintenance</html>');

  Assert.WillRaise(
    procedure
    begin
      FDrive.GetItem('Notes/Today.md');
    end,
    EGraphDriveMappingError);
end;

end.
