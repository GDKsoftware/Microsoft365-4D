unit MSGraph.Graph.SharePoint;

interface

uses
  System.SysUtils,
  System.JSON,
  MSGraph.OAuth2.Types,
  MSGraph.Graph.Http,
  MSGraph.Graph.SharePoint.Types,
  MSGraph.Graph.SharePoint.Interfaces;

type
  TSharePointClient = class(TInterfacedObject, ISharePointClient)
  strict private
    FGraphClient: TGraphHttpClient;
    FOwnsClient: Boolean;

    function SiteEndpoint(const SiteId: string): string;

    class function ParseSite(const SiteObj: TJSONObject): TSite; static;
    class function ParseDriveItem(const ItemObj: TJSONObject): TDriveItem; static;
    class function ParseSitePage(const PageObj: TJSONObject): TSitePage; static;
    class function ExtractCanvasHtml(const LayoutObj: TJSONObject): string; static;
    class procedure SortDriveItems(var Items: TArray<TDriveItem>; const OrderBy: string); static;

    const
      EndpointSites = '/sites';
  public
    constructor Create(const AccessToken: string; const LogProc: TLogProc = nil); overload;
    constructor Create(const GraphClient: TGraphHttpClient; const OwnsClient: Boolean = False); overload;
    destructor Destroy; override;

    function ListSites(const Query: string; const Top: Integer = 25): TArray<TSite>;
    function GetSite(const SiteId: string): TSite;
    function ListDriveItems(const SiteId: string; const FolderId: string;
      const Top: Integer = 50; const OrderBy: string = ''): TArray<TDriveItem>;
    function SearchDriveItems(const SiteId: string; const Query: string;
      const Top: Integer = 25; const OrderBy: string = ''): TArray<TDriveItem>;
    function GetDriveItemContent(const SiteId: string; const ItemId: string): TDriveItem;
    function DownloadDriveItemBytes(const SiteId: string; const Item: TDriveItem): TBytes;

    function ListSitePages(const SiteId: string; const Top: Integer = 25): TArray<TSitePage>;
    function GetSitePageContent(const SiteId: string; const PageId: string): TSitePage;

    property GraphClient: TGraphHttpClient read FGraphClient;
  end;

implementation

uses
  System.NetEncoding,
  System.Generics.Collections,
  System.Generics.Defaults,
  MSGraph.Graph.JsonHelper;

constructor TSharePointClient.Create(const AccessToken: string; const LogProc: TLogProc);
begin
  inherited Create;
  FGraphClient := TGraphHttpClient.Create(AccessToken, LogProc);
  FOwnsClient := True;
end;

constructor TSharePointClient.Create(const GraphClient: TGraphHttpClient; const OwnsClient: Boolean);
begin
  inherited Create;
  FGraphClient := GraphClient;
  FOwnsClient := OwnsClient;
end;

destructor TSharePointClient.Destroy;
begin
  if FOwnsClient then
    FGraphClient.Free;
  inherited;
end;

function TSharePointClient.SiteEndpoint(const SiteId: string): string;
begin
  Result := EndpointSites + '/' + TNetEncoding.URL.Encode(SiteId);
end;

class function TSharePointClient.ParseSite(const SiteObj: TJSONObject): TSite;
begin
  Result := Default(TSite);
  if not Assigned(SiteObj) then
    Exit;
  Result.Id := TGraphJson.GetString(SiteObj, 'id');
  Result.Name := TGraphJson.GetString(SiteObj, 'name');
  Result.DisplayName := TGraphJson.GetString(SiteObj, 'displayName');
  Result.WebUrl := TGraphJson.GetString(SiteObj, 'webUrl');
  Result.Description := TGraphJson.GetString(SiteObj, 'description');
  Result.CreatedDateTime := TGraphJson.GetString(SiteObj, 'createdDateTime');
  Result.LastModifiedDateTime := TGraphJson.GetString(SiteObj, 'lastModifiedDateTime');
end;

class function TSharePointClient.ParseDriveItem(const ItemObj: TJSONObject): TDriveItem;
begin
  Result := Default(TDriveItem);
  if not Assigned(ItemObj) then
    Exit;

  Result.Id := TGraphJson.GetString(ItemObj, 'id');
  Result.Name := TGraphJson.GetString(ItemObj, 'name');
  Result.Size := TGraphJson.GetInt64(ItemObj, 'size');
  Result.WebUrl := TGraphJson.GetString(ItemObj, 'webUrl');
  Result.CreatedDateTime := TGraphJson.GetString(ItemObj, 'createdDateTime');
  Result.LastModifiedDateTime := TGraphJson.GetString(ItemObj, 'lastModifiedDateTime');
  Result.DownloadUrl := TGraphJson.GetString(ItemObj, '@microsoft.graph.downloadUrl');

  var FileObj := TGraphJson.GetObject(ItemObj, 'file');
  if Assigned(FileObj) then
  begin
    Result.ItemType := 'file';
    Result.MimeType := TGraphJson.GetString(FileObj, 'mimeType');
  end;

  var FolderObj := TGraphJson.GetObject(ItemObj, 'folder');
  if Assigned(FolderObj) then
  begin
    Result.ItemType := 'folder';
    Result.ChildCount := TGraphJson.GetInt(FolderObj, 'childCount');
  end;

  var ParentRef := TGraphJson.GetObject(ItemObj, 'parentReference');
  if Assigned(ParentRef) then
    Result.ParentPath := TGraphJson.GetString(ParentRef, 'path');
end;

function TSharePointClient.ListSites(const Query: string; const Top: Integer): TArray<TSite>;
begin
  Result := nil;

  var ActualTop := Top;
  if ActualTop < 1 then
    ActualTop := 25
  else if ActualTop > 100 then
    ActualTop := 100;

  var SearchTerm := Query.Trim;
  if SearchTerm.IsEmpty then
    SearchTerm := '*';

  var QueryParams := Format('search=%s&$top=%d&$select=id,name,displayName,webUrl,description,createdDateTime,lastModifiedDateTime',
    [TNetEncoding.URL.Encode(SearchTerm), ActualTop]);

  var Response := FGraphClient.Get(EndpointSites, QueryParams);
  try
    if TGraphJson.HasError(Response) then
      raise EGraphApiException.Create(TGraphJson.GetErrorMessage(Response));

    var ValueArray := TGraphJson.GetArray(Response, 'value');
    if not Assigned(ValueArray) then
      Exit;

    SetLength(Result, ValueArray.Count);
    for var Index := 0 to ValueArray.Count - 1 do
      Result[Index] := ParseSite(TGraphJson.ArrayItem(ValueArray, Index));
  finally
    Response.Free;
  end;
end;

function TSharePointClient.GetSite(const SiteId: string): TSite;
begin
  var Response := FGraphClient.Get(SiteEndpoint(SiteId),
    '$select=id,name,displayName,webUrl,description,createdDateTime,lastModifiedDateTime,root,siteCollection');
  try
    if TGraphJson.HasError(Response) then
      raise EGraphApiException.Create(TGraphJson.GetErrorMessage(Response));

    Result := ParseSite(Response);
  finally
    Response.Free;
  end;
end;

function TSharePointClient.ListDriveItems(const SiteId: string; const FolderId: string;
  const Top: Integer; const OrderBy: string): TArray<TDriveItem>;
begin
  Result := nil;

  var ActualTop := Top;
  if ActualTop < 1 then
    ActualTop := 50
  else if ActualTop > 200 then
    ActualTop := 200;

  var Endpoint := '';
  if FolderId.Trim.IsEmpty then
    Endpoint := SiteEndpoint(SiteId) + '/drive/root/children'
  else
    Endpoint := SiteEndpoint(SiteId) + '/drive/items/' +
      TNetEncoding.URL.Encode(FolderId) + '/children';

  var QueryParams := Format('$top=%d&$select=id,name,size,webUrl,createdDateTime,lastModifiedDateTime,file,folder,parentReference',
    [ActualTop]);

  if not OrderBy.Trim.IsEmpty then
    QueryParams := QueryParams + '&$orderby=' + TNetEncoding.URL.Encode(OrderBy);

  var Response := FGraphClient.Get(Endpoint, QueryParams);
  try
    if TGraphJson.HasError(Response) then
      raise EGraphApiException.Create(TGraphJson.GetErrorMessage(Response));

    var ValueArray := TGraphJson.GetArray(Response, 'value');
    if not Assigned(ValueArray) then
      Exit;

    SetLength(Result, ValueArray.Count);
    for var Index := 0 to ValueArray.Count - 1 do
      Result[Index] := ParseDriveItem(TGraphJson.ArrayItem(ValueArray, Index));
  finally
    Response.Free;
  end;
end;

function TSharePointClient.SearchDriveItems(const SiteId: string; const Query: string;
  const Top: Integer; const OrderBy: string): TArray<TDriveItem>;
begin
  Result := nil;

  var ActualTop := Top;
  if ActualTop < 1 then
    ActualTop := 25
  else if ActualTop > 200 then
    ActualTop := 200;

  var Endpoint := SiteEndpoint(SiteId) +
    '/drive/root/search(q=''' + TNetEncoding.URL.Encode(Query) + ''')';

  var QueryParams := Format('$top=%d&$select=id,name,size,webUrl,createdDateTime,lastModifiedDateTime,file,folder,parentReference',
    [ActualTop]);

  var Response := FGraphClient.Get(Endpoint, QueryParams);
  try
    if TGraphJson.HasError(Response) then
      raise EGraphApiException.Create(TGraphJson.GetErrorMessage(Response));

    var ValueArray := TGraphJson.GetArray(Response, 'value');
    if not Assigned(ValueArray) then
      Exit;

    SetLength(Result, ValueArray.Count);
    for var Index := 0 to ValueArray.Count - 1 do
      Result[Index] := ParseDriveItem(TGraphJson.ArrayItem(ValueArray, Index));

    if not OrderBy.Trim.IsEmpty then
      SortDriveItems(Result, OrderBy);
  finally
    Response.Free;
  end;
end;

class procedure TSharePointClient.SortDriveItems(var Items: TArray<TDriveItem>; const OrderBy: string);
begin
  if Length(Items) <= 1 then
    Exit;

  var Field := OrderBy.Trim.ToLower;
  var Descending := False;

  if Field.EndsWith(' desc') then
  begin
    Descending := True;
    Field := Field.Substring(0, Field.Length - 5).Trim;
  end
  else if Field.EndsWith(' asc') then
    Field := Field.Substring(0, Field.Length - 4).Trim;

  TArray.Sort<TDriveItem>(Items, TComparer<TDriveItem>.Construct(
    function(const Left, Right: TDriveItem): Integer
    begin
      if Field = 'name' then
        Result := CompareText(Left.Name, Right.Name)
      else if Field = 'lastmodifieddatetime' then
        Result := CompareText(Left.LastModifiedDateTime, Right.LastModifiedDateTime)
      else if Field = 'createddatetime' then
        Result := CompareText(Left.CreatedDateTime, Right.CreatedDateTime)
      else if Field = 'size' then
      begin
        if Left.Size < Right.Size then
          Result := -1
        else if Left.Size > Right.Size then
          Result := 1
        else
          Result := 0;
      end
      else
        Result := 0;

      if Descending then
        Result := -Result;
    end
  ));
end;

function TSharePointClient.GetDriveItemContent(const SiteId: string; const ItemId: string): TDriveItem;
begin
  var Endpoint := SiteEndpoint(SiteId) + '/drive/items/' + TNetEncoding.URL.Encode(ItemId);

  var Response := FGraphClient.Get(Endpoint,
    '$select=id,name,size,webUrl,createdDateTime,lastModifiedDateTime,file,folder,parentReference');
  try
    if TGraphJson.HasError(Response) then
      raise EGraphApiException.Create(TGraphJson.GetErrorMessage(Response));

    Result := ParseDriveItem(Response);
  finally
    Response.Free;
  end;
end;

function TSharePointClient.DownloadDriveItemBytes(const SiteId: string; const Item: TDriveItem): TBytes;
begin
  var Endpoint := SiteEndpoint(SiteId) + '/drive/items/' + TNetEncoding.URL.Encode(Item.Id) + '/content';
  Result := FGraphClient.GetRawBytes(Endpoint);
end;

{ TSharePointClient - Site Pages }

class function TSharePointClient.ParseSitePage(const PageObj: TJSONObject): TSitePage;
begin
  Result := Default(TSitePage);
  if not Assigned(PageObj) then
    Exit;

  Result.Id := TGraphJson.GetString(PageObj, 'id');
  Result.Title := TGraphJson.GetString(PageObj, 'title');
  Result.Name := TGraphJson.GetString(PageObj, 'name');
  Result.WebUrl := TGraphJson.GetString(PageObj, 'webUrl');
  Result.Description := TGraphJson.GetString(PageObj, 'description');
  Result.CreatedDateTime := TGraphJson.GetString(PageObj, 'createdDateTime');
  Result.LastModifiedDateTime := TGraphJson.GetString(PageObj, 'lastModifiedDateTime');
end;

class function TSharePointClient.ExtractCanvasHtml(const LayoutObj: TJSONObject): string;
begin
  Result := '';
  if not Assigned(LayoutObj) then
    Exit;

  var Sections := TGraphJson.GetArray(LayoutObj, 'horizontalSections');
  if not Assigned(Sections) then
    Exit;

  var Builder := TStringBuilder.Create;
  try
    for var SectionIndex := 0 to Sections.Count - 1 do
    begin
      var Section := TGraphJson.ArrayItem(Sections, SectionIndex);
      if not Assigned(Section) then
        Continue;

      var Columns := TGraphJson.GetArray(Section, 'columns');
      if not Assigned(Columns) then
        Continue;

      for var ColumnIndex := 0 to Columns.Count - 1 do
      begin
        var Column := TGraphJson.ArrayItem(Columns, ColumnIndex);
        if not Assigned(Column) then
          Continue;

        var Webparts := TGraphJson.GetArray(Column, 'webparts');
        if not Assigned(Webparts) then
          Continue;

        for var WebpartIndex := 0 to Webparts.Count - 1 do
        begin
          var Webpart := TGraphJson.ArrayItem(Webparts, WebpartIndex);
          if not Assigned(Webpart) then
            Continue;

          var InnerHtml := TGraphJson.GetString(Webpart, 'innerHtml');
          if not InnerHtml.IsEmpty then
          begin
            if Builder.Length > 0 then
              Builder.AppendLine;
            Builder.Append(InnerHtml);
          end;
        end;
      end;
    end;

    Result := Builder.ToString;
  finally
    Builder.Free;
  end;
end;

function TSharePointClient.ListSitePages(const SiteId: string; const Top: Integer): TArray<TSitePage>;
begin
  Result := nil;

  var ActualTop := Top;
  if ActualTop < 1 then
    ActualTop := 25
  else if ActualTop > 100 then
    ActualTop := 100;

  var Endpoint := SiteEndpoint(SiteId) + '/pages';
  var QueryParams := Format('$top=%d&$select=id,title,name,webUrl,description,createdDateTime,lastModifiedDateTime',
    [ActualTop]);

  var Response := FGraphClient.Get(Endpoint, QueryParams);
  try
    if TGraphJson.HasError(Response) then
      raise EGraphApiException.Create(TGraphJson.GetErrorMessage(Response));

    var ValueArray := TGraphJson.GetArray(Response, 'value');
    if not Assigned(ValueArray) then
      Exit;

    SetLength(Result, ValueArray.Count);
    for var Index := 0 to ValueArray.Count - 1 do
      Result[Index] := ParseSitePage(TGraphJson.ArrayItem(ValueArray, Index));
  finally
    Response.Free;
  end;
end;

function TSharePointClient.GetSitePageContent(const SiteId: string; const PageId: string): TSitePage;
begin
  var Endpoint := SiteEndpoint(SiteId) + '/pages/' + TNetEncoding.URL.Encode(PageId) +
    '/microsoft.graph.sitePage';

  var Response := FGraphClient.Get(Endpoint,
    '$select=id,title,name,webUrl,description,createdDateTime,lastModifiedDateTime&$expand=canvasLayout');
  try
    if TGraphJson.HasError(Response) then
      raise EGraphApiException.Create(TGraphJson.GetErrorMessage(Response));

    Result := ParseSitePage(Response);

    var CanvasLayout := TGraphJson.GetObject(Response, 'canvasLayout');
    Result.ContentHtml := ExtractCanvasHtml(CanvasLayout);
  finally
    Response.Free;
  end;
end;

end.
