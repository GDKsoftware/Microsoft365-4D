unit MSGraph.Graph.Drive;

interface

uses
  System.SysUtils,
  System.Net.URLClient,
  MSGraph.Graph.Http.Types,
  MSGraph.Graph.Http.Interfaces,
  MSGraph.Graph.Drive.Types,
  MSGraph.Graph.Drive.Interfaces;

type
  TGraphDriveClient = class(TInterfacedObject, IGraphDriveClient)
  strict private
    FConfig: TGraphDriveConfig;
    FTransport: IGraphHttpTransport;

    const
      MethodGet = 'GET';
      MethodPost = 'POST';
      MethodPatch = 'PATCH';
      MethodPut = 'PUT';
      MethodDelete = 'DELETE';
      HeaderAuthorization = 'Authorization';
      HeaderContentType = 'Content-Type';
      HeaderIfMatch = 'If-Match';
      HeaderRetryAfter = 'Retry-After';
      BearerFormat = 'Bearer %s';
      ContentTypeJson = 'application/json';
      ContentTypeOctetStream = 'application/octet-stream';
      DeltaSuffix = '/delta';
      LatestDeltaSuffix = '/delta?token=latest';
      AddressFormat = '%s%s';
      ConflictBehaviorReplace = '?@microsoft.graph.conflictBehavior=replace';
      ConflictBehaviorFail = '?@microsoft.graph.conflictBehavior=fail';
      KeyName = 'name';
      KeyFolder = 'folder';
      KeyConflictBehavior = '@microsoft.graph.conflictBehavior';
      ValueFail = 'fail';
      KeyParentReference = 'parentReference';
      KeyId = 'id';

    function RootAddress: string;
    function DriveUrl(const Suffix: string): string;
    function AuthorizationHeader: TNetHeader;
    function JsonHeaders: TArray<TNetHeader>;
    function UploadHeaders(const Request: TGraphDriveUploadRequest): TArray<TNetHeader>;
    function UploadUrl(const Request: TGraphDriveUploadRequest): string;
    function Send(const Request: TGraphHttpRequest): TGraphHttpResponse;
    function SendBinary(const Request: TGraphHttpRequest): TGraphHttpResponse;
    function GetJson(const Url: string): string;
    procedure GuardGraphLink(const Link: string);
    procedure GuardSuccess(const Response: TGraphHttpResponse);
    procedure GuardConfig;

    class function NewRequest(const Method, Url: string; const Headers: TArray<TNetHeader>): TGraphHttpRequest; static;
    class function FolderBody(const Name: string): string; static;
    class function MoveBody(const NewParentId, NewName: string): string; static;
  public
    constructor Create(const Config: TGraphDriveConfig); overload;
    constructor Create(const Config: TGraphDriveConfig; const Transport: IGraphHttpTransport); overload;

    function GetItem(const Path: string): TGraphDriveItem;
    function ListChildren(const Path: string): TGraphDrivePage;
    function ListNext(const NextLink: string): TGraphDrivePage;
    function DownloadContent(const DownloadUrl: string): TBytes;
    function Upload(const Request: TGraphDriveUploadRequest; const Content: TBytes): TGraphDriveItem;
    function CreateFolder(const ParentPath, Name: string): TGraphDriveItem;
    function Move(const Path, NewParentId, NewName: string): TGraphDriveItem;
    procedure Delete(const Path: string);
    function Delta(const Link: string): TGraphDrivePage;
    function LatestDeltaLink: string;
  end;

implementation

uses
  System.JSON,
  MSGraph.Graph.Http.Transport,
  MSGraph.Graph.Drive.Path,
  MSGraph.Graph.Drive.Mapper,
  MSGraph.Graph.Drive.Exceptions;

const
  SMissingAccessToken = 'A drive client needs an access token';
  SMissingDrive = 'A drive client needs a base URL and a drive path';
  SMissingTransport = 'A drive client needs an HTTP transport';
  SForeignLink = 'Refusing to follow a link outside %s';
  SMissingDownloadUrl = 'The item has no download URL';
  SMissingETag = 'An If-Match upload needs the eTag of the item';
  SMissingDeltaLink = 'Microsoft Graph returned no delta link';
  SUnreachable = 'Microsoft Graph could not be reached: %s';
  SUnsupportedMode = 'Unsupported drive write mode: %d';

constructor TGraphDriveClient.Create(const Config: TGraphDriveConfig);
begin
  Create(Config, TNetHttpTransport.Create);
end;

constructor TGraphDriveClient.Create(const Config: TGraphDriveConfig; const Transport: IGraphHttpTransport);
begin
  inherited Create;
  FConfig    := Config;
  FTransport := Transport;
  GuardConfig;
end;

function TGraphDriveClient.GetItem(const Path: string): TGraphDriveItem;
begin
  const Address = TGraphDrivePath.ItemAddress(RootAddress, Path);
  const Url = DriveUrl(Address);
  Result := TGraphDriveMapper.ItemFromJson(GetJson(Url));
end;

function TGraphDriveClient.ListChildren(const Path: string): TGraphDrivePage;
begin
  const Address = TGraphDrivePath.ChildrenAddress(RootAddress, Path);
  const Url = DriveUrl(Address);
  Result := TGraphDriveMapper.PageFromJson(GetJson(Url));
end;

function TGraphDriveClient.ListNext(const NextLink: string): TGraphDrivePage;
begin
  GuardGraphLink(NextLink);
  Result := TGraphDriveMapper.PageFromJson(GetJson(NextLink));
end;

function TGraphDriveClient.DownloadContent(const DownloadUrl: string): TBytes;
begin
  if DownloadUrl.Trim.IsEmpty then
    raise EGraphDriveValidationError.Create(0, '', SMissingDownloadUrl);

  const Request = NewRequest(MethodGet, DownloadUrl, nil);
  const Response = SendBinary(Request);
  GuardSuccess(Response);
  Result := Response.ContentBytes;
end;

function TGraphDriveClient.Upload(const Request: TGraphDriveUploadRequest; const Content: TBytes): TGraphDriveItem;
begin
  var HttpRequest := NewRequest(MethodPut, UploadUrl(Request), UploadHeaders(Request));
  HttpRequest.BodyBytes := Content;

  const Response = Send(HttpRequest);
  GuardSuccess(Response);
  Result := TGraphDriveMapper.ItemFromJson(Response.Content);
end;

function TGraphDriveClient.CreateFolder(const ParentPath, Name: string): TGraphDriveItem;
begin
  const Address = TGraphDrivePath.ChildrenAddress(RootAddress, ParentPath);
  const Url = DriveUrl(Address);
  var HttpRequest := NewRequest(MethodPost, Url, JsonHeaders);
  HttpRequest.Body := FolderBody(Name);

  const Response = Send(HttpRequest);
  GuardSuccess(Response);
  Result := TGraphDriveMapper.ItemFromJson(Response.Content);
end;

function TGraphDriveClient.Move(const Path, NewParentId, NewName: string): TGraphDriveItem;
begin
  const Address = TGraphDrivePath.ItemAddress(RootAddress, Path);
  const Url = DriveUrl(Address);
  var HttpRequest := NewRequest(MethodPatch, Url, JsonHeaders);
  HttpRequest.Body := MoveBody(NewParentId, NewName);

  const Response = Send(HttpRequest);
  GuardSuccess(Response);
  Result := TGraphDriveMapper.ItemFromJson(Response.Content);
end;

procedure TGraphDriveClient.Delete(const Path: string);
begin
  const Address = TGraphDrivePath.ItemAddress(RootAddress, Path);
  const Url = DriveUrl(Address);
  const Response = Send(NewRequest(MethodDelete, Url, [AuthorizationHeader]));
  GuardSuccess(Response);
end;

function TGraphDriveClient.Delta(const Link: string): TGraphDrivePage;
begin
  if Link.IsEmpty then
  begin
    const Root = RootAddress;
    const DeltaAddress = Format(AddressFormat, [Root, DeltaSuffix]);
    const Url = DriveUrl(DeltaAddress);
    const Json = GetJson(Url);
    Result := TGraphDriveMapper.PageFromJson(Json);
    Exit;
  end;

  GuardGraphLink(Link);
  Result := TGraphDriveMapper.PageFromJson(GetJson(Link));
end;

function TGraphDriveClient.LatestDeltaLink: string;
begin
  const Root = RootAddress;
  const LatestAddress = Format(AddressFormat, [Root, LatestDeltaSuffix]);
  const Url = DriveUrl(LatestAddress);
  const Json = GetJson(Url);
  const Page = TGraphDriveMapper.PageFromJson(Json);
  Result := Page.DeltaLink;
  if Result.IsEmpty then
    raise EGraphDriveMappingError.Create(0, '', SMissingDeltaLink);
end;

function TGraphDriveClient.RootAddress: string;
begin
  Result := FConfig.RootAddress;
  if Result.IsEmpty then
    Result := GraphDriveRootAddress;
end;

function TGraphDriveClient.DriveUrl(const Suffix: string): string;
begin
  Result := FConfig.BaseUrl + FConfig.DrivePath + Suffix;
end;

function TGraphDriveClient.AuthorizationHeader: TNetHeader;
begin
  Result := TNetHeader.Create(HeaderAuthorization, Format(BearerFormat, [FConfig.AccessToken]));
end;

function TGraphDriveClient.JsonHeaders: TArray<TNetHeader>;
begin
  Result := [AuthorizationHeader, TNetHeader.Create(HeaderContentType, ContentTypeJson)];
end;

function TGraphDriveClient.UploadHeaders(const Request: TGraphDriveUploadRequest): TArray<TNetHeader>;
begin
  Result := [AuthorizationHeader, TNetHeader.Create(HeaderContentType, ContentTypeOctetStream)];

  const IsConditional = (Request.Mode = TGraphDriveWriteMode.IfMatch);
  if not IsConditional then
    Exit;

  if Request.ETag.Trim.IsEmpty then
    raise EGraphDriveValidationError.Create(0, '', SMissingETag);

  Result := Result + [TNetHeader.Create(HeaderIfMatch, Request.ETag)];
end;

function TGraphDriveClient.UploadUrl(const Request: TGraphDriveUploadRequest): string;
begin
  const Address = TGraphDrivePath.ContentAddress(RootAddress, Request.Path);
  const Url = DriveUrl(Address);

  case Request.Mode of
    TGraphDriveWriteMode.Replace      : Result := Url + ConflictBehaviorReplace;
    TGraphDriveWriteMode.FailIfExists : Result := Url + ConflictBehaviorFail;
    TGraphDriveWriteMode.IfMatch      : Result := Url;
  else
    raise EGraphDriveValidationError.Create(0, '', Format(SUnsupportedMode, [Ord(Request.Mode)]));
  end;
end;

function TGraphDriveClient.Send(const Request: TGraphHttpRequest): TGraphHttpResponse;
begin
  try
    Result := FTransport.Execute(Request);
  except
    on E: ENetException do
      raise EGraphDriveNetworkError.Create(0, '', Format(SUnreachable, [E.Message]));
  end;
end;

function TGraphDriveClient.SendBinary(const Request: TGraphHttpRequest): TGraphHttpResponse;
begin
  try
    Result := FTransport.ExecuteBinary(Request);
  except
    on E: ENetException do
      raise EGraphDriveNetworkError.Create(0, '', Format(SUnreachable, [E.Message]));
  end;
end;

function TGraphDriveClient.GetJson(const Url: string): string;
begin
  const Response = Send(NewRequest(MethodGet, Url, [AuthorizationHeader]));
  GuardSuccess(Response);
  Result := Response.Content;
end;

procedure TGraphDriveClient.GuardGraphLink(const Link: string);
begin
  const Prefix = FConfig.BaseUrl + '/';
  const IsGraphLink = Link.StartsWith(Prefix, True);
  if not IsGraphLink then
    raise EGraphDriveValidationError.Create(0, '', Format(SForeignLink, [FConfig.BaseUrl]));
end;

procedure TGraphDriveClient.GuardSuccess(const Response: TGraphHttpResponse);
begin
  if Response.IsSuccess then
    Exit;

  const RetryAfter = Response.HeaderValue(HeaderRetryAfter);
  TGraphDriveMapper.RaiseFor(Response.StatusCode, Response.Content, RetryAfter);
end;

procedure TGraphDriveClient.GuardConfig;
begin
  if FConfig.AccessToken.Trim.IsEmpty then
    raise EGraphDriveValidationError.Create(0, '', SMissingAccessToken);

  const HasDrive = (not FConfig.BaseUrl.IsEmpty and not FConfig.DrivePath.IsEmpty);
  if not HasDrive then
    raise EGraphDriveValidationError.Create(0, '', SMissingDrive);

  if not Assigned(FTransport) then
    raise EGraphDriveValidationError.Create(0, '', SMissingTransport);
end;

class function TGraphDriveClient.NewRequest(const Method, Url: string; const Headers: TArray<TNetHeader>): TGraphHttpRequest;
begin
  Result         := Default(TGraphHttpRequest);
  Result.Method  := Method;
  Result.Url     := Url;
  Result.Headers := Headers;
end;

class function TGraphDriveClient.FolderBody(const Name: string): string;
begin
  const Body = TJSONObject.Create;
  try
    Body.AddPair(KeyName, Name);
    Body.AddPair(KeyFolder, TJSONObject.Create);
    Body.AddPair(KeyConflictBehavior, ValueFail);
    Result := Body.ToJSON;
  finally
    Body.Free;
  end;
end;

class function TGraphDriveClient.MoveBody(const NewParentId, NewName: string): string;
begin
  const Body = TJSONObject.Create;
  try
    const Parent = TJSONObject.Create;
    Parent.AddPair(KeyId, NewParentId);
    Body.AddPair(KeyParentReference, Parent);
    Body.AddPair(KeyName, NewName);
    Result := Body.ToJSON;
  finally
    Body.Free;
  end;
end;

end.
