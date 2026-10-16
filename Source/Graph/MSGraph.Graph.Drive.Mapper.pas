unit MSGraph.Graph.Drive.Mapper;

interface

uses
  System.JSON,
  MSGraph.Graph.Drive.Types;

type
  TGraphDriveMapper = class
  strict private
    const
      KeyValue = 'value';
      KeyNextLink = '@odata.nextLink';
      KeyDeltaLink = '@odata.deltaLink';
      KeyId = 'id';
      KeyName = 'name';
      KeyETag = 'eTag';
      KeyCTag = 'cTag';
      KeySize = 'size';
      KeyLastModified = 'lastModifiedDateTime';
      KeyParentReference = 'parentReference';
      KeyDriveId = 'driveId';
      KeyRemoteItem = 'remoteItem';
      KeyFile = 'file';
      KeyFolder = 'folder';
      KeyDeleted = 'deleted';
      KeyDownloadUrl = '@microsoft.graph.downloadUrl';
      KeyError = 'error';
      KeyCode = 'code';
      KeyMessage = 'message';
      DefaultRetryAfterSeconds = 5;

    class function ParseObject(const Json: string): TJSONObject; static;
    class function KeyString(const Obj: TJSONObject; const Key: string): string; static;
    class function KindOf(const Obj, RemoteItem: TJSONObject): TGraphDriveItemKind; static;
    class function HasFacet(const Obj: TJSONObject; const Key: string): Boolean; static;
    class function RetryAfterSeconds(const HeaderValue: string): Integer; static;
    class procedure ReadError(const Body: string; out ErrorCode, ErrorMessage: string); static;
  public
    class function ItemFromObject(const Obj: TJSONObject): TGraphDriveItem; static;
    class function ItemFromJson(const Json: string): TGraphDriveItem; static;
    class function PageFromJson(const Json: string): TGraphDrivePage; static;
    class procedure RaiseFor(const StatusCode: Integer; const Body, RetryAfterHeader: string); static;
  end;

implementation

uses
  System.SysUtils,
  MSGraph.Graph.JsonHelper,
  MSGraph.Graph.Drive.Exceptions;

const
  SUnexpectedJson = 'Microsoft Graph returned a response that is not a JSON object';
  SDriveHttpError = 'Microsoft Graph drive request failed with HTTP %d (%s): %s';
  SNoErrorDetails = 'no details';

class function TGraphDriveMapper.ItemFromObject(const Obj: TJSONObject): TGraphDriveItem;
begin
  Result := Default(TGraphDriveItem);

  const Parent = TGraphJson.GetObject(Obj, KeyParentReference);
  const RemoteItem = TGraphJson.GetObject(Obj, KeyRemoteItem);
  const RemoteParent = TGraphJson.GetObject(RemoteItem, KeyParentReference);

  Result.Id           := TGraphJson.GetString(Obj, KeyId);
  Result.Name         := TGraphJson.GetString(Obj, KeyName);
  Result.Kind         := KindOf(Obj, RemoteItem);
  Result.ETag         := TGraphJson.GetString(Obj, KeyETag);
  Result.CTag         := TGraphJson.GetString(Obj, KeyCTag);
  Result.Size         := TGraphJson.GetInt64(Obj, KeySize);
  Result.LastModified := TGraphJson.GetString(Obj, KeyLastModified);
  Result.IsDeleted    := Assigned(Obj.GetValue(KeyDeleted));
  Result.DownloadUrl  := KeyString(Obj, KeyDownloadUrl);
  Result.Raw          := Obj.ToJSON;

  Result.ParentId      := TGraphJson.GetString(Parent, KeyId);
  Result.ParentDriveId := TGraphJson.GetString(Parent, KeyDriveId);

  Result.RemoteDriveId  := TGraphJson.GetString(RemoteParent, KeyDriveId);
  Result.RemoteItemId   := TGraphJson.GetString(RemoteItem, KeyId);
  Result.RemoteIsFolder := HasFacet(RemoteItem, KeyFolder);
end;

class function TGraphDriveMapper.ItemFromJson(const Json: string): TGraphDriveItem;
begin
  const Obj = ParseObject(Json);
  try
    Result := ItemFromObject(Obj);
  finally
    Obj.Free;
  end;
end;

class function TGraphDriveMapper.PageFromJson(const Json: string): TGraphDrivePage;
begin
  Result := Default(TGraphDrivePage);

  const Obj = ParseObject(Json);
  try
    Result.NextLink  := KeyString(Obj, KeyNextLink);
    Result.DeltaLink := KeyString(Obj, KeyDeltaLink);

    const Values = TGraphJson.GetArray(Obj, KeyValue);
    if not Assigned(Values) then
      Exit;

    for var Element in Values do
    begin
      if not (Element is TJSONObject) then
        raise EGraphDriveMappingError.Create(0, '', SUnexpectedJson);

      Result.Items := Result.Items + [ItemFromObject(TJSONObject(Element))];
    end;
  finally
    Obj.Free;
  end;
end;

class procedure TGraphDriveMapper.RaiseFor(const StatusCode: Integer; const Body, RetryAfterHeader: string);
begin
  var ErrorCode: string;
  var ErrorMessage: string;
  ReadError(Body, ErrorCode, ErrorMessage);
  const Msg = Format(SDriveHttpError, [StatusCode, ErrorCode, ErrorMessage]);

  case StatusCode of
    401:
      raise EGraphDriveAuthenticationError.Create(StatusCode, ErrorCode, Msg);
    403:
      raise EGraphDriveAccessDeniedError.Create(StatusCode, ErrorCode, Msg);
    404:
      raise EGraphDriveNotFoundError.Create(StatusCode, ErrorCode, Msg);
    409:
      raise EGraphDriveConflictError.Create(StatusCode, ErrorCode, Msg);
    410:
      raise EGraphDriveResyncRequiredError.Create(StatusCode, ErrorCode, Msg);
    412:
      raise EGraphDrivePreconditionFailedError.Create(StatusCode, ErrorCode, Msg);
    429, 503:
      raise EGraphDriveRateLimitError.Create(StatusCode, ErrorCode, Msg, RetryAfterSeconds(RetryAfterHeader));
    500..502, 504..599:
      raise EGraphDriveServerError.Create(StatusCode, ErrorCode, Msg);
  else
    raise EGraphDriveError.Create(StatusCode, ErrorCode, Msg);
  end;
end;

class function TGraphDriveMapper.ParseObject(const Json: string): TJSONObject;
begin
  const Parsed = TJSONObject.ParseJSONValue(Json);
  if not (Parsed is TJSONObject) then
  begin
    Parsed.Free;
    raise EGraphDriveMappingError.Create(0, '', SUnexpectedJson);
  end;

  Result := TJSONObject(Parsed);
end;

class function TGraphDriveMapper.KeyString(const Obj: TJSONObject; const Key: string): string;
begin
  const Value = Obj.GetValue(Key);
  if Value is TJSONString then
    Result := TJSONString(Value).Value
  else
    Result := '';
end;

class function TGraphDriveMapper.HasFacet(const Obj: TJSONObject; const Key: string): Boolean;
begin
  Result := Assigned(Obj) and Assigned(Obj.GetValue(Key));
end;

class function TGraphDriveMapper.KindOf(const Obj, RemoteItem: TJSONObject): TGraphDriveItemKind;
begin
  if HasFacet(Obj, KeyFolder) then
    Exit(TGraphDriveItemKind.FolderItem);

  if HasFacet(Obj, KeyFile) then
    Exit(TGraphDriveItemKind.FileItem);

  if HasFacet(RemoteItem, KeyFolder) then
    Exit(TGraphDriveItemKind.FolderItem);

  if HasFacet(RemoteItem, KeyFile) then
    Exit(TGraphDriveItemKind.FileItem);

  Result := TGraphDriveItemKind.Unknown;
end;

class function TGraphDriveMapper.RetryAfterSeconds(const HeaderValue: string): Integer;
begin
  const IsNumber = TryStrToInt(HeaderValue.Trim, Result);
  const IsUsable = (IsNumber and (Result > 0));
  if not IsUsable then
    Result := DefaultRetryAfterSeconds;
end;

class procedure TGraphDriveMapper.ReadError(const Body: string; out ErrorCode, ErrorMessage: string);
begin
  ErrorCode    := '';
  ErrorMessage := SNoErrorDetails;

  const Parsed = TJSONObject.ParseJSONValue(Body);
  try
    if not (Parsed is TJSONObject) then
      Exit;

    const Error = TGraphJson.GetObject(TJSONObject(Parsed), KeyError);
    ErrorCode := TGraphJson.GetString(Error, KeyCode);

    const Message = TGraphJson.GetString(Error, KeyMessage);
    if not Message.IsEmpty then
      ErrorMessage := Message;
  finally
    Parsed.Free;
  end;
end;

end.
