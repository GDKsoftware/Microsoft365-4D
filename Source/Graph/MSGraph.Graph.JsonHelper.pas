unit MSGraph.Graph.JsonHelper;

interface

uses
  System.JSON;

type
  TGraphJson = class
  public
    class function GetString(const Obj: TJSONObject; const Path: string): string;
    class function GetInt(const Obj: TJSONObject; const Path: string; const Default: Integer = 0): Integer;
    class function GetInt64(const Obj: TJSONObject; const Path: string; const Default: Int64 = 0): Int64;
    class function GetBool(const Obj: TJSONObject; const Path: string; const Default: Boolean = False): Boolean;

    class function GetObject(const Obj: TJSONObject; const Path: string): TJSONObject;
    class function GetArray(const Obj: TJSONObject; const Path: string): TJSONArray;
    class function ArrayItem(const Arr: TJSONArray; const Index: Integer): TJSONObject;

    class function HasError(const Obj: TJSONObject): Boolean;
    class function GetErrorMessage(const Obj: TJSONObject): string;
    class function HasNextPage(const Obj: TJSONObject): Boolean;
  end;

implementation

uses
  System.SysUtils,
  System.Generics.Collections;

class function TGraphJson.GetString(const Obj: TJSONObject; const Path: string): string;
begin
  Result := '';
  if not Assigned(Obj) then
    Exit;

  var Value := Obj.FindValue(Path);
  if Assigned(Value) and (Value is TJSONString) then
    Result := TJSONString(Value).Value;
end;

class function TGraphJson.GetInt(const Obj: TJSONObject; const Path: string; const Default: Integer): Integer;
begin
  Result := Default;
  if not Assigned(Obj) then
    Exit;

  var Value := Obj.FindValue(Path);
  if Assigned(Value) and (Value is TJSONNumber) then
    Result := TJSONNumber(Value).AsInt;
end;

class function TGraphJson.GetInt64(const Obj: TJSONObject; const Path: string; const Default: Int64): Int64;
begin
  Result := Default;
  if not Assigned(Obj) then
    Exit;

  var Value := Obj.FindValue(Path);
  if Assigned(Value) and (Value is TJSONNumber) then
    Result := TJSONNumber(Value).AsInt64;
end;

class function TGraphJson.GetBool(const Obj: TJSONObject; const Path: string; const Default: Boolean): Boolean;
begin
  Result := Default;
  if not Assigned(Obj) then
    Exit;

  var Value := Obj.FindValue(Path);
  if Assigned(Value) then
    Result := (Value is TJSONTrue);
end;

class function TGraphJson.GetObject(const Obj: TJSONObject; const Path: string): TJSONObject;
begin
  Result := nil;
  if not Assigned(Obj) then
    Exit;

  var Value := Obj.FindValue(Path);
  if Assigned(Value) and (Value is TJSONObject) then
    Result := TJSONObject(Value);
end;

class function TGraphJson.GetArray(const Obj: TJSONObject; const Path: string): TJSONArray;
begin
  Result := nil;
  if not Assigned(Obj) then
    Exit;

  var Value := Obj.FindValue(Path);
  if Assigned(Value) and (Value is TJSONArray) then
    Result := TJSONArray(Value);
end;

class function TGraphJson.ArrayItem(const Arr: TJSONArray; const Index: Integer): TJSONObject;
begin
  Result := nil;
  if not Assigned(Arr) then
    Exit;

  if (Index < 0) or (Index >= Arr.Count) then
    Exit;

  if Arr.Items[Index] is TJSONObject then
    Result := TJSONObject(Arr.Items[Index]);
end;

class function TGraphJson.HasError(const Obj: TJSONObject): Boolean;
begin
  Result := Assigned(Obj) and Assigned(Obj.FindValue('error'));
end;

class function TGraphJson.GetErrorMessage(const Obj: TJSONObject): string;
begin
  Result := '';
  if not Assigned(Obj) then
    Exit;

  var ErrorValue := Obj.FindValue('error');
  if not Assigned(ErrorValue) then
    Exit;

  if ErrorValue is TJSONString then
  begin
    Result := TJSONString(ErrorValue).Value;
    Exit;
  end;

  if not (ErrorValue is TJSONObject) then
    Exit;

  var MessageValue := TJSONObject(ErrorValue).FindValue('message');
  if Assigned(MessageValue) and (MessageValue is TJSONString) then
    Result := TJSONString(MessageValue).Value
  else
    Result := ErrorValue.ToString;
end;

class function TGraphJson.HasNextPage(const Obj: TJSONObject): Boolean;
begin
  Result := Assigned(Obj) and Assigned(Obj.FindValue('@odata.nextLink'));
end;

end.
