unit MSGraph.Graph.Drive.Path;

interface

type
  /// <summary>
  /// Builds the path-based item addresses of the drive API. A drive-relative
  /// path such as <c>Notes/2026-10-14 Idee #1.md</c> becomes
  /// <c>/root:/Notes/2026-10-14%20Idee%20%231.md:</c>; an empty path is
  /// <c>/root</c>.
  /// </summary>
  TGraphDrivePath = class
  strict private
    const
      Separator = '/';
      ParentSegment = '..';
      CurrentSegment = '.';
      PathAddressFormat = '%s:/%s:';
      ChildrenSuffix = '/children';
      PathChildrenFormat = '%s:/%s:/children';
      ContentFormat = '%s:/%s:/content';
      PercentFormat = '%%%.2X';

    class function IsUnreserved(const Value: Byte): Boolean; static;
  public
    class function Normalize(const Path: string): string; static;
    class function Encode(const Path: string): string; static;
    class function EncodeSegment(const Segment: string): string; static;
    class function ItemAddress(const RootAddress, Path: string): string; static;
    class function ChildrenAddress(const RootAddress, Path: string): string; static;
    class function ContentAddress(const RootAddress, Path: string): string; static;
  end;

implementation

uses
  System.SysUtils,
  MSGraph.Graph.Drive.Exceptions;

const
  SEmptyContentPath = 'A drive upload needs the path of a file';
  SUnsafeDrivePath = 'Drive path "%s" contains an empty, "." or ".." segment';

class function TGraphDrivePath.Normalize(const Path: string): string;
begin
  Result := Path.Trim([Separator]);
  if Result.IsEmpty then
    Exit;

  const Segments = Result.Split([Separator]);
  for var Segment in Segments do
  begin
    const IsUnsafe = (Segment.IsEmpty or (Segment = CurrentSegment) or (Segment = ParentSegment));
    if IsUnsafe then
      raise EGraphDriveValidationError.Create(0, '', Format(SUnsafeDrivePath, [Path]));
  end;
end;

class function TGraphDrivePath.Encode(const Path: string): string;
begin
  const Normalized = Normalize(Path);
  if Normalized.IsEmpty then
    Exit('');

  const Segments = Normalized.Split([Separator]);
  var Encoded: TArray<string>;
  SetLength(Encoded, Length(Segments));
  for var Index := 0 to High(Segments) do
  begin
    Encoded[Index] := EncodeSegment(Segments[Index]);
  end;

  Result := string.Join(Separator, Encoded);
end;

class function TGraphDrivePath.ItemAddress(const RootAddress, Path: string): string;
begin
  const Encoded = Encode(Path);
  if Encoded.IsEmpty then
    Result := RootAddress
  else
    Result := Format(PathAddressFormat, [RootAddress, Encoded]);
end;

class function TGraphDrivePath.ChildrenAddress(const RootAddress, Path: string): string;
begin
  const Encoded = Encode(Path);
  if Encoded.IsEmpty then
    Result := RootAddress + ChildrenSuffix
  else
    Result := Format(PathChildrenFormat, [RootAddress, Encoded]);
end;

class function TGraphDrivePath.ContentAddress(const RootAddress, Path: string): string;
begin
  const Encoded = Encode(Path);
  if Encoded.IsEmpty then
    raise EGraphDriveValidationError.Create(0, '', SEmptyContentPath);

  Result := Format(ContentFormat, [RootAddress, Encoded]);
end;

class function TGraphDrivePath.EncodeSegment(const Segment: string): string;
begin
  const Bytes = TEncoding.UTF8.GetBytes(Segment);
  const Builder = TStringBuilder.Create;
  try
    for var Value in Bytes do
    begin
      if IsUnreserved(Value) then
        Builder.Append(Char(Value))
      else
        Builder.AppendFormat(PercentFormat, [Value]);
    end;

    Result := Builder.ToString;
  finally
    Builder.Free;
  end;
end;

class function TGraphDrivePath.IsUnreserved(const Value: Byte): Boolean;
begin
  Result := CharInSet(Char(Value), ['A'..'Z', 'a'..'z', '0'..'9', '-', '.', '_', '~']);
end;

end.
