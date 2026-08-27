unit MSGraph.Graph.Http.Redaction;

interface

type
  TGraphUrlRedactor = class
  strict private
    const
      RedactedValue = '***';
      QueryParameterSeparator = '&';
      SensitiveQueryParameters: array[0..0] of string = ('authtoken');

    class function RedactParameter(const Parameter: string): string; static;
    class function IsSensitive(const ParameterName: string): Boolean; static;

  public
    class function Redact(const Url: string): string; static;
  end;

implementation

uses
  System.SysUtils;

class function TGraphUrlRedactor.Redact(const Url: string): string;
begin
  const QueryStart = Url.IndexOf('?');
  const HasQuery = (QueryStart >= 0);
  if not HasQuery then
  begin
    Result := Url;
    Exit;
  end;

  const BaseUrl = Url.Substring(0, QueryStart + 1);
  const Parameters = Url.Substring(QueryStart + 1).Split([QueryParameterSeparator]);

  var RedactedParameters: TArray<string>;
  for var Parameter in Parameters do
  begin
    RedactedParameters := RedactedParameters + [RedactParameter(Parameter)];
  end;

  const RedactedQuery = string.Join(QueryParameterSeparator, RedactedParameters);

  Result := Format('%s%s', [BaseUrl, RedactedQuery]);
end;

class function TGraphUrlRedactor.RedactParameter(const Parameter: string): string;
begin
  const Separator = Parameter.IndexOf('=');
  const HasValue = (Separator >= 0);
  if not HasValue then
  begin
    Result := Parameter;
    Exit;
  end;

  const Name = Parameter.Substring(0, Separator);
  if not IsSensitive(Name) then
  begin
    Result := Parameter;
    Exit;
  end;

  Result := Format('%s=%s', [Name, RedactedValue]);
end;

class function TGraphUrlRedactor.IsSensitive(const ParameterName: string): Boolean;
begin
  Result := False;

  for var SensitiveParameter in SensitiveQueryParameters do
  begin
    if SameText(ParameterName, SensitiveParameter) then
    begin
      Result := True;
      Exit;
    end;
  end;
end;

end.
