unit MSGraph.Graph.Http.Redaction;

interface

type
  TGraphUrlRedactor = class
  strict private
    const
      RedactedValue = '***';
      SensitiveQueryParameters: array[0..0] of string = ('authtoken');

    class function RedactParameter(const Parameter: string): string; static;
    class function IsSensitive(const Name: string): Boolean; static;
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
    Exit(Url);

  const BaseUrl = Url.Substring(0, QueryStart + 1);
  const Parameters = Url.Substring(QueryStart + 1).Split(['&']);

  var RedactedParameters: TArray<string>;
  for var Parameter in Parameters do
  begin
    RedactedParameters := RedactedParameters + [RedactParameter(Parameter)];
  end;

  const RedactedQuery = string.Join('&', RedactedParameters);

  Result := Format('%s%s', [BaseUrl, RedactedQuery]);
end;

class function TGraphUrlRedactor.RedactParameter(const Parameter: string): string;
begin
  const Separator = Parameter.IndexOf('=');
  const HasValue = (Separator >= 0);
  if not HasValue then
    Exit(Parameter);

  const Name = Parameter.Substring(0, Separator);
  if not IsSensitive(Name) then
    Exit(Parameter);

  Result := Format('%s=%s', [Name, RedactedValue]);
end;

class function TGraphUrlRedactor.IsSensitive(const Name: string): Boolean;
begin
  Result := False;

  for var SensitiveParameter in SensitiveQueryParameters do
  begin
    if SameText(Name, SensitiveParameter) then
      Exit(True);
  end;
end;

end.
