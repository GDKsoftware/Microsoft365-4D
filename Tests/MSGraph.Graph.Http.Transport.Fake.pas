unit MSGraph.Graph.Http.Transport.Fake;

interface

uses
  System.SysUtils,
  System.Net.URLClient,
  MSGraph.Graph.Http.Types,
  MSGraph.Graph.Http.Interfaces;

type
  IFakeGraphHttpTransport = interface(IGraphHttpTransport)
    ['{DDE14D52-A795-4C90-A1E4-7F2CCAB873CD}']
    procedure EnqueueResponse(const StatusCode: Integer; const Content: string);
    procedure EnqueueResponseWithHeaders(const StatusCode: Integer; const Content: string;
      const Headers: TArray<TNetHeader>);
    procedure EnqueueBinaryResponse(const StatusCode: Integer; const ContentBytes: TBytes);

    function RequestCount: Integer;
    function RequestAt(const Index: Integer): TGraphHttpRequest;
    function LastRequest: TGraphHttpRequest;
    function HeaderValue(const RequestIndex: Integer; const Name: string): string;
  end;

  TFakeGraphHttpTransport = class(TInterfacedObject, IFakeGraphHttpTransport)
  strict private
    FRequests: TArray<TGraphHttpRequest>;
    FResponses: TArray<TGraphHttpResponse>;
    FNextResponse: Integer;

    function TakeNextResponse(const Request: TGraphHttpRequest): TGraphHttpResponse;
  public
    procedure EnqueueResponse(const StatusCode: Integer; const Content: string);
    procedure EnqueueResponseWithHeaders(const StatusCode: Integer; const Content: string;
      const Headers: TArray<TNetHeader>);
    procedure EnqueueBinaryResponse(const StatusCode: Integer; const ContentBytes: TBytes);

    function RequestCount: Integer;
    function RequestAt(const Index: Integer): TGraphHttpRequest;
    function LastRequest: TGraphHttpRequest;
    function HeaderValue(const RequestIndex: Integer; const Name: string): string;

    function Execute(const Request: TGraphHttpRequest): TGraphHttpResponse;
    function ExecuteBinary(const Request: TGraphHttpRequest): TGraphHttpResponse;
  end;

implementation

procedure TFakeGraphHttpTransport.EnqueueResponse(const StatusCode: Integer; const Content: string);
begin
  EnqueueResponseWithHeaders(StatusCode, Content, nil);
end;

procedure TFakeGraphHttpTransport.EnqueueResponseWithHeaders(const StatusCode: Integer;
  const Content: string; const Headers: TArray<TNetHeader>);
begin
  var Response := Default(TGraphHttpResponse);
  Response.StatusCode := StatusCode;
  Response.Content := Content;
  Response.Headers := Headers;

  FResponses := FResponses + [Response];
end;

procedure TFakeGraphHttpTransport.EnqueueBinaryResponse(const StatusCode: Integer; const ContentBytes: TBytes);
begin
  var Response := Default(TGraphHttpResponse);
  Response.StatusCode := StatusCode;
  Response.ContentBytes := ContentBytes;

  FResponses := FResponses + [Response];
end;

function TFakeGraphHttpTransport.RequestCount: Integer;
begin
  Result := Length(FRequests);
end;

function TFakeGraphHttpTransport.RequestAt(const Index: Integer): TGraphHttpRequest;
begin
  if (Index < 0) or (Index > High(FRequests)) then
    raise ERangeError.CreateFmt('Fake transport: no request at index %d (%d recorded)',
      [Index, Length(FRequests)]);

  Result := FRequests[Index];
end;

function TFakeGraphHttpTransport.LastRequest: TGraphHttpRequest;
begin
  Result := RequestAt(High(FRequests));
end;

function TFakeGraphHttpTransport.HeaderValue(const RequestIndex: Integer; const Name: string): string;
begin
  Result := '';
  for var Header in RequestAt(RequestIndex).Headers do
  begin
    if SameText(Header.Name, Name) then
    begin
      Result := Header.Value;
      Exit;
    end;
  end;
end;

function TFakeGraphHttpTransport.TakeNextResponse(const Request: TGraphHttpRequest): TGraphHttpResponse;
begin
  FRequests := FRequests + [Request];

  if FNextResponse > High(FResponses) then
    raise Exception.CreateFmt('Fake transport: unexpected %s %s (no response queued)',
      [Request.Method, Request.Url]);

  Result := FResponses[FNextResponse];
  Inc(FNextResponse);
end;

function TFakeGraphHttpTransport.Execute(const Request: TGraphHttpRequest): TGraphHttpResponse;
begin
  Result := TakeNextResponse(Request);
end;

function TFakeGraphHttpTransport.ExecuteBinary(const Request: TGraphHttpRequest): TGraphHttpResponse;
begin
  Result := TakeNextResponse(Request);
end;

end.
