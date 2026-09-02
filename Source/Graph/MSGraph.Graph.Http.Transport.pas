unit MSGraph.Graph.Http.Transport;

interface

uses
  System.Net.HttpClient,
  MSGraph.Graph.Http.Types,
  MSGraph.Graph.Http.Interfaces;

type
  TNetHttpTransport = class(TInterfacedObject, IGraphHttpTransport)
  strict private
    function Send(const HttpClient: THTTPClient; const Request: TGraphHttpRequest): IHTTPResponse;
    function SendPost(const HttpClient: THTTPClient; const Request: TGraphHttpRequest): IHTTPResponse;
    function SendPatch(const HttpClient: THTTPClient; const Request: TGraphHttpRequest): IHTTPResponse;
    function SendPut(const HttpClient: THTTPClient; const Request: TGraphHttpRequest): IHTTPResponse;
    function BuildResponse(const Source: IHTTPResponse): TGraphHttpResponse;
  public
    function Execute(const Request: TGraphHttpRequest): TGraphHttpResponse;
    function ExecuteBinary(const Request: TGraphHttpRequest): TGraphHttpResponse;
  end;

implementation

uses
  System.Classes,
  System.SysUtils,
  MSGraph.OAuth2.Types;

const
  MethodGet = 'GET';
  MethodPost = 'POST';
  MethodPatch = 'PATCH';
  MethodPut = 'PUT';
  MethodDelete = 'DELETE';
  UploadConnectionTimeout = 60000;
  UploadResponseTimeout = 120000;

function TNetHttpTransport.Execute(const Request: TGraphHttpRequest): TGraphHttpResponse;
begin
  var HttpClient := THTTPClient.Create;
  try
    const Response = Send(HttpClient, Request);

    Result := BuildResponse(Response);
    Result.Content := Response.ContentAsString(TEncoding.UTF8);
  finally
    HttpClient.Free;
  end;
end;

function TNetHttpTransport.ExecuteBinary(const Request: TGraphHttpRequest): TGraphHttpResponse;
begin
  var HttpClient := THTTPClient.Create;
  try
    var ResponseStream := TBytesStream.Create;
    try
      const Response = HttpClient.Get(Request.Url, ResponseStream, Request.Headers);

      Result := BuildResponse(Response);
      Result.ContentBytes := ResponseStream.Bytes;
      SetLength(Result.ContentBytes, ResponseStream.Size);
    finally
      ResponseStream.Free;
    end;
  finally
    HttpClient.Free;
  end;
end;

function TNetHttpTransport.Send(const HttpClient: THTTPClient; const Request: TGraphHttpRequest): IHTTPResponse;
begin
  if Request.Method = MethodGet then
    Result := HttpClient.Get(Request.Url, nil, Request.Headers)
  else if Request.Method = MethodPost then
    Result := SendPost(HttpClient, Request)
  else if Request.Method = MethodPatch then
    Result := SendPatch(HttpClient, Request)
  else if Request.Method = MethodPut then
    Result := SendPut(HttpClient, Request)
  else if Request.Method = MethodDelete then
    Result := HttpClient.Delete(Request.Url, nil, Request.Headers)
  else
    raise EGraphApiException.CreateFmt('Unsupported HTTP method: %s', [Request.Method]);
end;

function TNetHttpTransport.SendPost(const HttpClient: THTTPClient; const Request: TGraphHttpRequest): IHTTPResponse;
begin
  var Content: TStringStream := nil;
  if not Request.Body.IsEmpty then
    Content := TStringStream.Create(Request.Body, TEncoding.UTF8);
  try
    Result := HttpClient.Post(Request.Url, Content, nil, Request.Headers);
  finally
    Content.Free;
  end;
end;

function TNetHttpTransport.SendPatch(const HttpClient: THTTPClient; const Request: TGraphHttpRequest): IHTTPResponse;
begin
  var Content := TStringStream.Create(Request.Body, TEncoding.UTF8);
  try
    Result := HttpClient.Patch(Request.Url, Content, nil, Request.Headers);
  finally
    Content.Free;
  end;
end;

function TNetHttpTransport.SendPut(const HttpClient: THTTPClient; const Request: TGraphHttpRequest): IHTTPResponse;
begin
  HttpClient.ConnectionTimeout := UploadConnectionTimeout;
  HttpClient.ResponseTimeout   := UploadResponseTimeout;

  var Content := TBytesStream.Create(Request.BodyBytes);
  try
    Result := HttpClient.Put(Request.Url, Content, nil, Request.Headers);
  finally
    Content.Free;
  end;
end;

function TNetHttpTransport.BuildResponse(const Source: IHTTPResponse): TGraphHttpResponse;
begin
  Result := Default(TGraphHttpResponse);

  Result.StatusCode := Source.StatusCode;
  Result.Headers    := Source.Headers;
end;

end.
