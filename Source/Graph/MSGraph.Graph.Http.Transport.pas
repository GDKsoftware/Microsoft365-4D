unit MSGraph.Graph.Http.Transport;

interface

uses
  MSGraph.Graph.Http.Types,
  MSGraph.Graph.Http.Interfaces;

type
  TNetHttpTransport = class(TInterfacedObject, IGraphHttpTransport)
  public
    function Execute(const Request: TGraphHttpRequest): TGraphHttpResponse;
    function ExecuteBinary(const Request: TGraphHttpRequest): TGraphHttpResponse;
  end;

implementation

uses
  System.Classes,
  System.SysUtils,
  System.Net.HttpClient,
  MSGraph.OAuth2.Types;

const
  MethodGet = 'GET';
  MethodPost = 'POST';
  MethodPatch = 'PATCH';
  MethodDelete = 'DELETE';

function TNetHttpTransport.Execute(const Request: TGraphHttpRequest): TGraphHttpResponse;
begin
  Result := Default(TGraphHttpResponse);

  var HttpClient := THTTPClient.Create;
  try
    var Response: IHTTPResponse;

    if Request.Method = MethodGet then
      Response := HttpClient.Get(Request.Url, nil, Request.Headers)
    else if Request.Method = MethodPost then
    begin
      var Content: TStringStream := nil;
      if not Request.Body.IsEmpty then
        Content := TStringStream.Create(Request.Body, TEncoding.UTF8);
      try
        Response := HttpClient.Post(Request.Url, Content, nil, Request.Headers);
      finally
        Content.Free;
      end;
    end
    else if Request.Method = MethodPatch then
    begin
      var Content := TStringStream.Create(Request.Body, TEncoding.UTF8);
      try
        Response := HttpClient.Patch(Request.Url, Content, nil, Request.Headers);
      finally
        Content.Free;
      end;
    end
    else if Request.Method = MethodDelete then
      Response := HttpClient.Delete(Request.Url, nil, Request.Headers)
    else
      raise EGraphApiException.Create('Unsupported HTTP method: ' + Request.Method);

    Result.StatusCode := Response.StatusCode;
    Result.Content := Response.ContentAsString(TEncoding.UTF8);
  finally
    HttpClient.Free;
  end;
end;

function TNetHttpTransport.ExecuteBinary(const Request: TGraphHttpRequest): TGraphHttpResponse;
begin
  Result := Default(TGraphHttpResponse);

  var HttpClient := THTTPClient.Create;
  try
    var ResponseStream := TBytesStream.Create;
    try
      var Response := HttpClient.Get(Request.Url, ResponseStream, Request.Headers);

      Result.StatusCode := Response.StatusCode;
      Result.ContentBytes := ResponseStream.Bytes;
      SetLength(Result.ContentBytes, ResponseStream.Size);
    finally
      ResponseStream.Free;
    end;
  finally
    HttpClient.Free;
  end;
end;

end.
