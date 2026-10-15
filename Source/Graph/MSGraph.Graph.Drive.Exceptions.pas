unit MSGraph.Graph.Drive.Exceptions;

interface

uses
  MSGraph.OAuth2.Types;

type
  /// <summary>
  /// Base class of every drive-client error. <c>StatusCode</c> is 0 when no
  /// HTTP response was received; <c>ErrorCode</c> is Graph's <c>error.code</c>
  /// (for example <c>itemNotFound</c> or <c>nameAlreadyExists</c>) when the
  /// response carried one.
  /// </summary>
  EGraphDriveError = class(EGraphApiException)
  strict private
    FStatusCode: Integer;
    FErrorCode: string;
  public
    constructor Create(const StatusCode: Integer; const ErrorCode, Msg: string); reintroduce;

    property StatusCode: Integer read FStatusCode;
    property ErrorCode: string read FErrorCode;
  end;

  /// Invalid caller input or an unsafe link, detected before a request is sent.
  EGraphDriveValidationError = class(EGraphDriveError);
  /// Transport failure reaching Graph (DNS, refused connection, timeout).
  EGraphDriveNetworkError = class(EGraphDriveError);
  /// A response body could not be parsed into the expected JSON shape.
  EGraphDriveMappingError = class(EGraphDriveError);
  /// HTTP 401 - the access token is missing, invalid or expired.
  EGraphDriveAuthenticationError = class(EGraphDriveError);
  /// HTTP 403 - the token lacks the scope or the account lacks access.
  EGraphDriveAccessDeniedError = class(EGraphDriveError);
  /// HTTP 404 - the item does not exist.
  EGraphDriveNotFoundError = class(EGraphDriveError);
  /// HTTP 409 - an item with that name already exists.
  EGraphDriveConflictError = class(EGraphDriveError);
  /// HTTP 410 - the delta link expired; start again with a full listing.
  EGraphDriveResyncRequiredError = class(EGraphDriveError);
  /// HTTP 412 - the item changed since the eTag sent in If-Match.
  EGraphDrivePreconditionFailedError = class(EGraphDriveError);
  /// HTTP 5xx other than 503 - Graph failed to handle the request.
  EGraphDriveServerError = class(EGraphDriveError);

  /// HTTP 429 or 503 - the caller is being throttled.
  EGraphDriveRateLimitError = class(EGraphDriveError)
  strict private
    FRetryAfterSeconds: Integer;
  public
    constructor Create(const StatusCode: Integer; const ErrorCode, Msg: string; const RetryAfterSeconds: Integer); reintroduce;

    property RetryAfterSeconds: Integer read FRetryAfterSeconds;
  end;

implementation

constructor EGraphDriveError.Create(const StatusCode: Integer; const ErrorCode, Msg: string);
begin
  inherited Create(Msg);
  FStatusCode := StatusCode;
  FErrorCode  := ErrorCode;
end;

constructor EGraphDriveRateLimitError.Create(const StatusCode: Integer; const ErrorCode, Msg: string;
  const RetryAfterSeconds: Integer);
begin
  inherited Create(StatusCode, ErrorCode, Msg);
  FRetryAfterSeconds := RetryAfterSeconds;
end;

end.
