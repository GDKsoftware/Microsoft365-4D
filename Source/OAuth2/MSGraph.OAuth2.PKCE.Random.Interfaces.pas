unit MSGraph.OAuth2.PKCE.Random.Interfaces;

interface

uses
  System.SysUtils;

type
  IOAuth2PKCERandom = interface
    ['{59862DAE-DFA5-4F9F-9598-9BFD824F55E9}']

    function Base64UrlEncode(const Data: TBytes): string;
    function GenerateSecureRandomString(const Length: Integer): string;
    function GenerateCodeChallenge(const CodeVerifier: string): string;
    function GenerateNonce: string;
  end;

implementation

end.
