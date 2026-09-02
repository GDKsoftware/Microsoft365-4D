unit MSGraph.Graph.Http.Redaction.Tests;

interface

uses
  DUnitX.TestFramework;

type
  [TestFixture]
  TGraphUrlRedactorTests = class
  public
    [Test]
    procedure Redact_UrlWithoutQuery_IsUnchanged;
    [Test]
    procedure Redact_UrlWithoutParameterValue_IsUnchanged;
    [Test]
    procedure Redact_AuthToken_IsMaskedAndOtherParametersAreKept;
    [Test]
    procedure Redact_AuthToken_IsMatchedRegardlessOfCase;
  end;

implementation

uses
  MSGraph.Graph.Http.Redaction;

procedure TGraphUrlRedactorTests.Redact_UrlWithoutQuery_IsUnchanged;
begin
  const Url = 'https://graph.microsoft.com/v1.0/me/messages';

  Assert.AreEqual(Url, TGraphUrlRedactor.Redact(Url));
end;

procedure TGraphUrlRedactorTests.Redact_UrlWithoutParameterValue_IsUnchanged;
begin
  const Url = 'https://outlook.office.com/upload?authtoken';

  Assert.AreEqual(Url, TGraphUrlRedactor.Redact(Url));
end;

procedure TGraphUrlRedactorTests.Redact_AuthToken_IsMaskedAndOtherParametersAreKept;
begin
  const Redacted = TGraphUrlRedactor.Redact(
    'https://outlook.office.com/upload?authtoken=SECRET&sessionid=42');

  Assert.AreEqual('https://outlook.office.com/upload?authtoken=***&sessionid=42', Redacted);
end;

procedure TGraphUrlRedactorTests.Redact_AuthToken_IsMatchedRegardlessOfCase;
begin
  const Redacted = TGraphUrlRedactor.Redact('https://outlook.office.com/upload?AuthToken=SECRET');

  Assert.AreEqual('https://outlook.office.com/upload?AuthToken=***', Redacted);
end;

initialization
  TDUnitX.RegisterTestFixture(TGraphUrlRedactorTests);

end.
