unit MSGraph.OAuth2.PKCE.Random.Windows.Tests;

interface

uses
  DUnitX.TestFramework,
  System.SysUtils,
  MSGraph.OAuth2.PKCE.Random.Interfaces;

type
  [TestFixture]
  [Category('OAuth2.PKCE')]
  TOAuth2PKCERandomWindowsTests = class
  private
    FRandom: IOAuth2PKCERandom;

    function ContainsOnly(const Value: string; const Allowed: TSysCharSet): Boolean;

  public
    [Setup]
    procedure Setup;
    [TearDown]
    procedure TearDown;

    [Test]
    [TestCase('Zero', '0')]
    [TestCase('Minimum', '43')]
    [TestCase('Maximum', '128')]
    procedure GenerateSecureRandomString_GivenLength_ReturnsThatLength(const RequestedLength: Integer);

    [Test]
    procedure GenerateSecureRandomString_Generated_ContainsOnlyUrlSafeCharacters;
    [Test]
    procedure GenerateSecureRandomString_TwoCalls_ReturnDifferentValues;

    [Test]
    procedure GenerateCodeChallenge_Rfc7636Example_ReturnsExpectedChallenge;
  end;

implementation

uses
  MSGraph.OAuth2.PKCE.Random.Windows;

const
  UrlSafeCharacters = ['A'..'Z', 'a'..'z', '0'..'9', '-', '_'];

procedure TOAuth2PKCERandomWindowsTests.Setup;
begin
  FRandom := TOAuth2PKCERandomWindows.Create;
end;

procedure TOAuth2PKCERandomWindowsTests.TearDown;
begin
  FRandom := nil;
end;

procedure TOAuth2PKCERandomWindowsTests.GenerateSecureRandomString_GivenLength_ReturnsThatLength(const RequestedLength: Integer);
begin
  const Generated = FRandom.GenerateSecureRandomString(RequestedLength);

  Assert.AreEqual(RequestedLength, Generated.Length);
end;

procedure TOAuth2PKCERandomWindowsTests.GenerateSecureRandomString_Generated_ContainsOnlyUrlSafeCharacters;
begin
  const Generated = FRandom.GenerateSecureRandomString(128);

  Assert.IsTrue(ContainsOnly(Generated, UrlSafeCharacters), Generated);
end;

procedure TOAuth2PKCERandomWindowsTests.GenerateSecureRandomString_TwoCalls_ReturnDifferentValues;
begin
  const First  = FRandom.GenerateSecureRandomString(43);
  const Second = FRandom.GenerateSecureRandomString(43);

  Assert.AreNotEqual(First, Second);
end;

procedure TOAuth2PKCERandomWindowsTests.GenerateCodeChallenge_Rfc7636Example_ReturnsExpectedChallenge;
begin
  const CodeVerifier = 'dBjftJeZ4CVP-mB92K27uhbUJU1p1r_wW1gFWFOEjXk';

  const Challenge = FRandom.GenerateCodeChallenge(CodeVerifier);

  Assert.AreEqual('E9Melhoa2OwvFrEMTJguCHaoeK1t8URWbuGJSstw-cM', Challenge);
end;

function TOAuth2PKCERandomWindowsTests.ContainsOnly(const Value: string; const Allowed: TSysCharSet): Boolean;
begin
  for var Character in Value do
  begin
    if not CharInSet(Character, Allowed) then
      Exit(False);
  end;

  Result := True;
end;

end.
