unit MSGraph.OAuth2.PKCE.Random.Tests;

interface

uses
  DUnitX.TestFramework,
  System.SysUtils,
  MSGraph.OAuth2.PKCE.Random.Interfaces;

type
  [TestFixture]
  [Category('OAuth2.PKCE')]
  TOAuth2PKCERandomTests = class
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
    procedure GenerateSecureRandomString_KnownBytes_ReturnsTheirBase64Url;

    [Test]
    procedure GenerateNonce_Generated_Returns32HexCharacters;
    [Test]
    procedure GenerateNonce_KnownBytes_ReturnsTheirHex;
  end;
implementation

uses
  MSGraph.OAuth2.PKCE.Random.Stub;

const
  UrlSafeCharacters = ['A'..'Z', 'a'..'z', '0'..'9', '-', '_'];
  HexCharacters     = ['0'..'9', 'A'..'F', 'a'..'f'];

procedure TOAuth2PKCERandomTests.Setup;
begin
  FRandom := TOAuth2PKCERandomStub.Create;
end;

procedure TOAuth2PKCERandomTests.TearDown;
begin
  FRandom := nil;
end;

procedure TOAuth2PKCERandomTests.GenerateSecureRandomString_GivenLength_ReturnsThatLength(const RequestedLength: Integer);
begin
  const Generated = FRandom.GenerateSecureRandomString(RequestedLength);

  Assert.AreEqual(RequestedLength, Generated.Length);
end;

procedure TOAuth2PKCERandomTests.GenerateSecureRandomString_Generated_ContainsOnlyUrlSafeCharacters;
begin
  const Generated = FRandom.GenerateSecureRandomString(128);

  Assert.IsTrue(ContainsOnly(Generated, UrlSafeCharacters), Generated);
end;

procedure TOAuth2PKCERandomTests.GenerateSecureRandomString_KnownBytes_ReturnsTheirBase64Url;
begin
  const Generated = FRandom.GenerateSecureRandomString(8);

  Assert.AreEqual('AAECAwQF', Generated);
end;

procedure TOAuth2PKCERandomTests.GenerateNonce_Generated_Returns32HexCharacters;
begin
  const Nonce = FRandom.GenerateNonce;

  Assert.AreEqual(32, Nonce.Length);
  Assert.IsTrue(ContainsOnly(Nonce, HexCharacters), Nonce);
end;

procedure TOAuth2PKCERandomTests.GenerateNonce_KnownBytes_ReturnsTheirHex;
begin
  const Nonce = FRandom.GenerateNonce;

  Assert.AreEqual('000102030405060708090a0b0c0d0e0f', Nonce);
end;

function TOAuth2PKCERandomTests.ContainsOnly(const Value: string; const Allowed: TSysCharSet): Boolean;
begin
  for var Character in Value do
  begin
    if not CharInSet(Character, Allowed) then
      Exit(False);
  end;

  Result := True;
end;

end.
