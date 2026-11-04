unit MSGraph.OAuth2.PKCE.Random.Tests;

interface

uses
  DUnitX.TestFramework,
  MSGraph.OAuth2.PKCE.Random.Interfaces;

type
  [TestFixture]
  [Category('OAuth2.PKCE')]
  TOAuth2PKCERandomTests = class
  private
    FRandom: IOAuth2PKCERandom;

  public
    [Setup]
    procedure Setup;
    [TearDown]
    procedure TearDown;

    [Test]
    [TestCase('Empty', '|', '|')]
    [TestCase('OneBytePadding', '66|Zg', '|')]
    [TestCase('TwoBytePadding', '666F|Zm8', '|')]
    [TestCase('NoPadding', '666F6F|Zm9v', '|')]
    [TestCase('PlusAndSlash', 'FBFFBF|-_-_', '|')]
    [TestCase('PlusSlashAndPadding', 'FBFF|-_8', '|')]
    procedure Base64UrlEncode_KnownBytes_ReturnsUrlSafeUnpadded(const HexData, Expected: string);

    [Test]
    procedure Base64UrlEncode_LongData_ContainsNoLineBreaks;

    [Test]
    procedure GenerateCodeChallenge_Rfc7636Example_ReturnsExpectedChallenge;
    [Test]
    procedure GenerateCodeChallenge_SameVerifier_ReturnsSameChallenge;
    [Test]
    procedure GenerateCodeChallenge_AnyVerifier_Returns43UrlSafeCharacters;

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
    procedure GenerateNonce_Generated_Returns32HexCharacters;
    [Test]
    procedure GenerateNonce_TwoCalls_ReturnDifferentValues;
  end;

implementation

uses
  System.SysUtils,
  System.Classes,
  MSGraph.OAuth2.PKCE.Random;

const
  UrlSafeCharacters = ['A'..'Z', 'a'..'z', '0'..'9', '-', '_'];
  HexCharacters     = ['0'..'9', 'A'..'F', 'a'..'f'];

function HexToBytes(const Hex: string): TBytes;
begin
  SetLength(Result, Hex.Length div 2);
  HexToBin(PChar(Hex), 0, Result, 0, Length(Result));
end;

function ContainsOnly(const Value: string; const Allowed: TSysCharSet): Boolean;
begin
  for var Character in Value do
    if not CharInSet(Character, Allowed) then
      Exit(False);

  Result := True;
end;

procedure TOAuth2PKCERandomTests.Setup;
begin
  FRandom := TOAuth2PKCERandom.Create;
end;

procedure TOAuth2PKCERandomTests.TearDown;
begin
  FRandom := nil;
end;

procedure TOAuth2PKCERandomTests.Base64UrlEncode_KnownBytes_ReturnsUrlSafeUnpadded(const HexData, Expected: string);
begin
  const Data = HexToBytes(HexData);

  const Encoded = FRandom.Base64UrlEncode(Data);

  Assert.AreEqual(Expected, Encoded);
end;

procedure TOAuth2PKCERandomTests.Base64UrlEncode_LongData_ContainsNoLineBreaks;
begin
  var Data: TBytes;
  SetLength(Data, 96);
  FillChar(Data[0], Length(Data), $AB);

  const Encoded = FRandom.Base64UrlEncode(Data);

  Assert.AreEqual(128, Encoded.Length);
  Assert.IsTrue(ContainsOnly(Encoded, UrlSafeCharacters), Encoded);
end;

procedure TOAuth2PKCERandomTests.GenerateCodeChallenge_Rfc7636Example_ReturnsExpectedChallenge;
begin
  const CodeVerifier = 'dBjftJeZ4CVP-mB92K27uhbUJU1p1r_wW1gFWFOEjXk';

  const Challenge = FRandom.GenerateCodeChallenge(CodeVerifier);

  Assert.AreEqual('E9Melhoa2OwvFrEMTJguCHaoeK1t8URWbuGJSstw-cM', Challenge);
end;

procedure TOAuth2PKCERandomTests.GenerateCodeChallenge_SameVerifier_ReturnsSameChallenge;
begin
  const CodeVerifier = FRandom.GenerateSecureRandomString(64);

  const First  = FRandom.GenerateCodeChallenge(CodeVerifier);
  const Second = FRandom.GenerateCodeChallenge(CodeVerifier);

  Assert.AreEqual(First, Second);
end;

procedure TOAuth2PKCERandomTests.GenerateCodeChallenge_AnyVerifier_Returns43UrlSafeCharacters;
begin
  const CodeVerifier = FRandom.GenerateSecureRandomString(128);

  const Challenge = FRandom.GenerateCodeChallenge(CodeVerifier);

  Assert.AreEqual(43, Challenge.Length);
  Assert.IsTrue(ContainsOnly(Challenge, UrlSafeCharacters), Challenge);
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

procedure TOAuth2PKCERandomTests.GenerateSecureRandomString_TwoCalls_ReturnDifferentValues;
begin
  const First  = FRandom.GenerateSecureRandomString(43);
  const Second = FRandom.GenerateSecureRandomString(43);

  Assert.AreNotEqual(First, Second);
end;

procedure TOAuth2PKCERandomTests.GenerateNonce_Generated_Returns32HexCharacters;
begin
  const Nonce = FRandom.GenerateNonce;

  Assert.AreEqual(32, Nonce.Length);
  Assert.IsTrue(ContainsOnly(Nonce, HexCharacters), Nonce);
end;

procedure TOAuth2PKCERandomTests.GenerateNonce_TwoCalls_ReturnDifferentValues;
begin
  const First  = FRandom.GenerateNonce;
  const Second = FRandom.GenerateNonce;

  Assert.AreNotEqual(First, Second);
end;

initialization
  TDUnitX.RegisterTestFixture(TOAuth2PKCERandomTests);

end.
