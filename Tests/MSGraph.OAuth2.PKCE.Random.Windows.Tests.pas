unit MSGraph.OAuth2.PKCE.Random.Windows.Tests;

interface

uses
  DUnitX.TestFramework,
  System.SysUtils,
  MSGraph.OAuth2.PKCE.Random.Interfaces;

const
  MissingLibraryName       = 'MSGraph.Missing.dll';
  LibraryWithoutBCryptName = 'kernel32.dll';

type
  [TestFixture]
  [Category('OAuth2.PKCE')]
  TOAuth2PKCERandomWindowsTests = class
  private
    FRandom: IOAuth2PKCERandom;

    function ContainsOnly(const Value: string; const Allowed: TSysCharSet): Boolean;
    procedure PlantLibrary(const PlantedPath: string);

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

    [Test]
    [TestCase('LibraryMissing', MissingLibraryName)]
    [TestCase('FunctionMissing', LibraryWithoutBCryptName)]
    procedure GenerateSecureRandomString_BCryptUnavailable_RaisesRandomException(const LibraryName: string);
    [Test]
    procedure GenerateSecureRandomString_ZeroLengthWithBCryptUnavailable_ReturnsEmpty;
    [Test]
    procedure GenerateSecureRandomString_LibraryOnlyInApplicationDirectory_RaisesRandomException;

    [Test]
    procedure GenerateNonce_TwoCalls_ReturnDifferentValues;
    [Test]
    procedure GenerateNonce_BCryptUnavailable_RaisesRandomException;
  end;

implementation

uses
  System.IOUtils,
  MSGraph.OAuth2.Types,
  MSGraph.OAuth2.PKCE.Random.Windows,
  MSGraph.OAuth2.PKCE.Random.Windows.Stub;

const
  UrlSafeCharacters   = ['A'..'Z', 'a'..'z', '0'..'9', '-', '_'];
  PlantedLibraryName  = 'MSGraph.Planted.dll';
  SystemLibraryFolder = 'System32';
  SystemBCryptName    = 'bcrypt.dll';

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

procedure TOAuth2PKCERandomWindowsTests.GenerateSecureRandomString_BCryptUnavailable_RaisesRandomException(const LibraryName: string);
begin
  const UnavailableRandom: IOAuth2PKCERandom = TOAuth2PKCERandomWindowsStub.Create(LibraryName);

  Assert.WillRaise(
    procedure
    begin
      UnavailableRandom.GenerateSecureRandomString(43);
    end,
    EOAuth2RandomException);
end;

procedure TOAuth2PKCERandomWindowsTests.GenerateSecureRandomString_ZeroLengthWithBCryptUnavailable_ReturnsEmpty;
begin
  const UnavailableRandom: IOAuth2PKCERandom = TOAuth2PKCERandomWindowsStub.Create(MissingLibraryName);

  const Generated = UnavailableRandom.GenerateSecureRandomString(0);

  Assert.IsEmpty(Generated);
end;

procedure TOAuth2PKCERandomWindowsTests.GenerateSecureRandomString_LibraryOnlyInApplicationDirectory_RaisesRandomException;
begin
  const ApplicationPath      = ParamStr(0);
  const ApplicationDirectory = TPath.GetDirectoryName(ApplicationPath);
  const PlantedPath          = TPath.Combine(ApplicationDirectory, PlantedLibraryName);

  const PlantedRandom: IOAuth2PKCERandom = TOAuth2PKCERandomWindowsStub.Create(PlantedLibraryName);

  PlantLibrary(PlantedPath);
  try
    Assert.WillRaise(
      procedure
      begin
        PlantedRandom.GenerateSecureRandomString(43);
      end,
      EOAuth2RandomException);
  finally
    TFile.Delete(PlantedPath);
  end;
end;

procedure TOAuth2PKCERandomWindowsTests.GenerateNonce_TwoCalls_ReturnDifferentValues;
begin
  const First  = FRandom.GenerateNonce;
  const Second = FRandom.GenerateNonce;

  Assert.AreNotEqual(First, Second);
end;

procedure TOAuth2PKCERandomWindowsTests.GenerateNonce_BCryptUnavailable_RaisesRandomException;
begin
  const UnavailableRandom: IOAuth2PKCERandom = TOAuth2PKCERandomWindowsStub.Create(MissingLibraryName);

  Assert.WillRaise(
    procedure
    begin
      UnavailableRandom.GenerateNonce;
    end,
    EOAuth2RandomException);
end;

procedure TOAuth2PKCERandomWindowsTests.PlantLibrary(const PlantedPath: string);
begin
  const WindowsDirectory = GetEnvironmentVariable('SystemRoot');
  const SystemDirectory  = TPath.Combine(WindowsDirectory, SystemLibraryFolder);
  const SystemBCryptPath = TPath.Combine(SystemDirectory, SystemBCryptName);

  TFile.Copy(SystemBCryptPath, PlantedPath, True);
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
