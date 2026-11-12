unit MSGraph.OAuth2.PKCE.Tests;

interface

uses
  DUnitX.TestFramework;

type
  [TestFixture]
  [Category('OAuth2.PKCE')]
  TOAuth2PKCETests = class
  public
    [Test]
    procedure GenerateCodeChallenge_Rfc7636Example_ReturnsExpectedChallenge;

    [Test]
    procedure Generate_GivenRandom_TakesVerifierAndStateFromThatRandom;
    [Test]
    procedure Generate_GivenRandom_ChallengeBelongsToVerifier;
  end;

implementation

uses
  MSGraph.OAuth2.Types,
  MSGraph.OAuth2.PKCE,
  MSGraph.OAuth2.PKCE.Random.Interfaces,
  MSGraph.OAuth2.PKCE.Random.Stub;

procedure TOAuth2PKCETests.GenerateCodeChallenge_Rfc7636Example_ReturnsExpectedChallenge;
begin
  const CodeVerifier = 'dBjftJeZ4CVP-mB92K27uhbUJU1p1r_wW1gFWFOEjXk';

  const Challenge = TOAuth2PKCE.GenerateCodeChallenge(CodeVerifier);

  Assert.AreEqual('E9Melhoa2OwvFrEMTJguCHaoeK1t8URWbuGJSstw-cM', Challenge);
end;

procedure TOAuth2PKCETests.Generate_GivenRandom_TakesVerifierAndStateFromThatRandom;
begin
  const Random: IOAuth2PKCERandom = TOAuth2PKCERandomStub.Create;

  const Session = TOAuth2PKCE.Generate(Random);

  Assert.AreEqual(128, Session.CodeVerifier.Length);
  Assert.StartsWith('AAECAwQF', Session.CodeVerifier);
  Assert.AreEqual('000102030405060708090a0b0c0d0e0f', Session.State);
end;

procedure TOAuth2PKCETests.Generate_GivenRandom_ChallengeBelongsToVerifier;
begin
  const Random: IOAuth2PKCERandom = TOAuth2PKCERandomStub.Create;

  const Session           = TOAuth2PKCE.Generate(Random);
  const ExpectedChallenge = TOAuth2PKCE.GenerateCodeChallenge(Session.CodeVerifier);

  Assert.AreEqual(ExpectedChallenge, Session.CodeChallenge);
end;

end.
