unit MSGraph.Graph.Drive.Path.Tests;

interface

uses
  DUnitX.TestFramework;

type
  [TestFixture]
  TGraphDrivePathTests = class
  public
    [Test]
    [TestCase('Plain', 'Notes/Today.md|Notes/Today.md', '|')]
    [TestCase('SpaceAndHash', 'Notes/2026-10-14 Idee #1.md|Notes/2026-10-14%20Idee%20%231.md', '|')]
    [TestCase('Percent', '100% klaar.md|100%25%20klaar.md', '|')]
    [TestCase('OuterSlashes', '/Notes/|Notes', '|')]
    procedure Encode_PercentEncodesEachSegment(const Path, Expected: string);

    [Test]
    procedure Encode_NonAscii_EncodesTheUtf8Bytes;

    [Test]
    [TestCase('Root', '|/root', '|')]
    [TestCase('Nested', 'Notes/A b.md|/root:/Notes/A%20b.md:', '|')]
    procedure ItemAddress_UsesPathAddressing(const Path, Expected: string);

    [Test]
    [TestCase('Root', '|/root/children', '|')]
    [TestCase('Folder', 'Notes|/root:/Notes:/children', '|')]
    procedure ChildrenAddress_UsesPathAddressing(const Path, Expected: string);

    [Test]
    procedure ContentAddress_EmptyPath_RaisesValidationError;

    [Test]
    [TestCase('Parent', 'Notes/../Secret.md')]
    [TestCase('Current', 'Notes/./Today.md')]
    [TestCase('EmptySegment', 'Notes//Today.md')]
    procedure Encode_UnsafeSegment_RaisesValidationError(const Path: string);
  end;

implementation

uses
  MSGraph.Graph.Drive.Path,
  MSGraph.Graph.Drive.Exceptions;

procedure TGraphDrivePathTests.Encode_PercentEncodesEachSegment(const Path, Expected: string);
begin
  const Encoded = TGraphDrivePath.Encode(Path);

  Assert.AreEqual(Expected, Encoded);
end;

procedure TGraphDrivePathTests.Encode_NonAscii_EncodesTheUtf8Bytes;
begin
  const Path: string = Char($C9) + Char($E9) + 'n/caf' + Char($E9) + '.md';

  const Encoded = TGraphDrivePath.Encode(Path);

  Assert.AreEqual('%C3%89%C3%A9n/caf%C3%A9.md', Encoded);
end;

procedure TGraphDrivePathTests.ItemAddress_UsesPathAddressing(const Path, Expected: string);
begin
  const Address = TGraphDrivePath.ItemAddress(Path);

  Assert.AreEqual(Expected, Address);
end;

procedure TGraphDrivePathTests.ChildrenAddress_UsesPathAddressing(const Path, Expected: string);
begin
  const Address = TGraphDrivePath.ChildrenAddress(Path);

  Assert.AreEqual(Expected, Address);
end;

procedure TGraphDrivePathTests.ContentAddress_EmptyPath_RaisesValidationError;
begin
  Assert.WillRaise(
    procedure
    begin
      TGraphDrivePath.ContentAddress('');
    end,
    EGraphDriveValidationError);
end;

procedure TGraphDrivePathTests.Encode_UnsafeSegment_RaisesValidationError(const Path: string);
begin
  Assert.WillRaise(
    procedure
    begin
      TGraphDrivePath.Encode(Path);
    end,
    EGraphDriveValidationError);
end;

end.
