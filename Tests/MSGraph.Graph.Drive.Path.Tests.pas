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
    procedure ItemAddress_PathUnderFolder_AddressesRelativeToTheFolderItem;
    [Test]
    procedure ItemAddress_EmptyPathUnderFolder_IsTheFolderItself;
    [Test]
    procedure ChildrenAddress_UnderFolder_AddressesRelativeToTheFolderItem;
    [Test]
    procedure ChildrenAddress_EmptyPathUnderFolder_ListsTheFolderChildren;
    [Test]
    procedure ContentAddress_UnderFolder_AddressesRelativeToTheFolderItem;
    [Test]
    procedure ItemAddress_SpecialCharactersUnderFolder_EncodesEachSegment;

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
  MSGraph.Graph.Drive.Types,
  MSGraph.Graph.Drive.Path,
  MSGraph.Graph.Drive.Exceptions;

const
  FolderAddress = '/items/ABC';

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
  const Address = TGraphDrivePath.ItemAddress(GraphDriveRootAddress, Path);

  Assert.AreEqual(Expected, Address);
end;

procedure TGraphDrivePathTests.ChildrenAddress_UsesPathAddressing(const Path, Expected: string);
begin
  const Address = TGraphDrivePath.ChildrenAddress(GraphDriveRootAddress, Path);

  Assert.AreEqual(Expected, Address);
end;

procedure TGraphDrivePathTests.ItemAddress_PathUnderFolder_AddressesRelativeToTheFolderItem;
begin
  const Address = TGraphDrivePath.ItemAddress(FolderAddress, 'Notes/a.md');

  Assert.AreEqual('/items/ABC:/Notes/a.md:', Address);
end;

procedure TGraphDrivePathTests.ItemAddress_EmptyPathUnderFolder_IsTheFolderItself;
begin
  const Address = TGraphDrivePath.ItemAddress(FolderAddress, '');

  Assert.AreEqual(FolderAddress, Address);
end;

procedure TGraphDrivePathTests.ChildrenAddress_UnderFolder_AddressesRelativeToTheFolderItem;
begin
  const Address = TGraphDrivePath.ChildrenAddress(FolderAddress, 'Notes');

  Assert.AreEqual('/items/ABC:/Notes:/children', Address);
end;

procedure TGraphDrivePathTests.ChildrenAddress_EmptyPathUnderFolder_ListsTheFolderChildren;
begin
  const Address = TGraphDrivePath.ChildrenAddress(FolderAddress, '');

  Assert.AreEqual('/items/ABC/children', Address);
end;

procedure TGraphDrivePathTests.ContentAddress_UnderFolder_AddressesRelativeToTheFolderItem;
begin
  const Address = TGraphDrivePath.ContentAddress(FolderAddress, 'Notes/a.md');

  Assert.AreEqual('/items/ABC:/Notes/a.md:/content', Address);
end;

procedure TGraphDrivePathTests.ItemAddress_SpecialCharactersUnderFolder_EncodesEachSegment;
begin
  const Path: string = 'Idee #1 ' + Char($E9) + '.md';

  const Address = TGraphDrivePath.ItemAddress(FolderAddress, Path);

  Assert.AreEqual('/items/ABC:/Idee%20%231%20%C3%A9.md:', Address);
end;

procedure TGraphDrivePathTests.ContentAddress_EmptyPath_RaisesValidationError;
begin
  Assert.WillRaise(
    procedure
    begin
      TGraphDrivePath.ContentAddress(GraphDriveRootAddress, '');
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
