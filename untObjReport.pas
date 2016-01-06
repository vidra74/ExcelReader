unit untObjReport;

interface

uses Classes;

type TObjReport = class (TObject)
  private
    mPath: String;
    mSheetList: TStringList;
    procedure clearSheets;
  public
    constructor Create(ReportPath: String);
    destructor Destroy; override;
    function analyzeExcelReport: Boolean;
    property Path: String read mPath;
    property Sheets: TStringList read mSheetList write mSheetList;
end;

implementation

{ TObjReport }

function TObjReport.analyzeExcelReport: Boolean;
var
  I: Integer;
  bBilanca, bRDG, bNT: Boolean;
begin

  bBilanca  := false;
  bRDG      := false;
  bNT       := false;

  for I := 0 to Sheets.Count do
  begin
    if not bBilanca then
      bBilanca := (Pos('Bilanca', Sheets.ValueFromIndex[I]) > -1);
    if not bRDG then
      bRDG := (Pos('RDG', Sheets.ValueFromIndex[I]) > -1);
    if not bNT then
      bNT := (Pos('NT_I', Sheets.ValueFromIndex[I]) > -1);
  end;
  Result := bBilanca and bRDG and bNT;
end;

procedure TObjReport.clearSheets;
begin
  Sheets.Clear;
  Sheets.Free;
end;

constructor TObjReport.Create(ReportPath: String);
begin
  inherited Create;

  mPath := ReportPath;
  if Sheets <> nil then clearSheets;

  Sheets := TStringList.Create;

end;

destructor TObjReport.Destroy;
begin
  clearSheets;
  inherited;
end;

end.
