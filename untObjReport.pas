unit untObjReport;

interface

uses Classes;

type TObjReport = class (TObject)
  private
    mPath: String;
    mSheetList: TStringList;
    mStatus: Boolean;
    mStatusMessage: String;
    procedure clearSheets;
  public
    constructor Create(ReportPath: String);
    destructor Destroy; override;
    function analyzeExcelReport: Boolean;
    property Path: String read mPath;
    property Sheets: TStringList read mSheetList write mSheetList;
    property Status: Boolean read mStatus;
    property StatusMessage: String read mStatusMessage;
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

  for I := 0 to Sheets.Count - 1 do
  begin
    if not bBilanca then
      bBilanca := (Pos('Bilanca', Sheets[I]) > 0);
    if not bRDG then
      bRDG := (Pos('RDG', Sheets[I]) > 0);
    if not bNT then
      bNT := (Pos('NT_I', Sheets[I]) > 0);
  end;

  mStatus := bBilanca and bRDG and bNT;
  if not mStatus then
    mStatusMessage := 'Excel file is not proper Croatian report. Check content for ' + Path
  else
    mStatusMessage := 'Correct Excel file ' + Path;
  Result := mStatus;
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

  mStatus := True;
  mStatusMessage := 'Empty Report object created for ' + ReportPath;
end;

destructor TObjReport.Destroy;
begin
  clearSheets;
  inherited;
end;

end.
