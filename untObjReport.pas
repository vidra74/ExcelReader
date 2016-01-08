unit untObjReport;

interface

uses Classes;

type
  TIzvjestajPodaci = record
    ID: Integer;
    Ticker: String;
    DatumOd: TDate;
    DatumDo: TDate;
    Opis: String;
  end;

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
    function saveReportInfo(Slog: TIzvjestajPodaci): Boolean;
    property Path: String read mPath;
    property Sheets: TStringList read mSheetList write mSheetList;
    property Status: Boolean read mStatus;
    property StatusMessage: String read mStatusMessage;
end;

implementation

{ TObjReport }

uses untDMMain,
      SysUtils;

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

function TObjReport.saveReportInfo(Slog: TIzvjestajPodaci): Boolean;
begin
  try

    Result := False;

    if not FileExists(DMMain.cdsPregledIzvjestaja.FileName) then
      DMMain.cdsPregledIzvjestaja.CreateDataSet
    else
      DMMain.cdsPregledIzvjestaja.Open;

    if DMMain.cdsPregledIzvjestaja.Locate('OPIS', Slog.Opis, []) = True then begin
      mStatusMessage := DMMain.cdsPregledIzvjestaja.FileName +
                          ' ima ID ' + DMMain.cdsPregledIzvjestaja.FieldByName('ID').AsString;
      Exit;
    end;

    if DMMain.cdsPregledIzvjestaja.IsEmpty then
      Slog.ID := 1
    else begin
      Slog.ID := DMMain.cdsPregledIzvjestaja.RecordCount + 1;
    end;



    try
      DMMain.cdsPregledIzvjestaja.Insert;
      DMMain.cdsPregledIzvjestaja.FieldByName('ID').AsInteger := Slog.ID;
      DMMain.cdsPregledIzvjestaja.FieldByName('PATH').AsString := Self.Path;
      DMMain.cdsPregledIzvjestaja.FieldByName('TICKER').AsString := Slog.Ticker;
      DMMain.cdsPregledIzvjestaja.FieldByName('DATUMUNOSA').AsDateTime := Date;
      DMMain.cdsPregledIzvjestaja.FieldByName('DATUMIZVJESTAJA').AsDateTime := Slog.DatumDo;
      DMMain.cdsPregledIzvjestaja.FieldByName('OPIS').AsString := Slog.Opis;
      DMMain.cdsPregledIzvjestaja.Post;
    except
      On E:Exception do
        mStatusMessage := 'Puknuo insert: ' + E.Message;
    end;

    Result := True;
  finally
    DMMain.cdsPregledIzvjestaja.Close;
  end;

end;

end.
