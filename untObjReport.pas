unit untObjReport;

interface

uses Classes,
      untObjReportSheet;

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
    mFieldList: TStringList;
    mStatus: Boolean;
    mStatusMessage: String;
    procedure clearSheets;
  public
    objSheet: TSheet;
    constructor Create(ReportPath: String);
    destructor Destroy; override;
    function analyzeExcelReport: Boolean;
    function napuniIzvjestajRecord(var Slog: TIzvjestajPodaci): Boolean;
    function otvoriOdabraniSheet(Sheet: String): Boolean;
    function Open: Boolean;
    procedure Close;
    property Path: String read mPath;
    property Sheets: TStringList read mSheetList write mSheetList;
    property Fields: TStringList read mFieldList write mFieldList;
    property Status: Boolean read mStatus;
    property StatusMessage: String read mStatusMessage;
end;

implementation

{ TObjReport }

uses untDMMain,
      SysUtils,
      DateUtils,
      TypInfo,   // GetEnumType
      DB;        // TFieldType;

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
  Fields.Clear;
  Fields.Free;
  objSheet.Destroy;
end;

procedure TObjReport.Close;
begin
  DMMain.adoConectExcel.Close;
end;

constructor TObjReport.Create(ReportPath: String);
begin
  inherited Create;

  mPath := ReportPath;
  if Sheets <> nil then clearSheets;
  Sheets := TStringList.Create;
  Fields := TStringList.Create;
  objSheet := TSheet.Create(1, 'Bilanca');
  mStatus := True;
  mStatusMessage := 'Empty Report object created for ' + ReportPath;
end;

destructor TObjReport.Destroy;
begin
  clearSheets;
  inherited;
end;

function TObjReport.napuniIzvjestajRecord(var Slog: TIzvjestajPodaci): Boolean;
begin
  DMMain.qryIzvjestajPodaci.Close;
  DMMain.qryIzvjestajPodaci.SQL.Text :=  'select * from [OPÆI PODACI$]';
  try
    try
      DMMain.qryIzvjestajPodaci.Open;

      if DMMain.qryIzvjestajPodaci.IsEmpty then
      begin
        mStatusMessage := 'Nema podataka u sheetu OPÆI PODACI';
        Result := false;
        Exit;
      end;

      Slog.DatumDo  := StrToDate(DMMain.qryIzvjestajPodaci.FieldByName('F8').AsString);
      try

        Slog.DatumOd  := StrToDate(DMMain.qryIzvjestajPodaci.FieldByName('F5').AsString);
      except
        Slog.DatumOd  := EncodeDate(YearOf(Slog.DatumDo), 1, 1);
      end;
      Slog.Opis     := ExtractFileName(Self.Path);
      Slog.Ticker   := Copy(Slog.Opis, 1, 4);
      Result := True;
    except
      On E:Exception do begin
        mStatusMessage := 'Ne mogu napuniti poodatke izvještaja ' + E.Message;
        Result := False;
      end;
    end;

  finally
    DMMain.qryIzvjestajPodaci.Close;
  end;
end;

function TObjReport.Open: Boolean;
begin

  Result := false;
  DMMain.adoConectExcel.Close;
  DMMain.adoConectExcel.ConnectionString := 'Provider=Microsoft.Jet.OLEDB.4.0;Data Source=' +
    Self.Path + ';Extended Properties=Excel 8.0;Persist Security Info=True;';

  try

    DMMain.adoConectExcel.Open;
    DMMain.adoConectExcel.GetTableNames(Self.Sheets, True);
    Result := True;
  except
    On E:Exception do begin
      mStatus := false;
      mStatusMessage := 'Otvaranje Excel datoteke: ' + E.Message;
      Result := false;
    end;
  end;
end;

function TObjReport.otvoriOdabraniSheet(Sheet: String): Boolean;
var
  i      : integer;
  ft     : TFieldType;
  sft    : string;
  fname  : string;
begin

  DMMain.qryExcel.Close;
  DMMain.qryExcel.SQL.Text :=  'select * from [' + Sheet + ']';
  try
    DMMain.qryExcel.Open;
    Result := DMMain.qryExcel.Active;

    Fields.Clear;
    for i := 0 to DMMain.qryExcel.Fields.Count - 1 do
    begin
      ft := DMMain.qryExcel.Fields[i].DataType;
      sft := GetEnumName(TypeInfo(TFieldType), Integer(ft));
      fname:= DMMain.qryExcel.Fields[i].FieldName;

      Fields.Add(Format('%d) NAME: %s TYPE: %s', [1+i, fname, sft]));
    end;
  except
    mStatusMessage := 'Ne mogu otvoriti Sheet ' + Sheet;
    Result := false;
  end;
end;

end.
