unit untObjReportSheet;

interface

uses Classes;

type
  TSheetIznosi = record
    ID_Report: Integer;       // ID izvještaja koji se gleda
    ID_Sheet: Integer;        // ID sheet-a u izvještaju kojeg èitamo (Bilanca, RDG, NT...)
    AOP: Integer;             // AOD ID broj sloga izvještaja, propisano
    Pret_Tromjesec: Currency; // Prethodni period - tromjeseèje
    Pret_Kumulativ: Currency; // Prethodni period - zajedno
    Tren_Tromjesec: Currency; // Trenutni period - tromjeseèje
    Tren_Kumulativ: Currency; // Trenutni period - zajedno
  end;

type
  TSheet = class (TObject)
    private

      mID_Report: Integer;
      mID_Sheet: Integer;
      mSheetName: String;
      mStatusMessage: String;
      mStatus: Boolean;
      mFieldList: TStringList;
      mIznosList: array of TSheetIznosi;
      mCount: Integer;
    public

      constructor Create(IDIzvjestaj: Integer; Sheet: String);
      destructor Destroy; override;
      function readSelectedSheet(Name, Path: String): Boolean;
      function showSelectedRecord(ID:Integer):TSheetIznosi;
      property ID_Report: Integer read mID_Report;
      property ID_Sheet: Integer read mID_Sheet;
      property SheetName: String read mSheetName;
      property Status: Boolean read mStatus;
      property StatusMessage: String read mStatusMessage;
      property Fields: TStringList read mFieldList write mFieldList;
      property Count: Integer read mCount;
  end;

implementation

{ TSheet }

uses untDMMain,
      SysUtils,
      untObjBilancaList,
      TypInfo,   // GetEnumType
      DB;        // TFieldType

constructor TSheet.Create(IDIzvjestaj: Integer; Sheet: String);
begin
  inherited Create();

  mID_Report := IDIzvjestaj;
  mID_Sheet := 0;
  mSheetName := Sheet;

  Fields := TStringList.Create;
  SetLength(mIznosList, 20);

  mCount := 0;

end;

destructor TSheet.Destroy;
begin
  Fields.Clear;
  Fields.Free;
  DMMain.qryExcel.Close;
  inherited;
end;

function TSheet.readSelectedSheet(Name, Path: String): Boolean;
var
  Slog: TSheetIznosi;
  i      : integer;
  ft     : TFieldType;
  sft    : string;
  fname  : string;
  DBSet  : TObjBilancaList;
begin

  Result := False;

  mID_Sheet := 0;
  if Pos('Bilanca', Name) > 0 then
    mID_Sheet := 1;
  if Pos('RDG', Name) > 0 then
    mID_Sheet := 2;
  if Pos('NT_I', Name) > 0 then
    mID_Sheet := 3;

  if ID_Sheet = 0 then Exit;
  mSheetName := Name;

  Slog.ID_Report        := ID_Report;
  Slog.ID_Sheet         := ID_Sheet;
  Slog.AOP              := 0;
  Slog.Pret_Tromjesec   := 0.0;
  Slog.Pret_Kumulativ   := 0.0;
  Slog.Tren_Tromjesec   := 0.0;
  Slog.Tren_Kumulativ   := 0.0;

  DMMain.qryExcel.Close;
  DMMain.qryExcel.SQL.Text :=  'select * from [' + Name + ']';
  try
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

      DMMain.qryExcel.First;

      DBSet := TObjBilancaList.Create(Path);
      DBSet.Open;
      mCount := 0;
      while not DMMain.qryExcel.EOF do
      begin
        if (Trim(DMMain.qryExcel.FieldByName('BILANCA').AsString) <> '') then
          if (Trim(DMMain.qryExcel.FieldByName('F9').AsString) <> '') then
          begin
            Slog.AOP := DMMain.qryExcel.FieldByName('F9').AsInteger;
            Slog.Pret_Tromjesec := 0.0;
            Slog.Pret_Kumulativ := DMMain.qryExcel.FieldByName('F10').AsCurrency;
            Slog.Tren_Tromjesec := 0.0;
            Slog.Tren_Kumulativ := DMMain.qryExcel.FieldByName('F11').AsCurrency;

            DBSet.AddNewAmmount(Slog);
            if Length(mIznosList) = Count then
              SetLength(mIznosList, Count + 40);

            mIznosList[mCount] := Slog;
            Inc(mCount);
          end;

        DMMain.qryExcel.Next;
      end;
    except
      mStatusMessage := 'Ne mogu otvoriti Sheet ' + SheetName;
      Result := false;
    end;
  finally
    ;
  end;
end;

function TSheet.showSelectedRecord(ID:Integer):TSheetIznosi;
begin
  Result := mIznosList[ID];
end;

end.
