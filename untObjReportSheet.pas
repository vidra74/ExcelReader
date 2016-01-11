unit untObjReportSheet;

interface

type
  TSheetIznosi = record
    ID_Report: Integer;       // ID izvještaja koji se gleda
    ID_Sheet: Integer;        // ID sheet-a u izvještaju kojeg èitamo (Bilanca, RDG, NT...)
    AOD: Integer;             // AOD ID broj sloga izvještaja, propisano
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
    public

      constructor Create(IDIzvjestaj: Integer; Sheet: String);
      destructor Destroy; override;
      function readSelectedSheet: Boolean;
      property ID_Report: Integer read mID_Report;
      property ID_Sheet: Integer read mID_Sheet;
      property SheetName: String read mSheetName;
      property Status: Boolean read mStatus;
      property StatusMessage: String read mStatusMessage;
  end;

implementation

{ TSheet }

uses untDMMain,
      SysUtils;

constructor TSheet.Create(IDIzvjestaj: Integer; Sheet: String);
begin
  inherited Create();

  mID_Report := IDIzvjestaj;
  mID_Sheet := 0;
  mSheetName := Sheet;

  if Pos('Bilanca', Sheet) > 0 then
    mID_Sheet := 1;
  if Pos('RDG', Sheet) > 0 then
    mID_Sheet := 2;
  if Pos('NT_I', Sheet) > 0 then
    mID_Sheet := 3;
end;

destructor TSheet.Destroy;
begin

  inherited;
end;

function TSheet.readSelectedSheet: Boolean;
var
  Slog: TSheetIznosi;
begin
  DMMain.qryIzvjestajPodaci.Close;
  DMMain.qryIzvjestajPodaci.SQL.Text :=  'select * from [' + SheetName + ']';
  try
    try
      DMMain.qryIzvjestajPodaci.Open;

      if DMMain.qryIzvjestajPodaci.IsEmpty then
      begin
        mStatusMessage := 'Nema podataka u sheetu ' + SheetName;
        Result := false;
        Exit;
      end;

      Slog.ID_Report        := ID_Report;
      Slog.ID_Sheet         := ID_Sheet;
      Slog.AOD              := 0;
      Slog.Pret_Tromjesec   := 0.0;
      Slog.Pret_Kumulativ   := 0.0;
      Slog.Tren_Tromjesec   := 0.0;
      Slog.Tren_Kumulativ   := 0.0;

      Result := True;
    except
      On E:Exception do begin
        mStatusMessage := 'Ne mogu napuniti podatke izvještaja ' + E.Message;
        Result := False;
      end;
    end;

  finally
    DMMain.qryIzvjestajPodaci.Close;
  end;
end;

end.
