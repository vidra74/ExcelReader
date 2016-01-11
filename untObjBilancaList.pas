unit untObjBilancaList;

interface

uses untObjReportSheet;

type TObjBilancaList = class (TObject)
  private

    mPath: String;
    mStatus: Boolean;
    mStatusMessage: String;
    function GetActive: Boolean;
  public

    constructor Create(CdsPath: String);
    destructor Destroy; override;
    procedure Open;
    procedure Close;
    function Locate(Keys: String; Values: Variant): Integer;
    function AddNewAmmount(Slog: TSheetIznosi): Integer;
    property Path: String read mPath;
    property Status: Boolean read mStatus;
    property StatusMessage: String read mStatusMessage;
    property Active: Boolean read GetActive;
end;

implementation

uses untDMMain,
      SysUtils;

{ TObjBilancaList }

function TObjBilancaList.AddNewAmmount(Slog: TSheetIznosi): Integer;
begin

  if DMMain.cdsBilanca.IsEmpty then
    Result := 1
  else begin
    Result := DMMain.cdsBilanca.RecordCount + 1;
  end;

  try
    DMMain.cdsBilanca.Insert;
    DMMain.cdsBilanca.FieldByName('ID_REPORT').AsInteger := Slog.ID_Report;
    DMMain.cdsBilanca.FieldByName('ID_SHEET').AsInteger := Slog.ID_Sheet;
    DMMain.cdsBilanca.FieldByName('AOP').AsInteger := Slog.AOP;
    DMMain.cdsBilanca.FieldByName('PRET_TROM').AsCurrency := Slog.Pret_Tromjesec;
    DMMain.cdsBilanca.FieldByName('PRET_TOT').AsCurrency := Slog.Pret_Kumulativ;
    DMMain.cdsBilanca.FieldByName('TREN_TROM').AsCurrency := Slog.Tren_Tromjesec;
    DMMain.cdsBilanca.FieldByName('TREN_TOT').AsCurrency := Slog.Tren_Kumulativ;
    DMMain.cdsBilanca.FieldByName('VRIJEME').AsDateTime := Now;
    DMMain.cdsBilanca.Post;
  except
    On E:Exception do begin

      mStatusMessage := 'Puknuo insert: ' + E.Message;
      Result := -1;
    end;
  end;

end;

procedure TObjBilancaList.Close;
begin
  DMMain.cdsBilanca.Close;
end;

constructor TObjBilancaList.Create(CdsPath: String);
begin
  mPath := cdsPath;
end;

destructor TObjBilancaList.Destroy;
begin
  Self.Close;
  inherited;
end;

function TObjBilancaList.GetActive: Boolean;
begin

  Result := DMMain.cdsBilanca.Active;
end;

function TObjBilancaList.Locate(Keys: String; Values: Variant): Integer;
begin

  Result := -1;
  if DMMain.cdsBilanca.Locate(Keys, Values, []) then
    Result := DMMain.cdsBilanca.FieldByName('ID').AsInteger;
end;

procedure TObjBilancaList.Open;
begin

  DMMain.cdsBilanca.FileName := Path + 'Bilanca.xml';
  if not FileExists(DMMain.cdsBilanca.FileName) then
    DMMain.cdsBilanca.CreateDataSet
  else
    DMMain.cdsBilanca.Open;
end;

end.
