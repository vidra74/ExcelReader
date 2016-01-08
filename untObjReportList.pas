unit untObjReportList;

interface

uses untObjReport;

type TObjReportList = class (TObject)
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
    function AddNewReport(Slog: TIzvjestajPodaci): Integer;
    property Path: String read mPath;
    property Status: Boolean read mStatus;
    property StatusMessage: String read mStatusMessage;
    property Active: Boolean read GetActive;
end;

implementation

uses untDMMain,
      SysUtils;

{ TObjReportList }

function TObjReportList.AddNewReport(Slog: TIzvjestajPodaci): Integer;
begin

  if DMMain.cdsPregledIzvjestaja.IsEmpty then
    Result := 1
  else begin
    Result := DMMain.cdsPregledIzvjestaja.RecordCount + 1;
  end;

  try
    DMMain.cdsPregledIzvjestaja.Insert;
    DMMain.cdsPregledIzvjestaja.FieldByName('ID').AsInteger := Result;
    DMMain.cdsPregledIzvjestaja.FieldByName('PATH').AsString := Self.Path;
    DMMain.cdsPregledIzvjestaja.FieldByName('TICKER').AsString := Slog.Ticker;
    DMMain.cdsPregledIzvjestaja.FieldByName('DATUMUNOSA').AsDateTime := Date;
    DMMain.cdsPregledIzvjestaja.FieldByName('DATUMIZVJESTAJA').AsDateTime := Slog.DatumDo;
    DMMain.cdsPregledIzvjestaja.FieldByName('OPIS').AsString := Slog.Opis;
    DMMain.cdsPregledIzvjestaja.Post;
  except
    On E:Exception do begin

      mStatusMessage := 'Puknuo insert: ' + E.Message;
      Result := -1;
    end;
  end;

end;

procedure TObjReportList.Close;
begin
  DMMain.cdsPregledIzvjestaja.Close;
end;

constructor TObjReportList.Create(CdsPath: String);
begin
  mPath := cdsPath;
end;

destructor TObjReportList.Destroy;
begin
  Self.Close;
  inherited;
end;

function TObjReportList.GetActive: Boolean;
begin

  Result := DMMain.cdsPregledIzvjestaja.Active;
end;

function TObjReportList.Locate(Keys: String; Values: Variant): Integer;
begin

  Result := -1;
  if DMMain.cdsPregledIzvjestaja.Locate(Keys, Values, []) then
    Result := DMMain.cdsPregledIzvjestaja.FieldByName('ID').AsInteger;
end;

procedure TObjReportList.Open;
begin

  DMMain.cdsPregledIzvjestaja.FileName := Path + 'Izvjestaji.xml';
  if not FileExists(DMMain.cdsPregledIzvjestaja.FileName) then
    DMMain.cdsPregledIzvjestaja.CreateDataSet
  else
    DMMain.cdsPregledIzvjestaja.Open;
end;

end.
