unit untExcelReader;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Grids, DBGrids, ExtCtrls, DB, ADODB, DBClient,
  untObjReport;

type
  TIzvjestajPodaci = record
    ID: Integer;
    Ticker: String;
    DatumOd: TDate;
    DatumDo: TDate;
    Opis: String;
  end;

type
  TFrmExcelReader = class(TForm)
    pnlDohvat: TPanel;
    dbgExcel: TDBGrid;
    pnlBotuni: TPanel;
    btnZatvori: TButton;
    btnOtvoriExcel: TButton;
    dlgOpenExcel: TOpenDialog;
    lblExcelDatoteka: TLabel;
    dsExcel: TDataSource;
    cboxExcelSheets: TComboBox;
    btnOtvoriSheet: TButton;
    ListBox1: TListBox;
    btnSpremiIzvjestaj: TButton;
    procedure btnZatvoriClick(Sender: TObject);
    procedure btnOtvoriExcelClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure btnOtvoriSheetClick(Sender: TObject);
    procedure btnSpremiIzvjestajClick(Sender: TObject);
  private
    { Private declarations }
    Izvjestaj: TObjReport;
    function otvoriOdabraniSheet(Sheet: String): Boolean;
    procedure posaljiSheetUGrid;
    procedure zatvoriGrid;
    procedure GetFieldInfo;
    function napuniIzvjestajRecord(var Slog: TIzvjestajPodaci): Boolean;
  public
    { Public declarations }
    IzvjestajiPodaci: TIzvjestajPodaci;
  end;

var
  FrmExcelReader: TFrmExcelReader;

implementation

{$R *.dfm}

uses typinfo,
      DateUtils,
      untDMMain;

procedure TFrmExcelReader.btnOtvoriExcelClick(Sender: TObject);
begin
  dlgOpenExcel.Execute();

  if dlgOpenExcel.FileName = '' then Exit;

  // ako je izvještaj veæ kreiran create æe ga ponovo rekreirati
  Izvjestaj := TObjReport.Create(dlgOpenExcel.FileName);

  lblExcelDatoteka.Caption := Izvjestaj.Path;
  DMMain.adoConectExcel.Close;
  DMMain.adoConectExcel.ConnectionString := 'Provider=Microsoft.Jet.OLEDB.4.0;Data Source=' +
    dlgOpenExcel.FileName + ';Extended Properties=Excel 8.0;Persist Security Info=True;';

  try
    DMMain.adoConectExcel.Open;
    DMMain.adoConectExcel.GetTableNames(Izvjestaj.Sheets, True);
    cboxExcelSheets.Items.AddStrings(Izvjestaj.Sheets);

    if not Izvjestaj.analyzeExcelReport then
    begin
      MessageDlg('Neispravan Excel izvještaj', mtError, [mbOk], 0);
    end else begin
      cboxExcelSheets.ItemIndex := 0;
      btnOtvoriSheetClick(Sender);
    end;

  except
    On E:Exception do
      MessageDlg('adoconnExcel.Open : ' + E.Message, mtError, [mbOk], 0);
  end;
end;

procedure TFrmExcelReader.btnOtvoriSheetClick(Sender: TObject);
begin
  if not (DMMain.adoConectExcel.Connected) then Exit;

  if otvoriOdabraniSheet(cboxExcelSheets.Items[cboxExcelSheets.ItemIndex]) then
    posaljiSheetUGrid
  else
    zatvoriGrid;
end;

procedure TFrmExcelReader.btnSpremiIzvjestajClick(Sender: TObject);
var cdsPath: String;
begin
  // ima li veæ izvještaj sa odabranim pathom ?

  cdsPath := ExtractFilePath(Application.ExeName) + 'Izvjestaji.xml';
  DMMain.cdsPregledIzvjestaja.FileName := cdsPath;

  if not napuniIzvjestajRecord(IzvjestajiPodaci) then Exit;

  try
    if not FileExists(cdsPath) then
      DMMain.cdsPregledIzvjestaja.CreateDataSet
    else
      DMMain.cdsPregledIzvjestaja.Open;

    if DMMain.cdsPregledIzvjestaja.Locate('OPIS', IzvjestajiPodaci.Opis, []) = True then begin
      ShowMessage(Izvjestaj.Path + ' ima ID ' + DMMain.cdsPregledIzvjestaja.FieldByName('ID').AsString);
      Exit;
    end;

    if DMMain.cdsPregledIzvjestaja.IsEmpty then
      IzvjestajiPodaci.ID := 1
    else begin
      IzvjestajiPodaci.ID := DMMain.cdsPregledIzvjestaja.RecordCount + 1;
    end;



    try
      DMMain.cdsPregledIzvjestaja.Insert;
      DMMain.cdsPregledIzvjestaja.FieldByName('ID').AsInteger := IzvjestajiPodaci.ID;
      DMMain.cdsPregledIzvjestaja.FieldByName('PATH').AsString := Izvjestaj.Path;
      DMMain.cdsPregledIzvjestaja.FieldByName('TICKER').AsString := IzvjestajiPodaci.Ticker;
      DMMain.cdsPregledIzvjestaja.FieldByName('DATUMUNOSA').AsDateTime := Date;
      DMMain.cdsPregledIzvjestaja.FieldByName('DATUMIZVJESTAJA').AsDateTime := IzvjestajiPodaci.DatumDo;
      DMMain.cdsPregledIzvjestaja.FieldByName('OPIS').AsString := IzvjestajiPodaci.Opis;
      DMMain.cdsPregledIzvjestaja.Post;
    except
      On E:Exception do
        ShowMessage('Puknuo insert: ' + E.Message);
    end;

    ShowMessage('Spremio podatke izvještaja: ' + IzvjestajiPodaci.Opis + ' Id: ' + IntToStr(IzvjestajiPodaci.ID));
  finally
    DMMain.cdsPregledIzvjestaja.Close;
  end;



end;

procedure TFrmExcelReader.btnZatvoriClick(Sender: TObject);
begin
  Close;
end;

procedure TFrmExcelReader.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  FreeAndNil(Izvjestaj);
  DMMain.qryExcel.Close;
  DMMain.adoConectExcel.Close;
end;

// Puni list box sa podacima o tipovima kolona iz odabranog sheeta

procedure TFrmExcelReader.GetFieldInfo;
var
  i      : integer;
  ft     : TFieldType;
  sft    : string;
  fname  : string;
begin
  ListBox1.Clear;
  for i := 0 to DMMain.qryExcel.Fields.Count - 1 do
  begin
    ft := DMMain.qryExcel.Fields[i].DataType;
    sft := GetEnumName(TypeInfo(TFieldType), Integer(ft));
    fname:= DMMain.qryExcel.Fields[i].FieldName;

    ListBox1.Items.Add(Format('%d) NAME: %s TYPE: %s', [1+i, fname, sft]));
  end;
end;

function TFrmExcelReader.napuniIzvjestajRecord(var Slog: TIzvjestajPodaci): Boolean;
begin
  DMMain.qryIzvjestajPodaci.Close;
  DMMain.qryIzvjestajPodaci.SQL.Text :=  'select * from [OPÆI PODACI$]';
  try
    try
      DMMain.qryIzvjestajPodaci.Open;

      if DMMain.qryIzvjestajPodaci.IsEmpty then
      begin
        ShowMessage('Nema podataka u sheetu OPÆI PODACI');
        Result := false;
        Exit;
      end;

      Slog.DatumDo  := StrToDate(DMMain.qryIzvjestajPodaci.FieldByName('F8').AsString);
      try

        Slog.DatumOd  := StrToDate(DMMain.qryIzvjestajPodaci.FieldByName('F5').AsString);
      except
        Slog.DatumOd  := EncodeDate(YearOf(Slog.DatumDo), 1, 1);
      end;
      Slog.Opis     := ExtractFileName(Izvjestaj.Path);
      Slog.Ticker   := Copy(Slog.Opis, 1, 4);
      Result := True;
    except
      On E:Exception do begin
        ShowMessage('Ne mogu napuniti poodatke izvještaja ' + E.Message);
        Result := False;
      end;
    end;

  finally
    DMMain.qryIzvjestajPodaci.Close;
  end;


end;

function TFrmExcelReader.otvoriOdabraniSheet(Sheet: String): Boolean;
begin
  DMMain.qryExcel.Close;
  DMMain.qryExcel.SQL.Text :=  'select * from [' + Sheet + ']';
  try
    DMMain.qryExcel.Open;
    Result := DMMain.qryExcel.Active;
  except
    ShowMessage('Ne mogu otvoriti Sheet ' + Sheet);
    Result := false;
  end;
end;

procedure TFrmExcelReader.posaljiSheetUGrid;
begin
  zatvoriGrid;
  GetFieldInfo;
end;

// Zatvaranje grida prazni grid i list kontrolu sa tipovima polja
// Raditi prije uèitavanja novog Excel-a ili uèitavanja sheeta

procedure TFrmExcelReader.zatvoriGrid;
begin
  dbgExcel.Columns.Clear;
  ListBox1.Clear;
end;

end.
