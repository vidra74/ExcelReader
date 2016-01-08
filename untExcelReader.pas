unit untExcelReader;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Grids, DBGrids, ExtCtrls, DB, ADODB, DBClient,
  untObjReport,
  untObjReportList;

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
    procedure FormShow(Sender: TObject);
  private
    { Private declarations }
    Izvjestaj: TObjReport;
    ListaIzvjestaja: TObjReportList;
    function otvoriOdabraniSheet(Sheet: String): Boolean;
    procedure posaljiSheetUGrid;
    procedure zatvoriGrid;
    procedure GetFieldInfo;
    procedure spremiIzvjestajInfo;
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

  // ako je izvje�taj ve� kreiran create �e ga ponovo rekreirati
  Izvjestaj := TObjReport.Create(dlgOpenExcel.FileName);
  lblExcelDatoteka.Caption := Izvjestaj.Path;

  if Izvjestaj.Open then
  begin

    if Izvjestaj.analyzeExcelReport then
    begin

      cboxExcelSheets.Items.AddStrings(Izvjestaj.Sheets);
      cboxExcelSheets.ItemIndex := 0;
      btnOtvoriSheetClick(Sender);
    end else begin

      MessageDlg(Izvjestaj.StatusMessage, mtError, [mbOk], 0);
    end;
  end else
    MessageDlg(Izvjestaj.StatusMessage, mtError, [mbOk], 0);
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
begin

  spremiIzvjestajInfo;
end;

procedure TFrmExcelReader.btnZatvoriClick(Sender: TObject);
begin
  Close;
end;

procedure TFrmExcelReader.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  Izvjestaj.Close;
  ListaIzvjestaja.Close;
  FreeAndNil(Izvjestaj);
  FreeAndNil(ListaIzvjestaja);
end;

procedure TFrmExcelReader.FormShow(Sender: TObject);
begin
  ListaIzvjestaja := TObjReportList.Create(ExtractFilePath(Application.ExeName));
  ListaIzvjestaja.Open;
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
  DMMain.qryIzvjestajPodaci.SQL.Text :=  'select * from [OP�I PODACI$]';
  try
    try
      DMMain.qryIzvjestajPodaci.Open;

      if DMMain.qryIzvjestajPodaci.IsEmpty then
      begin
        ShowMessage('Nema podataka u sheetu OP�I PODACI');
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
        ShowMessage('Ne mogu napuniti poodatke izvje�taja ' + E.Message);
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

procedure TFrmExcelReader.spremiIzvjestajInfo;
begin

  if not napuniIzvjestajRecord(IzvjestajiPodaci) then
  begin

    MessageDlg('Gre�ka pregleda izvje�taja', mtError, [mbOk], 0);
    Exit;
  end;

  // ima li ve� izvje�taj sa odabranim pathom ?
  IzvjestajiPodaci.ID := ListaIzvjestaja.Locate('OPIS', IzvjestajiPodaci.Opis);

  if IzvjestajiPodaci.ID < 1 then
  begin

    IzvjestajiPodaci.ID := ListaIzvjestaja.AddNewReport(IzvjestajiPodaci);
    if IzvjestajiPodaci.ID < 1 then
      MessageDlg(ListaIzvjestaja.StatusMessage, mtError, [mbOk], 0)
    else
      MessageDlg('Spremio podatke izvje�taja: ' + IzvjestajiPodaci.Opis + ' Id: ' + IntToStr(IzvjestajiPodaci.ID),
                mtInformation,
                [mbOk],
                0);
   end else
    MessageDlg('Izvje�taj: ' + IzvjestajiPodaci.Opis + ' postoji kao Id: ' + IntToStr(IzvjestajiPodaci.ID),
                mtInformation,
                [mbOk],
                0);
end;

// Zatvaranje grida prazni grid i list kontrolu sa tipovima polja
// Raditi prije u�itavanja novog Excel-a ili u�itavanja sheeta

procedure TFrmExcelReader.zatvoriGrid;
begin
  dbgExcel.Columns.Clear;
  ListBox1.Clear;
end;

end.
