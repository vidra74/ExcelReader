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
    procedure posaljiSheetUGrid;
    procedure zatvoriGrid;

  public
    { Public declarations }
    IzvjestajiPodaci: TIzvjestajPodaci;
    procedure otvoriExcel;
    procedure otvoriExcelSheet(SheetName: String);
    procedure spremiIzvjestajInfo;
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

  otvoriExcel;
end;

procedure TFrmExcelReader.btnOtvoriSheetClick(Sender: TObject);
begin

  otvoriExcelSheet(cboxExcelSheets.Items[cboxExcelSheets.ItemIndex]);
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

procedure TFrmExcelReader.otvoriExcel;
begin

  if Izvjestaj.Open then
  begin

    if Izvjestaj.analyzeExcelReport then
    begin

      cboxExcelSheets.Items.AddStrings(Izvjestaj.Sheets);
      cboxExcelSheets.ItemIndex := 0;
      otvoriExcelSheet(cboxExcelSheets.Items[0]);
    end else begin

      MessageDlg(Izvjestaj.StatusMessage, mtError, [mbOk], 0);
    end;
  end else
    MessageDlg(Izvjestaj.StatusMessage, mtError, [mbOk], 0);
end;

procedure TFrmExcelReader.otvoriExcelSheet(SheetName: String);
begin
  if Izvjestaj.otvoriOdabraniSheet(SheetName) then
    posaljiSheetUGrid
  else begin
    MessageDlg(Izvjestaj.StatusMessage, mtError, [mbOk], 0);
    zatvoriGrid;
  end;
end;

procedure TFrmExcelReader.posaljiSheetUGrid;
begin
  zatvoriGrid;
  ListBox1.Items.Clear;
  ListBox1.Items.AddStrings(Izvjestaj.Fields);
end;

procedure TFrmExcelReader.spremiIzvjestajInfo;
begin

  if not Izvjestaj.napuniIzvjestajRecord(IzvjestajiPodaci) then
  begin

    MessageDlg('Greška pregleda izvještaja ' + Izvjestaj.StatusMessage, mtError, [mbOk], 0);
    Exit;
  end;

  // ima li veæ izvještaj sa odabranim pathom ?
  IzvjestajiPodaci.ID := ListaIzvjestaja.Locate('OPIS', IzvjestajiPodaci.Opis);

  if IzvjestajiPodaci.ID < 1 then
  begin

    IzvjestajiPodaci.ID := ListaIzvjestaja.AddNewReport(IzvjestajiPodaci);
    if IzvjestajiPodaci.ID < 1 then
      MessageDlg(ListaIzvjestaja.StatusMessage, mtError, [mbOk], 0)
    else
      MessageDlg('Spremio podatke izvještaja: ' + IzvjestajiPodaci.Opis + ' Id: ' + IntToStr(IzvjestajiPodaci.ID),
                mtInformation,
                [mbOk],
                0);
   end else
    MessageDlg('Izvještaj: ' + IzvjestajiPodaci.Opis + ' postoji kao Id: ' + IntToStr(IzvjestajiPodaci.ID),
                mtInformation,
                [mbOk],
                0);
end;

// Zatvaranje grida prazni grid i list kontrolu sa tipovima polja
// Raditi prije uèitavanja novog Excel-a ili uèitavanja sheeta

procedure TFrmExcelReader.zatvoriGrid;
begin
  dbgExcel.Columns.Clear;
  ListBox1.Items.Clear;
end;

end.
