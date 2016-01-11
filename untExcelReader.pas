unit untExcelReader;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Grids, DBGrids, ExtCtrls, DB, ADODB, DBClient,
  untObjReport,
  untObjReportList,
  untObjReportSheet,
  ComCtrls, Menus;

type
  TFrmExcelReader = class(TForm)
    pnlDohvat: TPanel;
    dbgExcel: TDBGrid;
    pnlBotuni: TPanel;
    btnZatvori: TButton;
    btnOtvoriExcel: TButton;
    dlgOpenExcel: TOpenDialog;
    dsExcel: TDataSource;
    cboxExcelSheets: TComboBox;
    btnOtvoriSheet: TButton;
    ListBox1: TListBox;
    btnSpremiIzvjestaj: TButton;
    sbExcelStatus: TStatusBar;
    menExcelReader: TMainMenu;
    menuFile: TMenuItem;
    miOpenExcelReport: TMenuItem;
    miCloseReport: TMenuItem;
    miExit: TMenuItem;
    N1: TMenuItem;
    miSaveReportInfo: TMenuItem;
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
    objSheet: TSheet;
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
  objSheet := TSheet.Create(1, 'Bilanca');
  sbExcelStatus.Panels[0].Text := Izvjestaj.Path;

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
  objSheet.Destroy;
end;

procedure TFrmExcelReader.FormShow(Sender: TObject);
begin
  ListaIzvjestaja := TObjReportList.Create(ExtractFilePath(Application.ExeName));
  ListaIzvjestaja.Open;

  Izvjestaj := TObjReport.Create('');
  objSheet := TSheet.Create(1, 'Bilanca');
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

  sbExcelStatus.Panels[2].Text := Izvjestaj.StatusMessage;
end;

procedure TFrmExcelReader.otvoriExcelSheet(SheetName: String);
begin
  if not Izvjestaj.Status then
  begin
    MessageDlg(Izvjestaj.StatusMessage, mtError, [mbOk], 0);
    Exit;
  end;
  objSheet := TSheet.Create(IzvjestajiPodaci.ID, 'Bilanca');
  if objSheet.readSelectedSheet(SheetName, ExtractFilePath(Application.ExeName)) then
    posaljiSheetUGrid
  else begin
    MessageDlg(objSheet.StatusMessage, mtError, [mbOk], 0);
    zatvoriGrid;
  end;

  sbExcelStatus.Panels[2].Text := Izvjestaj.StatusMessage;
end;

procedure TFrmExcelReader.posaljiSheetUGrid;
begin
  zatvoriGrid;
  ListBox1.Items.Clear;
  ListBox1.Items.AddStrings(objSheet.Fields);
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

  sbExcelStatus.Panels[2].Text := Izvjestaj.StatusMessage;
end;

// Zatvaranje grida prazni grid i list kontrolu sa tipovima polja
// Raditi prije uèitavanja novog Excel-a ili uèitavanja sheeta

procedure TFrmExcelReader.zatvoriGrid;
begin
  dbgExcel.Columns.Clear;
  ListBox1.Items.Clear;
end;

end.
