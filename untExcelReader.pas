unit untExcelReader;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Grids, DBGrids, ExtCtrls, DB, ADODB, DBClient;

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
    adoconectExcel: TADOConnection;
    qryExcel: TADOQuery;
    dsExcel: TDataSource;
    cboxExcelSheets: TComboBox;
    btnOtvoriSheet: TButton;
    ListBox1: TListBox;
    btnSpremiIzvjestaj: TButton;
    cdsPregledIzvjestaja: TClientDataSet;
    cdsPregledIzvjestajaID: TIntegerField;
    cdsPregledIzvjestajaPATH: TStringField;
    cdsPregledIzvjestajaOPIS: TStringField;
    cdsPregledIzvjestajaTICKER: TStringField;
    cdsPregledIzvjestajaDATUMUNOSA: TDateField;
    cdsPregledIzvjestajaDATUMIZVJESTAJA: TDateField;
    qryIzvjestajPodaci: TADOQuery;
    procedure btnZatvoriClick(Sender: TObject);
    procedure btnOtvoriExcelClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure btnOtvoriSheetClick(Sender: TObject);
    procedure btnSpremiIzvjestajClick(Sender: TObject);
  private
    { Private declarations }
    IzvjestajPath: String;
    function otvoriOdabraniSheet(Sheet: String): Boolean;
    procedure posaljiSheetUGrid;
    procedure zatvoriGrid;
    procedure GetFieldInfo;
    function analizirajExcelDatoteku: Boolean;
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
      DateUtils;

function TFrmExcelReader.analizirajExcelDatoteku: Boolean;
var
  I: Integer;
  bBilanca, bRDG, bNT: Boolean;
begin

  bBilanca  := false;
  bRDG      := false;
  bNT       := false;

  for I := 0 to cboxExcelSheets.Items.Count do
  begin
    if not bBilanca then
      bBilanca := (Pos('Bilanca', cboxExcelSheets.Items.ValueFromIndex[I]) > -1);
    if not bRDG then
      bRDG := (Pos('RDG', cboxExcelSheets.Items.ValueFromIndex[I]) > -1);
    if not bNT then
      bNT := (Pos('NT_I', cboxExcelSheets.Items.ValueFromIndex[I]) > -1);
  end;
  Result := bBilanca and bRDG and bNT;
end;

procedure TFrmExcelReader.btnOtvoriExcelClick(Sender: TObject);
begin
  dlgOpenExcel.Execute();

  if dlgOpenExcel.FileName = '' then Exit;

  IzvjestajPath := (dlgOpenExcel.FileName);
  lblExcelDatoteka.Caption := IzvjestajPath;
  adoconectExcel.Close;
  adoconectExcel.ConnectionString := 'Provider=Microsoft.Jet.OLEDB.4.0;Data Source=' +
    dlgOpenExcel.FileName + ';Extended Properties=Excel 8.0;Persist Security Info=True;';

  try
    adoconectExcel.Open;
    adoconectExcel.GetTableNames(cboxExcelSheets.Items,True);

    if not analizirajExcelDatoteku then
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
  if not (adoconectExcel.Connected) then Exit;

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
  cdsPregledIzvjestaja.FileName := cdsPath;

  if not napuniIzvjestajRecord(IzvjestajiPodaci) then Exit;

  try
    if not FileExists(cdsPath) then
      cdsPregledIzvjestaja.CreateDataSet
    else
      cdsPregledIzvjestaja.Open;

    if cdsPregledIzvjestaja.Locate('OPIS', IzvjestajiPodaci.Opis, []) = True then begin
      ShowMessage(IzvjestajPath + ' ima ID ' + cdsPregledIzvjestaja.FieldByName('ID').AsString);
      Exit;
    end;

    if cdsPregledIzvjestaja.IsEmpty then
      IzvjestajiPodaci.ID := 1
    else begin
      IzvjestajiPodaci.ID := cdsPregledIzvjestaja.RecordCount + 1;
    end;



    try
      cdsPregledIzvjestaja.Insert;
      cdsPregledIzvjestaja.FieldByName('ID').AsInteger := IzvjestajiPodaci.ID;
      cdsPregledIzvjestaja.FieldByName('PATH').AsString := IzvjestajPath;
      cdsPregledIzvjestaja.FieldByName('TICKER').AsString := IzvjestajiPodaci.Ticker;
      cdsPregledIzvjestaja.FieldByName('DATUMUNOSA').AsDateTime := Date;
      cdsPregledIzvjestaja.FieldByName('DATUMIZVJESTAJA').AsDateTime := IzvjestajiPodaci.DatumDo;
      cdsPregledIzvjestaja.FieldByName('OPIS').AsString := IzvjestajiPodaci.Opis;
      cdsPregledIzvjestaja.Post;
    except
      On E:Exception do
        ShowMessage('Puknuo insert: ' + E.Message);
    end;


  finally
    cdsPregledIzvjestaja.Close;
  end;



end;

procedure TFrmExcelReader.btnZatvoriClick(Sender: TObject);
begin
  Close;
end;

procedure TFrmExcelReader.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  qryExcel.Close;
  adoconectExcel.Close;
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
  for i := 0 to qryExcel.Fields.Count - 1 do
  begin
    ft := qryExcel.Fields[i].DataType;
    sft := GetEnumName(TypeInfo(TFieldType), Integer(ft));
    fname:= qryExcel.Fields[i].FieldName;

    ListBox1.Items.Add(Format('%d) NAME: %s TYPE: %s', [1+i, fname, sft]));
  end;
end;

function TFrmExcelReader.napuniIzvjestajRecord(var Slog: TIzvjestajPodaci): Boolean;
begin
  qryIzvjestajPodaci.Close;
  qryIzvjestajPodaci.SQL.Text :=  'select * from [OPÆI PODACI$]';
  try
    try
      qryIzvjestajPodaci.Open;

      if qryIzvjestajPodaci.IsEmpty then
      begin
        ShowMessage('Nema podataka u sheetu OPÆI PODACI');
        Result := false;
        Exit;
      end;

      Slog.DatumDo  := StrToDate(qryIzvjestajPodaci.FieldByName('F8').AsString);
      try

        Slog.DatumOd  := StrToDate(qryIzvjestajPodaci.FieldByName('F5').AsString);
      except
        Slog.DatumOd  := EncodeDate(YearOf(Slog.DatumDo), 1, 1);
      end;
      Slog.Opis     := ExtractFileName(IzvjestajPath);
      Slog.Ticker   := Copy(Slog.Opis, 1, 4);
      Result := True;
    except
      On E:Exception do begin
        ShowMessage('Ne mogu napuniti poodatke izvještaja ' + E.Message);
        Result := False;
      end;
    end;

  finally
    qryIzvjestajPodaci.Close;
  end;


end;

function TFrmExcelReader.otvoriOdabraniSheet(Sheet: String): Boolean;
begin
  qryExcel.Close;
  qryExcel.SQL.Text :=  'select * from [' + Sheet + ']';
  try
    qryExcel.Open;
    Result := qryExcel.Active;
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
