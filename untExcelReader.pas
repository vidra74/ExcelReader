unit untExcelReader;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Grids, DBGrids, ExtCtrls, DB, ADODB, DBClient;

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
  public
    { Public declarations }
  end;

var
  FrmExcelReader: TFrmExcelReader;

implementation

{$R *.dfm}

uses typinfo;

function TFrmExcelReader.analizirajExcelDatoteku: Boolean;
var
  I: Integer;
  bBilanca, bRDG, bNT: Boolean;
begin

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

  lblExcelDatoteka.Caption := (dlgOpenExcel.FileName);

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
    nId: Integer;
begin
  // ima li veæ izvještaj sa odabranim pathom ?

  cdsPath := ExtractFilePath(Application.ExeName) + 'Izvjestaji.xml';
  cdsPregledIzvjestaja.FileName := cdsPath;

  nId := -1;
  try
    if not FileExists(cdsPath) then
      cdsPregledIzvjestaja.CreateDataSet
    else
      cdsPregledIzvjestaja.Open;

    if cdsPregledIzvjestaja.Locate('PATH', IzvjestajPath, []) = True then begin
      ShowMessage(IzvjestajPath + ' ima ID ' + cdsPregledIzvjestaja.FieldByName('ID').AsString);
      Exit;
    end;

    if cdsPregledIzvjestaja.IsEmpty then
      nId := 1
    else begin
      nId := cdsPregledIzvjestaja.RecordCount + 1;
    end;

    try
      cdsPregledIzvjestaja.Insert;
      cdsPregledIzvjestaja.FieldByName('ID').AsInteger := nId;
      cdsPregledIzvjestaja.FieldByName('PATH').AsString := IzvjestajPath;
      cdsPregledIzvjestaja.FieldByName('TICKER').AsString := 'RIVP-R-A';
      cdsPregledIzvjestaja.FieldByName('DATUMUNOSA').AsDateTime := Date;
      cdsPregledIzvjestaja.FieldByName('DATUMIZVJESTAJA').AsDateTime := Date;
      cdsPregledIzvjestaja.FieldByName('OPIS').AsString := ExtractFileName(IzvjestajPath);
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

procedure TFrmExcelReader.zatvoriGrid;
begin
  dbgExcel.Columns.Clear;
  ListBox1.Clear;
end;

end.
