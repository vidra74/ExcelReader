unit untExcelReader;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Grids, DBGrids, ExtCtrls, DB, ADODB;

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
    procedure btnZatvoriClick(Sender: TObject);
    procedure btnOtvoriExcelClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure btnOtvoriSheetClick(Sender: TObject);
  private
    { Private declarations }
    function otvoriOdabraniSheet(Sheet: String): Boolean;
    procedure posaljiSheetUGrid;
    procedure zatvoriGrid;
    procedure GetFieldInfo;
  public
    { Public declarations }
  end;

var
  FrmExcelReader: TFrmExcelReader;

implementation

{$R *.dfm}

uses typinfo;

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
  except
    On E:Exception do
      ShowMessage('adoconnExcel.Open : ' + E.Message);

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