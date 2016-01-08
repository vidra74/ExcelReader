unit untDMMain;

interface

uses
  SysUtils, Classes, DB, ADODB, DBClient;

type
  TDMMain = class(TDataModule)
    adoConectExcel: TADOConnection;
    cdsPregledIzvjestaja: TClientDataSet;
    cdsPregledIzvjestajaID: TIntegerField;
    cdsPregledIzvjestajaPATH: TStringField;
    cdsPregledIzvjestajaOPIS: TStringField;
    cdsPregledIzvjestajaTICKER: TStringField;
    cdsPregledIzvjestajaDATUMUNOSA: TDateField;
    cdsPregledIzvjestajaDATUMIZVJESTAJA: TDateField;
    qryIzvjestajPodaci: TADOQuery;
    qryExcel: TADOQuery;
    procedure DataModuleDestroy(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  DMMain: TDMMain;

implementation

{$R *.dfm}

procedure TDMMain.DataModuleDestroy(Sender: TObject);
begin
  qryIzvjestajPodaci.Close;
  qryExcel.Close;
  cdsPregledIzvjestaja.Close;
  adoConectExcel.Close;
end;

end.
