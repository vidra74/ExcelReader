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
    cdsBilanca: TClientDataSet;
    cdsBilancaID_REPORT: TIntegerField;
    cdsBilancaID_SHEET: TIntegerField;
    cdsBilancaAOP: TIntegerField;
    cdsBilancaPRET_TROM: TCurrencyField;
    cdsBilancaPRET_TOT: TCurrencyField;
    cdsBilancaTREN_TROM: TCurrencyField;
    cdsBilancaTREN_TOT: TCurrencyField;
    cdsBilancaVRIJEME: TDateTimeField;
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
