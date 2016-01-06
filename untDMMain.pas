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
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  DMMain: TDMMain;

implementation

{$R *.dfm}

end.
