unit untDMMain;

interface

uses
  SysUtils, Classes, DB, ADODB;

type
  TDMMain = class(TDataModule)
    adoConectExcel: TADOConnection;
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
