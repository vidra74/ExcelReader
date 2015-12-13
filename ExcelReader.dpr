program ExcelReader;

uses
  Forms,
  untExcelReader in 'untExcelReader.pas' {FrmExcelReader};

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TFrmExcelReader, FrmExcelReader);
  Application.Run;
end.
