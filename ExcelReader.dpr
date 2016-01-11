program ExcelReader;

uses
  Forms,
  untExcelReader in 'untExcelReader.pas' {FrmExcelReader},
  untDMMain in 'untDMMain.pas' {DMMain: TDataModule},
  untObjReport in 'untObjReport.pas',
  untObjReportList in 'untObjReportList.pas',
  untObjReportSheet in 'untObjReportSheet.pas',
  untObjBilancaList in 'untObjBilancaList.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TFrmExcelReader, FrmExcelReader);
  Application.CreateForm(TDMMain, DMMain);
  Application.Run;
end.
