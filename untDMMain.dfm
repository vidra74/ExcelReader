object DMMain: TDMMain
  OldCreateOrder = False
  Height = 386
  Width = 616
  object adoConectExcel: TADOConnection
    ConnectionString = 
      'Provider=Microsoft.Jet.OLEDB.4.0;Password="";Data Source=D:\Drop' +
      'box\Trziste\TSHC\TSHC-fin2011-1Y-NotREV-N-HR.xls;Persist Securit' +
      'y Info=True'
    KeepConnection = False
    LoginPrompt = False
    Provider = 'Microsoft.Jet.OLEDB.4.0'
    Left = 16
    Top = 16
  end
end
