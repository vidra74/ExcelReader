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
    Left = 32
    Top = 16
  end
  object cdsPregledIzvjestaja: TClientDataSet
    Aggregates = <>
    Params = <>
    Left = 32
    Top = 72
    object cdsPregledIzvjestajaID: TIntegerField
      FieldName = 'ID'
    end
    object cdsPregledIzvjestajaPATH: TStringField
      FieldName = 'PATH'
      Size = 100
    end
    object cdsPregledIzvjestajaOPIS: TStringField
      FieldName = 'OPIS'
      Size = 200
    end
    object cdsPregledIzvjestajaTICKER: TStringField
      FieldName = 'TICKER'
      Size = 10
    end
    object cdsPregledIzvjestajaDATUMUNOSA: TDateField
      FieldName = 'DATUMUNOSA'
    end
    object cdsPregledIzvjestajaDATUMIZVJESTAJA: TDateField
      FieldName = 'DATUMIZVJESTAJA'
    end
  end
  object qryIzvjestajPodaci: TADOQuery
    Connection = adoConectExcel
    EnableBCD = False
    Parameters = <>
    Left = 120
    Top = 16
  end
  object qryExcel: TADOQuery
    Connection = adoConectExcel
    EnableBCD = False
    Parameters = <>
    Left = 148
    Top = 16
  end
  object cdsBilanca: TClientDataSet
    Aggregates = <>
    Params = <>
    Left = 32
    Top = 128
    object cdsBilancaID_REPORT: TIntegerField
      FieldName = 'ID_REPORT'
    end
    object cdsBilancaID_SHEET: TIntegerField
      FieldName = 'ID_SHEET'
    end
    object cdsBilancaAOP: TIntegerField
      FieldName = 'AOP'
    end
    object cdsBilancaPRET_TROM: TCurrencyField
      FieldName = 'PRET_TROM'
    end
    object cdsBilancaPRET_TOT: TCurrencyField
      FieldName = 'PRET_TOT'
    end
    object cdsBilancaTREN_TROM: TCurrencyField
      FieldName = 'TREN_TROM'
    end
    object cdsBilancaTREN_TOT: TCurrencyField
      FieldName = 'TREN_TOT'
    end
    object cdsBilancaVRIJEME: TDateTimeField
      FieldName = 'VRIJEME'
    end
  end
end
