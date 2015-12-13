object FrmExcelReader: TFrmExcelReader
  Left = 0
  Top = 0
  Caption = 'Excel Reader'
  ClientHeight = 489
  ClientWidth = 700
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Position = poDesktopCenter
  OnClose = FormClose
  PixelsPerInch = 96
  TextHeight = 13
  object pnlDohvat: TPanel
    Left = 0
    Top = 0
    Width = 700
    Height = 41
    Align = alTop
    TabOrder = 0
    object lblExcelDatoteka: TLabel
      Left = 10
      Top = 8
      Width = 141
      Height = 13
      Caption = 'Nije odabrana Excel datoteka'
    end
    object cboxExcelSheets: TComboBox
      Left = 480
      Top = 5
      Width = 145
      Height = 21
      TabOrder = 0
      Text = 'Excel sheets'
    end
  end
  object dbgExcel: TDBGrid
    Left = 120
    Top = 41
    Width = 580
    Height = 351
    Align = alClient
    DataSource = dsExcel
    TabOrder = 1
    TitleFont.Charset = DEFAULT_CHARSET
    TitleFont.Color = clWindowText
    TitleFont.Height = -11
    TitleFont.Name = 'Tahoma'
    TitleFont.Style = []
  end
  object pnlBotuni: TPanel
    Left = 0
    Top = 41
    Width = 120
    Height = 351
    Align = alLeft
    TabOrder = 2
    object btnZatvori: TButton
      Left = 10
      Top = 224
      Width = 100
      Height = 25
      Caption = 'Zatvori'
      TabOrder = 0
      OnClick = btnZatvoriClick
    end
    object btnOtvoriExcel: TButton
      Left = 10
      Top = 6
      Width = 100
      Height = 25
      Caption = 'Otvori Excel'
      TabOrder = 1
      OnClick = btnOtvoriExcelClick
    end
    object btnOtvoriSheet: TButton
      Left = 10
      Top = 37
      Width = 100
      Height = 25
      Caption = 'Otvori Sheet'
      TabOrder = 2
      OnClick = btnOtvoriSheetClick
    end
  end
  object ListBox1: TListBox
    Left = 0
    Top = 392
    Width = 700
    Height = 97
    Align = alBottom
    ItemHeight = 13
    TabOrder = 3
  end
  object dlgOpenExcel: TOpenDialog
    Filter = 'Excel|*.xls|Novi Excel|*.xlsx'
    Left = 16
    Top = 8
  end
  object adoconectExcel: TADOConnection
    ConnectionString = 
      'Provider=Microsoft.Jet.OLEDB.4.0;Password="";Data Source=D:\Drop' +
      'box\Trziste\TSHC\TSHC-fin2011-1Y-NotREV-N-HR.xls;Persist Securit' +
      'y Info=True'
    KeepConnection = False
    LoginPrompt = False
    Provider = 'Microsoft.Jet.OLEDB.4.0'
    Left = 376
    Top = 8
  end
  object qryExcel: TADOQuery
    Connection = adoconectExcel
    EnableBCD = False
    Parameters = <>
    Left = 448
    Top = 8
  end
  object dsExcel: TDataSource
    DataSet = qryExcel
    Left = 520
    Top = 8
  end
end
