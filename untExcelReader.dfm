object FrmExcelReader: TFrmExcelReader
  Left = 0
  Top = 0
  Caption = 'Excel Reader'
  ClientHeight = 750
  ClientWidth = 1000
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  Menu = menExcelReader
  OldCreateOrder = False
  Position = poDesktopCenter
  OnClose = FormClose
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object pnlDohvat: TPanel
    Left = 0
    Top = 0
    Width = 1000
    Height = 41
    Align = alTop
    TabOrder = 0
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
    Width = 880
    Height = 593
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
    Height = 593
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
    object btnSpremiIzvjestaj: TButton
      Left = 10
      Top = 68
      Width = 100
      Height = 25
      Caption = 'Spremi Izvje'#353'taj'
      TabOrder = 3
      OnClick = btnSpremiIzvjestajClick
    end
  end
  object ListBox1: TListBox
    Left = 0
    Top = 634
    Width = 1000
    Height = 97
    Align = alBottom
    ItemHeight = 13
    TabOrder = 3
  end
  object sbExcelStatus: TStatusBar
    Left = 0
    Top = 731
    Width = 1000
    Height = 19
    Panels = <
      item
        Text = 'Nije odabrana Excel datoteka'
        Width = 400
      end
      item
        Width = 50
      end
      item
        Width = 400
      end>
  end
  object dlgOpenExcel: TOpenDialog
    Filter = 'Excel|*.xls|Novi Excel|*.xlsx'
    Left = 208
    Top = 128
  end
  object dsExcel: TDataSource
    DataSet = DMMain.qryExcel
    Left = 208
    Top = 208
  end
  object menExcelReader: TMainMenu
    Left = 208
    Top = 160
    object menuFile: TMenuItem
      Caption = '&File'
      object miOpenExcelReport: TMenuItem
        Caption = 'Open Excel Report'
        OnClick = btnOtvoriExcelClick
      end
      object miSaveReportInfo: TMenuItem
        Caption = 'Save report Info'
        OnClick = btnSpremiIzvjestajClick
      end
      object miCloseReport: TMenuItem
        Caption = 'Close Report'
      end
      object N1: TMenuItem
        Caption = '-'
      end
      object miExit: TMenuItem
        Caption = 'Exit'
        OnClick = btnZatvoriClick
      end
    end
  end
end
