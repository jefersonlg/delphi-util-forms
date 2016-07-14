object frmGeradorDeRelatorios: TfrmGeradorDeRelatorios
  Left = 0
  Top = 0
  BorderIcons = [biSystemMenu]
  BorderStyle = bsToolWindow
  Caption = '  Gerador de Relat'#243'rios'
  ClientHeight = 196
  ClientWidth = 328
  Color = clBtnFace
  Constraints.MaxHeight = 220
  Constraints.MinHeight = 220
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  OnClose = FormClose
  OnCreate = FormCreate
  OnResize = FormResize
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object panelProgressBar: TPanel
    AlignWithMargins = True
    Left = 3
    Top = 152
    Width = 322
    Height = 41
    Align = alBottom
    TabOrder = 0
    object lblProgressbar: TLabel
      Left = 1
      Top = 1
      Width = 320
      Height = 13
      Align = alTop
      Alignment = taCenter
      Caption = '0 %'
      ExplicitWidth = 20
    end
    object progressbar: TProgressBar
      AlignWithMargins = True
      Left = 4
      Top = 15
      Width = 314
      Height = 22
      Align = alBottom
      Smooth = True
      Step = 1
      TabOrder = 0
    end
  end
  object panel1: TPanel
    AlignWithMargins = True
    Left = 3
    Top = 3
    Width = 322
    Height = 134
    Align = alTop
    TabOrder = 1
    object labelArqDoc: TLabel
      AlignWithMargins = True
      Left = 4
      Top = 4
      Width = 314
      Height = 23
      Align = alTop
      Alignment = taCenter
      Caption = 'Modelo de Relat'#243'rio'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -19
      Font.Name = 'Tahoma'
      Font.Style = []
      ParentFont = False
      ExplicitWidth = 166
    end
    object labelArqDocSaveAs: TLabel
      AlignWithMargins = True
      Left = 4
      Top = 70
      Width = 314
      Height = 23
      Align = alTop
      Alignment = taCenter
      Caption = 'Relat'#243'rio a ser criado...'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -19
      Font.Name = 'Tahoma'
      Font.Style = []
      ParentFont = False
      ExplicitWidth = 195
    end
    object editArqDoc: TEdit
      AlignWithMargins = True
      Left = 11
      Top = 33
      Width = 300
      Height = 24
      Margins.Left = 10
      Margins.Right = 10
      Margins.Bottom = 10
      Align = alTop
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -13
      Font.Name = 'Tahoma'
      Font.Style = []
      ParentFont = False
      TabOrder = 0
      Text = '...'
    end
    object editArqDocSaveAs: TEdit
      AlignWithMargins = True
      Left = 11
      Top = 99
      Width = 300
      Height = 24
      Margins.Left = 10
      Margins.Right = 10
      Align = alTop
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -13
      Font.Name = 'Tahoma'
      Font.Style = []
      ParentFont = False
      TabOrder = 1
      Text = '...'
    end
  end
  object panel2: TPanel
    AlignWithMargins = True
    Left = 3
    Top = 143
    Width = 322
    Height = 4
    Align = alTop
    TabOrder = 2
    object btnGerarRelatorio: TButton
      AlignWithMargins = True
      Left = 11
      Top = 11
      Width = 150
      Height = 0
      Margins.Left = 10
      Margins.Top = 10
      Margins.Right = 0
      Margins.Bottom = 10
      Align = alLeft
      Caption = 'Gerar Relat'#243'rio'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -19
      Font.Name = 'Tahoma'
      Font.Style = []
      ParentFont = False
      TabOrder = 0
    end
    object btnCancelar: TButton
      AlignWithMargins = True
      Left = 171
      Top = 11
      Width = 150
      Height = 0
      Margins.Left = 10
      Margins.Top = 10
      Margins.Right = 0
      Margins.Bottom = 10
      Align = alLeft
      Caption = 'Cancelar'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -19
      Font.Name = 'Tahoma'
      Font.Style = []
      ParentFont = False
      TabOrder = 1
    end
  end
  object dtsRelatorio: TSQLDataSet
    CommandText = 'select '#39'Gerador de Relat'#243'rio'#39
    MaxBlobSize = -1
    Params = <>
    SQLConnection = Connection
    Left = 83
    Top = 59
  end
  object dsRelatorio: TDataSource
    AutoEdit = False
    DataSet = cdsRelatorio
    Left = 43
    Top = 59
  end
  object cdsRelatorio: TClientDataSet
    Aggregates = <>
    Params = <>
    ProviderName = 'dspRelatorio'
    ReadOnly = True
    Left = 163
    Top = 59
  end
  object dspRelatorio: TDataSetProvider
    DataSet = dtsRelatorio
    Left = 123
    Top = 59
  end
  object Connection: TSQLConnection
    LoginPrompt = False
    Left = 44
    Top = 16
  end
  object query: TSQLQuery
    MaxBlobSize = -1
    Params = <>
    SQL.Strings = (
      'select time(now)')
    SQLConnection = Connection
    Left = 43
    Top = 107
  end
end
