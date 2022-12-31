object frmMain: TfrmMain
  Left = 0
  Top = 0
  Caption = 'Read Email Settings'
  ClientHeight = 1056
  ClientWidth = 1924
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -25
  Font.Name = 'Tahoma'
  Font.Style = []
  OnCreate = FormCreate
  OnDestroy = FormDestroy
  PixelsPerInch = 240
  DesignSize = (
    1924
    1056)
  TextHeight = 30
  object mmoLog: TMemo
    Left = 52
    Top = 102
    Width = 1186
    Height = 908
    Margins.Left = 6
    Margins.Top = 6
    Margins.Right = 6
    Margins.Bottom = 6
    Anchors = [akLeft, akTop, akBottom]
    Lines.Strings = (
      'Memo1')
    ScrollBars = ssBoth
    TabOrder = 0
  end
  object btnEmailInfoTest: TButton
    Left = 1328
    Top = 11
    Width = 274
    Height = 50
    Margins.Left = 6
    Margins.Top = 6
    Margins.Right = 6
    Margins.Bottom = 6
    Caption = 'Email Info Test'
    TabOrder = 1
    OnClick = btnEmailInfoTestClick
  end
  object btnOutputProfileTest: TButton
    Left = 1328
    Top = 96
    Width = 274
    Height = 50
    Margins.Left = 6
    Margins.Top = 6
    Margins.Right = 6
    Margins.Bottom = 6
    Caption = 'outlook profile test'
    TabOrder = 2
    OnClick = btnOutputProfileTestClick
  end
  object btnWindowsProfileTest: TButton
    Left = 1328
    Top = 176
    Width = 274
    Height = 50
    Margins.Left = 6
    Margins.Top = 6
    Margins.Right = 6
    Margins.Bottom = 6
    Caption = 'windows profile test'
    TabOrder = 3
    OnClick = btnWindowsProfileTestClick
  end
  object btnLoadRegistryFile: TButton
    Left = 1328
    Top = 272
    Width = 274
    Height = 50
    Margins.Left = 6
    Margins.Top = 6
    Margins.Right = 6
    Margins.Bottom = 6
    Caption = 'test load registry file'
    TabOrder = 4
    OnClick = btnLoadRegistryFileClick
  end
  object btnExtractCDKeys: TButton
    Left = 1328
    Top = 368
    Width = 274
    Height = 50
    Margins.Left = 6
    Margins.Top = 6
    Margins.Right = 6
    Margins.Bottom = 6
    Caption = 'Extract CDKeys'
    TabOrder = 5
    OnClick = btnExtractCDKeysClick
  end
  object btnGetWindowsUsername: TButton
    Left = 1328
    Top = 480
    Width = 418
    Height = 50
    Margins.Left = 6
    Margins.Top = 6
    Margins.Right = 6
    Margins.Bottom = 6
    Caption = 'Get Windows Domain\Username'
    TabOrder = 6
    OnClick = btnGetWindowsUsernameClick
  end
  object FDConnection1: TFDConnection
    Params.Strings = (
      'CharacterSet=utf8mb4'
      'DriverID=MySQL')
    LoginPrompt = False
    Left = 1368
    Top = 616
  end
  object FDPhysMySQLDriverLink1: TFDPhysMySQLDriverLink
    Left = 1312
    Top = 650
  end
  object tblMailAccess: TFDTable
    Connection = FDConnection1
    TableName = 'dbispconfig.mail_access'
    Left = 1624
    Top = 608
  end
end
