object Form2: TForm2
  Left = 0
  Top = 0
  Caption = 'Form2'
  ClientHeight = 1338
  ClientWidth = 1936
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -25
  Font.Name = 'Tahoma'
  Font.Style = []
  OnCreate = FormCreate
  OnDestroy = FormDestroy
  PixelsPerInch = 240
  TextHeight = 30
  object Memo1: TMemo
    Left = 32
    Top = 128
    Width = 1186
    Height = 1170
    Margins.Left = 6
    Margins.Top = 6
    Margins.Right = 6
    Margins.Bottom = 6
    Lines.Strings = (
      'Memo1')
    TabOrder = 0
  end
  object Button1: TButton
    Left = 16
    Top = 16
    Width = 150
    Height = 50
    Margins.Left = 6
    Margins.Top = 6
    Margins.Right = 6
    Margins.Bottom = 6
    Caption = 'Button1'
    TabOrder = 1
    OnClick = Button1Click
  end
  object Button2: TButton
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
    OnClick = Button2Click
  end
  object Button3: TButton
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
    OnClick = Button3Click
  end
  object Button4: TButton
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
    OnClick = Button4Click
  end
  object Button5: TButton
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
    OnClick = Button5Click
  end
  object Button6: TButton
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
    OnClick = Button6Click
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
  object FDTable1: TFDTable
    Connection = FDConnection1
    TableName = 'dbispconfig.mail_access'
    Left = 1624
    Top = 608
  end
end
