Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    DefaultView =0
    AllowUpdating =4
    ScrollBars =0
    ViewsAllowed =1
    TabularCharSet =178
    TabularFamily =42
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =14570
    DatasheetFontHeight =10
    ItemSuffix =10
    Left =120
    Top =390
    Right =11640
    Bottom =6165
    HelpContextId =8016
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x4852978af96be340
    End
    Caption ="ليست چكهاي به حساب گذاشته شده"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    HelpFile ="dena.hlp"
    DatasheetFontName ="Arial (Arabic)"
    OnLoad ="[Event Procedure]"
    OnError ="[Event Procedure]"
    AllowDatasheetView =0
    Begin
        Begin Label
            BackStyle =0
            TextFontCharSet =178
            FontName ="Tahoma"
        End
        Begin CommandButton
            TextFontCharSet =178
            Width =1701
            Height =283
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Tahoma"
        End
        Begin TextBox
            SpecialEffect =2
            OldBorderStyle =0
            TextFontCharSet =178
            Width =1701
            LabelX =-1701
            FontName ="Tahoma"
        End
        Begin Subform
            SpecialEffect =2
            Width =1701
            Height =1701
        End
        Begin CustomControl
            SpecialEffect =2
            Width =4536
            Height =2835
        End
        Begin Section
            CanGrow = NotDefault
            Height =9127
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin Subform
                    OverlapFlags =247
                    Left =113
                    Top =453
                    Width =14445
                    Height =7590
                    Name ="CHK_V_PRINTSUB"
                    SourceObject ="Form.CHK_V_PRINTSUB"
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =340
                    Top =8675
                    Width =2539
                    Height =315
                    ColumnOrder =1
                    TabIndex =1
                    Name ="Text2"
                    ControlSource ="=[CHK_V_PRINTSUB].[Form]![smabl]"
                    Format ="#,###"
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =2958
                            Top =8674
                            Width =1260
                            Height =300
                            Name ="Label3"
                            Caption ="جمع مبلغ چكها:"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =93
                    AccessKey =141
                    TextFontCharSet =0
                    Left =3622
                    Top =75
                    Width =1170
                    Height =555
                    FontWeight =700
                    TabIndex =2
                    Name ="Command4"
                    Caption ="&چاپ برگه1"
                    OnClick ="[Event Procedure]"
                    ObjectPalette = Begin
                        0x0003100000000000800000000080000080800000000080008000800000808000 ,
                        0x80808000c0c0c000ff000000c0c0c000ffff00000000ff00c0c0c00000ffff00 ,
                        0xffffff0000000000
                    End
                    ControlTipText ="Find Next"
                End
                Begin CommandButton
                    OverlapFlags =93
                    AccessKey =200
                    TextFontCharSet =0
                    Left =4792
                    Top =75
                    Width =1290
                    Height =555
                    FontWeight =700
                    TabIndex =3
                    Name ="Command5"
                    Caption ="چاپ &برگه2"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Find Next"
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =93
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =11442
                    Top =226
                    ColumnOrder =2
                    TabIndex =5
                    Name ="CHK"
                    DefaultValue ="0"
                End
                Begin TextBox
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =6340
                    Top =226
                    Width =2256
                    Height =315
                    ColumnOrder =0
                    TabIndex =4
                    Name ="CITY"
                    DefaultValue ="\"يزد\""
                End
                Begin Label
                    OverlapFlags =93
                    TextFontCharSet =0
                    Left =8662
                    Top =225
                    Width =1755
                    Height =300
                    Name ="Label9"
                    Caption ="شهرستان محل پرداخت:"
                End
            End
        End
    End
End
CodeBehindForm
' See "CHK_V_PRINT.cls"
