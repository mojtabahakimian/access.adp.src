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
    Width =15090
    DatasheetFontHeight =10
    ItemSuffix =4
    Left =405
    Top =390
    Right =11520
    Bottom =5775
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x51cde666536ae240
    End
    Caption ="ليست چكهاي دريافتي نزد بانك"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Arial (Arabic)"
    OnError ="[Event Procedure]"
    Begin
        Begin Label
            BackStyle =0
            TextFontCharSet =178
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
        Begin Section
            CanGrow = NotDefault
            Height =9240
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin Subform
                    OverlapFlags =85
                    Left =90
                    Top =120
                    Width =15000
                    Height =8505
                    Name ="CHEK_VLIST"
                    SourceObject ="Form.CHEK_VLIST"
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1417
                    Top =8731
                    Width =2539
                    Height =315
                    FontWeight =700
                    TabIndex =1
                    Name ="Text2"
                    ControlSource ="=[CHEK_VLIST].[Form]![smabl]"
                    Format ="#,###"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =4035
                            Top =8730
                            Width =1260
                            Height =300
                            FontWeight =700
                            Name ="Label3"
                            Caption ="جمع مبلغ چكها:"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "CHEK_VOSUL_LES.cls"
