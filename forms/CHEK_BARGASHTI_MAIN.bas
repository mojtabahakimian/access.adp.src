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
    Width =15080
    DatasheetFontHeight =10
    ItemSuffix =4
    Left =405
    Right =11520
    Bottom =5385
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x011cde3db0a6e240
    End
    Caption ="ليست چكهاي برگشتي"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Arial (Arabic)"
    OnError ="[Event Procedure]"
    AllowDatasheetView =0
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
            Height =8787
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin Subform
                    OverlapFlags =85
                    Left =56
                    Width =14970
                    Height =8325
                    Name ="CHEK_BARGASHTI"
                    SourceObject ="Form.CHEK_BARGASHTI"
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2267
                    Top =8448
                    Width =2539
                    Height =315
                    TabIndex =1
                    Name ="Text2"
                    ControlSource ="=[CHEK_BARGASHTI].[Form]![smabl]"
                    Format ="#,###"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =4885
                            Top =8447
                            Width =1260
                            Height =300
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
' See "CHEK_BARGASHTI_MAIN.cls"
