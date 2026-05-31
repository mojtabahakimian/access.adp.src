Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    AllowEdits = NotDefault
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
    Width =14966
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
    Caption ="ليست چكهاي پرداختي"
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
            Height =9751
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin Subform
                    OverlapFlags =85
                    Left =113
                    Top =118
                    Width =14745
                    Height =9120
                    Name ="CHEK_PLIST"
                    SourceObject ="Form.CHEK_PLIST"
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1757
                    Top =9298
                    Width =2539
                    Height =315
                    TabIndex =1
                    Name ="Text2"
                    ControlSource ="=[CHEK_PLIST].[Form]![smabl]"
                    Format ="#,###"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =4375
                            Top =9297
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
' See "CHEK_PLISTS.cls"
