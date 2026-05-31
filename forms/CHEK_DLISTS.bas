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
    Width =15307
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
    Caption ="ليست چكهاي دريافتي"
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
            Height =9354
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin Subform
                    OverlapFlags =85
                    Left =100
                    Top =105
                    Width =15030
                    Height =8745
                    Name ="CHKE_DLIST"
                    SourceObject ="Form.CHKE_DLIST"
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1417
                    Top =8958
                    Width =2539
                    Height =315
                    TabIndex =1
                    Name ="Text2"
                    ControlSource ="=[CHKE_DLIST].[Form]![smabl]"
                    Format ="#,###"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =4035
                            Top =8957
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
' See "CHEK_DLISTS.cls"
