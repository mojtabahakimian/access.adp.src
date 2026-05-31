Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    DefaultView =0
    AllowUpdating =4
    ScrollBars =0
    ViewsAllowed =1
    TabularCharSet =178
    TabularFamily =0
    BorderStyle =3
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =8390
    DatasheetFontHeight =10
    ItemSuffix =4
    Left =6000
    Top =1875
    Right =14385
    Bottom =4515
    DatasheetGridlinesColor =12632256
    MaxRecords =1000000
    RecSrcDt = Begin
        0x00d79ef8d357e440
    End
    Caption ="استعلام (توضيحات)"
    DatasheetFontName ="Arial"
    OnLoad ="[Event Procedure]"
    AllowDatasheetView =0
    Begin
        Begin Label
            BackStyle =0
            TextFontCharSet =178
            FontName ="Tahoma"
        End
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            OldBorderStyle =0
            TextFontCharSet =178
            Width =1701
            LabelX =-1701
            FontName ="Tahoma"
            AsianLineBreak =255
        End
        Begin ListBox
            SpecialEffect =2
            TextFontCharSet =178
            Width =1701
            Height =1417
            LabelX =-1701
            FontName ="Tahoma"
        End
        Begin Section
            Height =2654
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =215
                    TextAlign =3
                    KeyboardLanguage =43
                    IMESentenceMode =3
                    Left =56
                    Top =56
                    Width =8277
                    Height =2551
                    Name ="LIST0"
                    OnDblClick ="[Event Procedure]"
                    OnKeyPress ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =7835
                            Top =56
                            Width =555
                            Height =240
                            Name ="Label3"
                            Caption ="Text2:"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "ESTELAM.cls"
