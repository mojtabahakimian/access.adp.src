Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    KeyPreview = NotDefault
    AllowEdits = NotDefault
    DefaultView =2
    AllowUpdating =4
    ViewsAllowed =2
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =7597
    RowHeight =255
    ItemSuffix =16
    Left =780
    Top =60
    Right =14565
    Bottom =7350
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x3295cdaecbe4e340
    End
    RecordSource ="SELECT SUM(MEGHKF) AS MEGHF, MPNAME, SUM(MABK) AS MABKK, SUM(MEGHkK) AS MEGHk FR"
        "OM FROOSH_COUNTALL_CITY_N_FANI(1, 999999) FROOSH_COUNTALL_CITY_N_FANI GROUP BY M"
        "PNAME ORDER BY MPNAME"
    Caption ="ليست فروش روزانه به تفكيك نوع كالا"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Tahoma"
    OnKeyPress ="[Event Procedure]"
    AllowFormView =0
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
        End
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            OldBorderStyle =0
            Width =1701
            LabelX =-1701
            FontName ="Tahoma"
            AsianLineBreak =255
        End
        Begin Subform
            SpecialEffect =2
            Width =1701
            Height =1701
        End
        Begin Section
            CanGrow = NotDefault
            Height =5215
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =4591
                    Top =113
                    ColumnOrder =2
                    Name ="MEGHk"
                    ControlSource ="MEGHk"
                    OnDblClick ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2777
                            Top =113
                            Width =570
                            Height =240
                            Name ="Label1"
                            Caption ="متراژ*"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =4591
                    Top =793
                    Width =2490
                    ColumnWidth =4425
                    ColumnOrder =1
                    TabIndex =1
                    Name ="MPNAME"
                    ControlSource ="MPNAME"
                    OnDblClick ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2784
                            Top =795
                            Width =540
                            Height =240
                            Name ="Label5"
                            Caption ="شهر*"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =4591
                    Top =1133
                    ColumnWidth =1065
                    TabIndex =2
                    Name ="MEGHF"
                    ControlSource ="MEGHF"
                    OnDblClick ="[Event Procedure]"
                    OnClick ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2777
                            Top =1133
                            Width =960
                            Height =240
                            Name ="Label7"
                            Caption ="كارتن*"
                        End
                    End
                End
                Begin Subform
                    OverlapFlags =87
                    Top =3231
                    Width =6981
                    TabIndex =3
                    Name ="FROOSH_COUNTALL_CITY2"
                    SourceObject ="Form.FROOSH_COUNTALL_CITY2"
                    LinkChildFields ="MPNAME"
                    LinkMasterFields ="MPNAME"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =1041
                            Top =2991
                            Width =2100
                            Height =240
                            Name ="FROOSH_COUNTALL_CITY2 Label"
                            Caption ="FROOSH_COUNTALL_CITY2"
                            EventProcPrefix ="FROOSH_COUNTALL_CITY2_Label"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =4591
                    Top =1493
                    TabIndex =4
                    Name ="MABKK"
                    ControlSource ="MABKK"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2777
                            Top =1493
                            Width =960
                            Height =240
                            Name ="Label15"
                            Caption ="مبلغ"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "FROOSH_COUNTALL_CITY1.cls"
