Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    KeyPreview = NotDefault
    OrderByOn = NotDefault
    AllowEdits = NotDefault
    DefaultView =2
    AllowUpdating =4
    ViewsAllowed =2
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =10500
    RowHeight =330
    ItemSuffix =19
    Left =90
    Top =375
    Right =15030
    Bottom =7890
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x918f665bece5e340
    End
    RecordSource ="SELECT SUM(FROOSH_COUNTALL_CITY_N_FANI.MEGHKF) AS MEGHF, SUM(FROOSH_COUNTALL_CIT"
        "Y_N_FANI.MABK) AS MABKK, SUM(FROOSH_COUNTALL_CITY_N_FANI.MEGHkK) AS MEGHk, TCOD_"
        "OSTAN.OSNAME, FROOSH_COUNTALL_CITY_N_FANI.OSCO FROM FROOSH_COUNTALL_CITY_N_FANI("
        "1, 999999) FROOSH_COUNTALL_CITY_N_FANI INNER JOIN TCOD_OSTAN ON FROOSH_COUNTALL_"
        "CITY_N_FANI.OSCO = TCOD_OSTAN.OSCODE GROUP BY TCOD_OSTAN.OSNAME, FROOSH_COUNTALL"
        "_CITY_N_FANI.OSCO ORDER BY SUM(FROOSH_COUNTALL_CITY_N_FANI.MEGHKF)"
    Caption ="ليست فروش روزانه به تفكيك نوع كالا"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Tahoma"
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
            Height =7871
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
                    Left =5705
                    Top =113
                    ColumnOrder =2
                    Name ="MEGHk"
                    ControlSource ="MEGHk"
                    Format ="Standard"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3891
                            Top =113
                            Width =570
                            Height =240
                            Name ="Label1"
                            Caption ="متراژ"
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
                    Left =5705
                    Top =793
                    Width =2490
                    ColumnWidth =4425
                    ColumnOrder =1
                    TabIndex =1
                    Name ="OSNAME"
                    ControlSource ="OSNAME"
                    OnDblClick ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3900
                            Top =795
                            Width =585
                            Height =240
                            Name ="Label5"
                            Caption ="استان*"
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
                    Left =5705
                    Top =1133
                    ColumnWidth =1065
                    TabIndex =2
                    BackColor =13434828
                    Name ="MEGHF"
                    ControlSource ="MEGHF"
                    OnDblClick ="[Event Procedure]"
                    ConditionalFormat = Begin
                        0x0100000066000000010000000000000004000000000000000200000001000000 ,
                        0x00000000ccffcc00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x300000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3891
                            Top =1133
                            Width =960
                            Height =240
                            Name ="Label7"
                            Caption ="كارتن*"
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
                    Left =5705
                    Top =1493
                    TabIndex =3
                    Name ="MABKK"
                    ControlSource ="MABKK"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3891
                            Top =1493
                            Width =960
                            Height =240
                            Name ="Label15"
                            Caption ="مبلغ"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =5640
                    Top =1814
                    TabIndex =4
                    Name ="OSCO"
                    ControlSource ="OSCO"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4037
                            Top =1870
                            Width =705
                            Height =240
                            Name ="Label16"
                            Caption ="كد استان"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "FROOSH_COUNTALL_CITY1_OSTAN.cls"
