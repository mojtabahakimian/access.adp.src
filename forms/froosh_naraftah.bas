Version =20
VersionRequired =20
Begin Form
    AllowDeletions = NotDefault
    AllowAdditions = NotDefault
    OrderByOn = NotDefault
    AllowEdits = NotDefault
    DefaultView =2
    AllowUpdating =4
    ViewsAllowed =2
    TabularCharSet =178
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =4940
    RowHeight =360
    ItemSuffix =22
    Left =270
    Top =210
    Right =16395
    Bottom =9210
    DatasheetForeColor =33554432
    DatasheetGridlinesColor =12632256
    MaxRecords =1000000
    OrderBy ="froosh_naraftah.MANDAH DESC"
    RecSrcDt = Begin
        0xe132591f93bfe440
    End
    RecordSource ="SELECT STUF_DEF.CODE, STUF_DEF.NAME, STUF_DEF.N_FANI, STUF_DEF.TOZIH, STUF_DEF.B"
        "_SEF, STUF_DEF.N_SEF, STUF_DEF.MIN_M, STUF_DEF.MAX_M, STUF_DEF.MABL_F, MOGUDI_KO"
        "L_ANBARHA.MANDAH, TCOD_STUFGROUP.NAMES FROM STUF_DEF INNER JOIN MOGUDI_KOL_ANBAR"
        "HA ON STUF_DEF.CODE = MOGUDI_KOL_ANBARHA.CODE INNER JOIN TCOD_STUFGROUP ON STUF_"
        "DEF.RADAH = TCOD_STUFGROUP.CODE WHERE (NOT (STUF_DEF.CODE IN (SELECT CODE FROM d"
        "bo.INVO_LST WHERE (TAG = 2) GROUP BY CODE))) ORDER BY TCOD_STUFGROUP.NAMES, STUF"
        "_DEF.NAME"
    Caption ="ليست كالاهاي فروش نرفته (راكد)"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Tahoma"
    AllowFormView =0
    RecordSourceQualifier ="dbo"
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
        Begin Section
            Height =4308
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =1
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1927
                    Top =113
                    Width =2490
                    ColumnWidth =705
                    ColumnOrder =0
                    Name ="CODE"
                    ControlSource ="CODE"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =113
                            Width =510
                            Height =240
                            Name ="Label1"
                            Caption ="كد كالا"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =1
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1927
                    Top =453
                    Width =2490
                    ColumnWidth =6135
                    ColumnOrder =2
                    TabIndex =1
                    Name ="NAME"
                    ControlSource ="NAME"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =450
                            Width =540
                            Height =240
                            Name ="Label3"
                            Caption ="نام كالا"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =1
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1927
                    Top =793
                    Width =2490
                    ColumnWidth =1965
                    ColumnOrder =9
                    TabIndex =2
                    Name ="N_FANI"
                    ControlSource ="N_FANI"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =793
                            Width =630
                            Height =240
                            Name ="Label5"
                            Caption ="باركد"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =1
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1927
                    Top =1133
                    Width =2490
                    ColumnOrder =10
                    TabIndex =3
                    Name ="TOZIH"
                    ControlSource ="TOZIH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =1140
                            Width =675
                            Height =240
                            Name ="Label7"
                            Caption ="توضيحات"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =1
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1927
                    Top =1474
                    ColumnOrder =5
                    TabIndex =4
                    Name ="B_SEF"
                    ControlSource ="B_SEF"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =1470
                            Width =690
                            Height =240
                            Name ="Label9"
                            Caption ="في خرده"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =1
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1927
                    Top =1814
                    ColumnOrder =6
                    TabIndex =5
                    Name ="N_SEF"
                    ControlSource ="N_SEF"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =1815
                            Width =1020
                            Height =240
                            Name ="Label11"
                            Caption ="نقطه سفارش"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =1
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1927
                    Top =2154
                    ColumnOrder =7
                    TabIndex =6
                    Name ="MIN_M"
                    ControlSource ="MIN_M"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =2160
                            Width =1125
                            Height =240
                            Name ="Label13"
                            Caption ="حداقل موجودي"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =1
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1927
                    Top =2494
                    ColumnWidth =1680
                    ColumnOrder =8
                    TabIndex =7
                    Name ="MAX_M"
                    ControlSource ="MAX_M"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =2490
                            Width =1365
                            Height =240
                            Name ="Label15"
                            Caption ="قيمت مصرف كننده"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =1
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1927
                    Top =2834
                    ColumnOrder =4
                    TabIndex =8
                    Name ="MABL_F"
                    ControlSource ="MABL_F"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =2835
                            Width =1200
                            Height =240
                            Name ="Label17"
                            Caption ="في عمده فروش"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =1
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1927
                    Top =3195
                    ColumnWidth =750
                    ColumnOrder =3
                    TabIndex =9
                    Name ="MANDAH"
                    ControlSource ="MANDAH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =3196
                            Width =1200
                            Height =240
                            Name ="Label19"
                            Caption ="موجودي"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =1
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1927
                    Top =3556
                    ColumnOrder =1
                    TabIndex =10
                    Name ="NAMES"
                    ControlSource ="NAMES"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =3557
                            Width =1200
                            Height =240
                            Name ="Label21"
                            Caption ="گروه"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "froosh_naraftah.cls"
