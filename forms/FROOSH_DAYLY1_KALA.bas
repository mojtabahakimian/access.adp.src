Version =20
VersionRequired =20
Begin Form
    DefaultView =2
    AllowUpdating =4
    ViewsAllowed =2
    TabularCharSet =178
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =7036
    RowHeight =330
    ItemSuffix =23
    Left =825
    Top =930
    Right =14850
    Bottom =7890
    DatasheetForeColor =33554432
    DatasheetGridlinesColor =12632256
    MaxRecords =10000000
    RecSrcDt = Begin
        0xf19472456ee8e340
    End
    RecordSource ="SELECT HEAD_LST.NUMBER, HEAD_LST.NUMBER1, HEAD_LST.CUST_NO, HEAD_LST.SADER, SUM("
        "INVO_LST.MEGH) AS MEGHs, SUM(INVO_LST.MEGHk) AS MEGHks, SUM(INVO_LST.TOTALARZ) A"
        "S TOTALARZs, HEAD_LST.DATE_N, HEAD_LST_1.DATE_N AS DATE_Nh, CUST_HESAB.NAME, CUS"
        "T_HESAB.hes, HEAD_LST.ANBARF, HEAD_LST_1.ANBAR, 1 AS kk, SUM(INVO_LST.MABL_K) AS"
        " MABL_KS, SUM(HEAD_LST.MBAA) AS MBAAs, SUM(HEAD_LST.TAKHFIF) AS TAKHFIFs, STUF_D"
        "EF.NAME AS MMNAME, HEAD_LST_1.MAS FROM HEAD_LST INNER JOIN INVO_LST ON HEAD_LST."
        "NUMBER = INVO_LST.NUMBER AND HEAD_LST.TAG - 11 = INVO_LST.TAG INNER JOIN HEAD_LS"
        "T HEAD_LST_1 ON INVO_LST.NUMBER = HEAD_LST_1.NUMBER AND INVO_LST.TAG = HEAD_LST_"
        "1.TAG INNER JOIN CUST_HESAB ON HEAD_LST.CUST_NO = CUST_HESAB.hes INNER JOIN STUF"
        "_DEF ON INVO_LST.CODE = STUF_DEF.CODE WHERE (HEAD_LST.TAG = 13) GROUP BY HEAD_LS"
        "T.NUMBER, HEAD_LST.NUMBER1, HEAD_LST.CUST_NO, HEAD_LST.SADER, HEAD_LST.DATE_N, H"
        "EAD_LST_1.DATE_N, CUST_HESAB.NAME, CUST_HESAB.hes, HEAD_LST.ANBARF, HEAD_LST_1.A"
        "NBAR, STUF_DEF.NAME, HEAD_LST_1.MAS"
    Caption ="فروش به تفكيك مشتري"
    DatasheetFontName ="Tahoma"
    AllowFormView =0
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
        Begin ComboBox
            SpecialEffect =2
            TextFontCharSet =178
            Width =1701
            LabelX =-1701
            FontName ="Tahoma"
        End
        Begin Section
            Height =6709
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1700
                    Top =623
                    ColumnWidth =1125
                    ColumnOrder =1
                    Name ="NUMBER"
                    ControlSource ="NUMBER"
                    OnDblClick ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3527
                            Top =623
                            Width =1020
                            Height =240
                            Name ="Label0"
                            Caption ="شماره حواله*"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1700
                    Top =963
                    ColumnWidth =1155
                    ColumnOrder =0
                    TabIndex =1
                    Name ="NUMBER1"
                    ControlSource ="NUMBER1"
                    OnDblClick ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3512
                            Top =963
                            Width =1050
                            Height =240
                            Name ="Label1"
                            Caption ="شماره فاكتور*"
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
                    Left =1700
                    Top =1303
                    ColumnWidth =945
                    ColumnOrder =9
                    TabIndex =2
                    Name ="hes"
                    ControlSource ="hes"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3917
                            Top =1303
                            Width =840
                            Height =240
                            Name ="Label2"
                            Caption ="كد مشتري"
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
                    Left =1700
                    Top =1984
                    ColumnWidth =720
                    ColumnOrder =6
                    TabIndex =3
                    Name ="MEGHs"
                    ControlSource ="MEGHs"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4472
                            Top =1984
                            Width =420
                            Height =240
                            Name ="Label4"
                            Caption ="كارتن"
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
                    Left =1700
                    Top =2324
                    ColumnWidth =885
                    ColumnOrder =7
                    TabIndex =4
                    Name ="MEGHks"
                    ControlSource ="MEGHks"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4397
                            Top =2324
                            Width =405
                            Height =240
                            Name ="Label5"
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
                    Left =1700
                    Top =2664
                    ColumnOrder =13
                    TabIndex =5
                    Name ="TOTALARZs"
                    ControlSource ="TOTALARZs"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4112
                            Top =2664
                            Width =735
                            Height =240
                            Name ="Label6"
                            Caption ="مبلغ ارزي"
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
                    Left =1700
                    Top =3004
                    ColumnWidth =1050
                    ColumnOrder =4
                    TabIndex =6
                    Name ="DATE_N"
                    ControlSource ="DATE_N"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4217
                            Top =3004
                            Width =825
                            Height =240
                            Name ="Label7"
                            Caption ="تاريخ فاكتور"
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
                    Left =1700
                    Top =3344
                    ColumnWidth =975
                    ColumnOrder =5
                    TabIndex =7
                    Name ="DATE_Nh"
                    ControlSource ="DATE_Nh"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4292
                            Top =3344
                            Width =795
                            Height =240
                            Name ="Label8"
                            Caption ="تاريخ حواله"
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
                    Left =1700
                    Top =3685
                    ColumnWidth =3135
                    ColumnOrder =2
                    TabIndex =8
                    Name ="NAME"
                    ControlSource ="NAME"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4545
                            Top =3690
                            Width =870
                            Height =240
                            Name ="Label9"
                            Caption ="نام مشتري"
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
                    Left =1700
                    Top =4365
                    ColumnWidth =1245
                    ColumnOrder =11
                    TabIndex =9
                    Name ="ANBARF"
                    ControlSource ="ANBARF"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4245
                            Top =4365
                            Width =1140
                            Height =240
                            Name ="Label11"
                            Caption ="شماره صادراتي"
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
                    Left =1700
                    Top =4705
                    ColumnWidth =1125
                    ColumnOrder =12
                    TabIndex =10
                    Name ="ANBAR"
                    ControlSource ="ANBAR"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4710
                            Top =4710
                            Width =1020
                            Height =240
                            Name ="Label12"
                            Caption ="شماره داخلي"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =1702
                    Top =1587
                    Height =270
                    ColumnWidth =915
                    ColumnOrder =10
                    TabIndex =11
                    ForeColor =8388608
                    Name ="SADER"
                    ControlSource ="SADER"
                    RowSourceType ="Value List"
                    RowSource ="0;\"داخلي\";1;\"خارجي\""
                    ColumnWidths ="0"
                    DefaultValue ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =3915
                            Top =1650
                            Width =810
                            Height =240
                            ForeColor =8388608
                            Name ="Label118"
                            Caption ="نوع فروش "
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1984
                    Top =5159
                    ColumnOrder =8
                    TabIndex =12
                    Name ="MABL_KS"
                    ControlSource ="MABL_KS"
                    Format ="Standard"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4591
                            Top =5159
                            Width =795
                            Height =240
                            Name ="Label14"
                            Caption ="مبلغ ريالي"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1984
                    Top =5519
                    ColumnOrder =14
                    TabIndex =13
                    Name ="MBAAs"
                    ControlSource ="MBAAs"
                    Format ="Standard"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5386
                            Top =5519
                            Width =900
                            Height =240
                            Name ="Label16"
                            Caption ="ارزش افزوده"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1984
                    Top =5879
                    ColumnOrder =15
                    TabIndex =14
                    Name ="TAKHFIFs"
                    ControlSource ="TAKHFIFs"
                    Format ="Standard"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4591
                            Top =5879
                            Width =525
                            Height =240
                            Name ="Label18"
                            Caption ="تخفيف"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1984
                    Top =6239
                    ColumnWidth =4185
                    ColumnOrder =3
                    TabIndex =15
                    Name ="MMNAME"
                    ControlSource ="MMNAME"
                    Format ="Standard"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4591
                            Top =6239
                            Width =525
                            Height =240
                            Name ="Label20"
                            Caption ="كالا"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Width =1521
                    Height =315
                    TabIndex =16
                    ForeColor =8388608
                    Name ="MAS"
                    ControlSource ="MAS"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT MPCODE, MPNAME FROM TCOD_MAP WHERE (MPP = 100) ORDER BY MPNAME"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =1560
                            Width =585
                            Height =300
                            ForeColor =8388608
                            Name ="Label116"
                            Caption ="مقصد "
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "FROOSH_DAYLY1_KALA.cls"
