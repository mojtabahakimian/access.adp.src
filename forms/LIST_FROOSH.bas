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
    Width =5896
    RowHeight =360
    ItemSuffix =37
    Left =270
    Top =1380
    Right =7965
    Bottom =5400
    DatasheetForeColor =33554432
    DatasheetGridlinesColor =12632256
    MaxRecords =10000000
    RecSrcDt = Begin
        0x72276195f32ee440
    End
    RecordSource ="SELECT     TOP 100 PERCENT dbo.KALAS.DATE_N, dbo.KALAS.code, dbo.KALAS.kala, dbo"
        ".KALAS.MEGH, dbo.KALAS.MEGHk, dbo.KALAS.CUSTNAME, dbo.KALAS.hes, \015\012       "
        "               dbo.KALAS.MABL, dbo.KALAS.MABL_K, dbo.KALAS.KHFR, dbo.KALAS.GHFR,"
        " dbo.KALAS.VAHCODE, dbo.KALAS.GRPCODE, dbo.KALAS.MOLAH, dbo.KALAS.SHARAYET, \015"
        "\012                      dbo.KALAS.FNUMCO, dbo.KALAS.MANDAH, dbo.KALAS.ANBARCOD"
        "E, dbo.KALAS.N_S, dbo.KALAS.USER_NAME, dbo.KALAS.MAS, dbo.KALAS.NUMBER, \015\012"
        "                      dbo.KALAS.NUMBER1, dbo.KALAS.BARGAH, dbo.KALAS.TAGCODE, db"
        "o.OTHER_DTL_SUB.CAM_KHALY, dbo.OTHER_DTL_SUB.CAM_POOR, \015\012                 "
        "     dbo.OTHER_DTL_SUB.MEGHk AS MEGHkg, dbo.OTHER_DTL.CAMIUN_NUM, dbo.OTHER_DTL."
        "CAMIUN, dbo.OTHER_DTL_SUB.TOZIH\015\012FROM         dbo.KALAS INNER JOIN\015\012"
        "                      dbo.OTHER_DTL_SUB ON dbo.KALAS.code = dbo.OTHER_DTL_SUB.CO"
        "DE AND dbo.KALAS.NUMBER1 = dbo.OTHER_DTL_SUB.NUMBER AND \015\012                "
        "      dbo.KALAS.TAG = dbo.OTHER_DTL_SUB.TAGG + 11 INNER JOIN\015\012            "
        "          dbo.OTHER_DTL ON dbo.KALAS.TAG = dbo.OTHER_DTL.TAG + 11 AND dbo.KALAS."
        "NUMBER1 = dbo.OTHER_DTL.NUMBER\015\012WHERE     (dbo.KALAS.TAGCODE = 13)\015\012"
        "ORDER BY dbo.KALAS.NUMBER1"
    Caption ="ليست فروش"
    OnOpen ="[Event Procedure]"
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
            Height =13449
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
                    Left =4195
                    Top =1700
                    ColumnWidth =930
                    ColumnOrder =1
                    Name ="DATE_N"
                    ControlSource ="DATE_N"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2494
                            Top =1700
                            Width =885
                            Height =240
                            Name ="Label0"
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
                    Left =4195
                    Top =2040
                    ColumnWidth =600
                    ColumnOrder =4
                    TabIndex =1
                    Name ="CODE"
                    ControlSource ="CODE"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2494
                            Top =2040
                            Width =555
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
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =4195
                    Top =2381
                    ColumnWidth =2775
                    ColumnOrder =5
                    TabIndex =2
                    Name ="KALA"
                    ControlSource ="KALA"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2490
                            Top =2385
                            Width =540
                            Height =240
                            Name ="Label2"
                            Caption ="نام كالا"
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
                    Left =4195
                    Top =2721
                    ColumnWidth =780
                    ColumnOrder =9
                    TabIndex =3
                    Name ="MEGH"
                    ControlSource ="MEGH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2494
                            Top =2721
                            Width =735
                            Height =240
                            Name ="Label3"
                            Caption ="مقداركالا"
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
                    Left =4195
                    Top =3061
                    ColumnWidth =1065
                    ColumnOrder =10
                    TabIndex =4
                    Name ="MEGHk"
                    ControlSource ="MEGHk"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2494
                            Top =3061
                            Width =1020
                            Height =240
                            Name ="Label4"
                            Caption ="مقدار كل كالا"
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
                    Left =4195
                    Top =3401
                    ColumnWidth =4320
                    ColumnOrder =8
                    TabIndex =5
                    Name ="CUSTNAME"
                    ControlSource ="CUSTNAME"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2494
                            Top =3401
                            Width =1245
                            Height =240
                            Name ="Label5"
                            Caption ="نام خريدار"
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
                    Left =4195
                    Top =3741
                    ColumnWidth =1020
                    ColumnOrder =7
                    TabIndex =6
                    Name ="HES"
                    ControlSource ="HES"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2490
                            Top =3735
                            Width =900
                            Height =240
                            Name ="Label6"
                            Caption ="كد خريدار"
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
                    Left =4195
                    Top =4081
                    ColumnWidth =900
                    ColumnOrder =11
                    TabIndex =7
                    Name ="MABL"
                    ControlSource ="MABL"
                    Format ="Standard"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2494
                            Top =4081
                            Width =435
                            Height =240
                            Name ="Label7"
                            Caption ="مبلغ"
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
                    Left =4195
                    Top =4422
                    ColumnOrder =12
                    TabIndex =8
                    Name ="MABL_K"
                    ControlSource ="MABL_K"
                    Format ="Standard"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2494
                            Top =4422
                            Width =675
                            Height =240
                            Name ="Label8"
                            Caption ="مبلغ كل"
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
                    Left =4195
                    Top =7143
                    ColumnOrder =22
                    TabIndex =10
                    Name ="MOLAH"
                    ControlSource ="MOLAH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2494
                            Top =7143
                            Width =750
                            Height =240
                            Name ="Label16"
                            Caption ="ملاحظات"
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
                    Left =4195
                    Top =8163
                    ColumnOrder =23
                    TabIndex =11
                    Name ="MANDAH"
                    ControlSource ="MANDAH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2494
                            Top =8163
                            Width =750
                            Height =240
                            Name ="Label19"
                            Caption ="توضيحات"
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
                    Left =4195
                    Top =8844
                    ColumnWidth =990
                    ColumnOrder =24
                    TabIndex =12
                    Name ="N_S"
                    ControlSource ="N_S"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2494
                            Top =8844
                            Width =945
                            Height =240
                            Name ="Label21"
                            Caption ="شماره سند"
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
                    Left =4195
                    Top =9184
                    ColumnOrder =25
                    TabIndex =13
                    Name ="USER_NAME"
                    ControlSource ="USER_NAME"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2494
                            Top =9184
                            Width =1035
                            Height =240
                            Name ="Label22"
                            Caption ="نام كاربر"
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
                    Left =4195
                    Top =9524
                    ColumnWidth =735
                    ColumnOrder =26
                    TabIndex =14
                    Name ="MAS"
                    ControlSource ="MAS"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2494
                            Top =9524
                            Width =450
                            Height =240
                            Name ="Label23"
                            Caption ="مدت"
                        End
                    End
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =4195
                    Top =6462
                    ColumnOrder =21
                    TabIndex =9
                    Name ="VAHCODE"
                    ControlSource ="VAHCODE"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT CODE, NAMES FROM TCOD_VAHEDS"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2494
                            Top =6462
                            Width =660
                            Height =240
                            Name ="Label14"
                            Caption ="واحد"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2154
                    Top =566
                    ColumnWidth =1050
                    ColumnOrder =2
                    TabIndex =15
                    Name ="NUMBER"
                    ControlSource ="NUMBER"
                    OnDblClick ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4530
                            Top =570
                            Width =1050
                            Height =240
                            Name ="Label24"
                            Caption ="شماره فاكتور*"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2154
                    Top =907
                    ColumnWidth =1110
                    ColumnOrder =3
                    TabIndex =16
                    Name ="NUMBER1"
                    ControlSource ="NUMBER1"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4161
                            Top =907
                            Width =1095
                            Height =240
                            Name ="Label25"
                            Caption ="شماره حواله "
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
                    Left =1360
                    Top =10714
                    ColumnWidth =825
                    ColumnOrder =13
                    TabIndex =17
                    Name ="CAM_KHALY"
                    ControlSource ="CAM_KHALY"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3742
                            Top =10714
                            Width =1020
                            Height =240
                            Name ="Label28"
                            Caption ="وزن خالي"
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
                    Left =1360
                    Top =11055
                    ColumnWidth =630
                    ColumnOrder =14
                    TabIndex =18
                    Name ="CAM_POOR"
                    ControlSource ="CAM_POOR"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3772
                            Top =11055
                            Width =990
                            Height =240
                            Name ="Label29"
                            Caption ="وزن پر"
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
                    Left =1360
                    Top =11395
                    ColumnWidth =855
                    ColumnOrder =15
                    TabIndex =19
                    Name ="MEGHkg"
                    ControlSource ="MEGHkg"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4035
                            Top =11400
                            Width =750
                            Height =240
                            Name ="Label30"
                            Caption ="وزن خالص"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =1
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1360
                    Top =12075
                    ColumnOrder =18
                    TabIndex =20
                    Name ="CAMIUN_NUM"
                    ControlSource ="CAMIUN_NUM"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3607
                            Top =12075
                            Width =1155
                            Height =240
                            Name ="Label32"
                            Caption ="شماره كاميون"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =1
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1360
                    Top =12415
                    ColumnOrder =19
                    TabIndex =21
                    Name ="CAMIUN"
                    ControlSource ="CAMIUN"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4020
                            Top =12420
                            Width =795
                            Height =240
                            Name ="Label33"
                            Caption ="نوع كاميون"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =1
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1360
                    Top =12755
                    ColumnWidth =1800
                    ColumnOrder =20
                    TabIndex =22
                    Name ="TOZIH"
                    ControlSource ="TOZIH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4162
                            Top =12755
                            Width =600
                            Height =240
                            Name ="Label34"
                            Caption ="توضيح"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "LIST_FROOSH.cls"
