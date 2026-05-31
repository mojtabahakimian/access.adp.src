Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    FilterOn = NotDefault
    OrderByOn = NotDefault
    DefaultView =0
    AllowUpdating =4
    TabularCharSet =178
    TabularFamily =42
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =11905
    DatasheetFontHeight =10
    ItemSuffix =318
    Left =1005
    Top =780
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x87768acc4ce3e340
    End
    RecordSource ="SELECT HEAD_LST.NUMBER, HEAD_LST.TAG, HEAD_LST.DATE_N, HEAD_LST.CUST_NO, HEAD_LS"
        "T.ANBAR, HEAD_LST.NUMBER1, HEAD_LST.TAH, HEAD_LST.MAS, HEAD_LST.VAS, HEAD_LST.N_"
        "S, HEAD_LST.M_NAGHD, HEAD_LST.MABL_VAR, HEAD_LST.MOIN_VAR, HEAD_LST.MABL_HAV, HE"
        "AD_LST.MOIN_HAV, HEAD_LST.MABL_HAZ, HEAD_LST.MOIN_HAZ, HEAD_LST.TAKHFIF, HEAD_LS"
        "T.MOIN_KHF, HEAD_LST.ANBARF, HEAD_LST.FNUMCO, INVO_LST.NUMBER AS INUMBER, INVO_L"
        "ST.TAG AS ITAG, INVO_LST.ANBAR AS IANBAR, INVO_LST.RADIF, INVO_LST.CODE, INVO_LS"
        "T.MEGH, INVO_LST.MEGHk, INVO_LST.MEGH_MAR, INVO_LST.MANDAH, INVO_LST.MABL, INVO_"
        "LST.MABL_K, INVO_LST.FROM_A, INVO_LST.N_RASID, INVO_LST.MEGH_R, INVO_LST.CUST_NO"
        " AS ICUST_NO, TCOD_ANBAR.NAMES, ISNULL(STUF_DEF.NAME, N' ') + N' ' + ISNULL(INVO"
        "_LST.MANDAH, N' ') AS KALA, TCOD_VAHEDS.NAMES AS VNAMES, HEAD_LST.MOLAH, HEAD_LS"
        "T.USER_NAME, CUST_HESAB.hes, CUST_HESAB.NAME, CUST_HESAB.ADDRESS, STUF_DEF.N_FAN"
        "I, HEAD_LST.SHARAYET, INVO_LST.VAHED_K, CUST_HESAB.TEL, CUST_HESAB.ECODE, CUST_H"
        "ESAB.PCODE, CUST_HESAB.CODE_E FROM STUF_DEF INNER JOIN TCOD_VAHEDS INNER JOIN HE"
        "AD_LST INNER JOIN INVO_LST ON HEAD_LST.NUMBER = INVO_LST.NUMBER AND HEAD_LST.TAG"
        " = INVO_LST.TAG ON TCOD_VAHEDS.CODE = INVO_LST.VAHED_K ON STUF_DEF.CODE = INVO_L"
        "ST.CODE INNER JOIN TCOD_ANBAR ON INVO_LST.ANBAR = TCOD_ANBAR.CODE LEFT OUTER JOI"
        "N CUST_HESAB ON HEAD_LST.CUST_NO = CUST_HESAB.hes WHERE (HEAD_LST.TAG = 2)"
    Caption ="فاكتور فروش"
    DatasheetFontName ="Arial (Arabic)"
    RecordSourceQualifier ="dbo"
    Begin
        Begin Label
            BackStyle =0
            TextFontCharSet =178
            TextAlign =1
            TextFontFamily =42
            FontSize =9
            FontWeight =700
            ForeColor =128
            FontName ="Arial (Arabic)"
        End
        Begin Rectangle
            BackStyle =0
            BorderWidth =1
            Width =850
            Height =850
        End
        Begin Line
            Width =1701
            BorderColor =128
        End
        Begin Image
            OldBorderStyle =0
            PictureAlignment =2
            Width =1701
            Height =1701
        End
        Begin CommandButton
            TextFontCharSet =178
            Width =1701
            Height =283
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Tahoma"
        End
        Begin OptionButton
            LabelX =230
            LabelY =-30
        End
        Begin CheckBox
            LabelX =230
            LabelY =-30
        End
        Begin OptionGroup
            BackStyle =1
            Width =1701
            Height =1701
        End
        Begin BoundObjectFrame
            Width =4536
            Height =2835
            LabelX =-1701
        End
        Begin TextBox
            OldBorderStyle =0
            TextFontCharSet =178
            TextFontFamily =42
            BackStyle =0
            Width =1701
            LabelX =-1701
            FontSize =9
            FontName ="Arial (Arabic)"
        End
        Begin ListBox
            TextFontCharSet =178
            TextFontFamily =42
            OldBorderStyle =0
            Width =1701
            Height =1417
            LabelX =-1701
            FontSize =9
            FontName ="Arial (Arabic)"
        End
        Begin ComboBox
            OldBorderStyle =0
            TextFontCharSet =178
            TextFontFamily =42
            BackStyle =0
            Width =1701
            LabelX =-1701
            FontSize =9
            FontName ="Arial (Arabic)"
        End
        Begin Subform
            OldBorderStyle =0
            Width =1701
            Height =1701
        End
        Begin UnboundObjectFrame
            Width =4536
            Height =2835
        End
        Begin PageBreak
            Width =283
        End
        Begin BreakLevel
            GroupFooter = NotDefault
            KeepTogether =1
            ControlSource ="ANBAR"
        End
        Begin PageHeader
            Height =4025
            Name ="PageHeader1"
            Begin
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =9622
                    Top =915
                    Width =1521
                    Height =375
                    FontSize =12
                    FontWeight =700
                    BorderColor =32768
                    Name ="DATE_N"
                    ControlSource ="DATE_N"
                    FontName ="Titr"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            Visible = NotDefault
                            TextAlign =3
                            TextFontFamily =2
                            Left =11195
                            Top =915
                            Width =570
                            Height =390
                            FontSize =12
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label92"
                            Caption ="تاريخ :"
                            FontName ="Titr"
                        End
                    End
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2385
                    Top =1938
                    Width =8031
                    Height =435
                    FontSize =12
                    TabIndex =1
                    BorderColor =32768
                    Name ="HESAB"
                    ControlSource ="NAME"
                    FontName ="Titr"
                    Begin
                        Begin Label
                            Visible = NotDefault
                            TextAlign =3
                            TextFontFamily =2
                            Left =6630
                            Top =60
                            Width =1860
                            Height =390
                            FontSize =12
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label95"
                            Caption ="نام شخص حقيقي/حقوقي:"
                            FontName ="Titr"
                        End
                    End
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2145
                    Top =2550
                    Width =7881
                    Height =375
                    FontSize =10
                    TabIndex =2
                    BorderColor =32768
                    Name ="ADDRESS"
                    ControlSource ="SHARAYET"
                    FontName ="Titr"
                    Begin
                        Begin Label
                            Visible = NotDefault
                            TextAlign =3
                            TextFontFamily =2
                            Left =10062
                            Top =2550
                            Width =630
                            Height =390
                            FontSize =12
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label97"
                            Caption ="نشاني:"
                            FontName ="Titr"
                        End
                    End
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Top =2552
                    Width =2331
                    Height =465
                    FontSize =12
                    FontWeight =700
                    TabIndex =3
                    BorderColor =32768
                    Name ="TEL"
                    ControlSource ="TEL"
                    FontName ="Titr"
                    Begin
                        Begin Label
                            Visible = NotDefault
                            TextAlign =3
                            TextFontFamily =2
                            Left =2594
                            Top =2550
                            Width =1260
                            Height =390
                            FontSize =12
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label98"
                            Caption ="شماره تلفن/نمابر:"
                            FontName ="Titr"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =3123
                    Top =113
                    FontSize =12
                    TabIndex =4
                    Name ="CODE"
                    ControlSource ="CODE"
                    FontName ="Titr"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Top =1920
                    Height =465
                    FontSize =12
                    FontWeight =700
                    TabIndex =5
                    BorderColor =32768
                    Name ="ECODE"
                    ControlSource ="ECODE"
                    FontName ="Titr"
                    InputMask ="#\"  \"#\"  \"#\"  \"#\"  \"#\"  \"#\"  \"#\"  \"#\"  \"#\"  \"#\"  \"#\"  \"#"
                End
                Begin TextBox
                    Visible = NotDefault
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1352
                    Top =1077
                    Width =3231
                    Height =495
                    FontSize =12
                    TabIndex =6
                    BorderColor =32768
                    Name ="PCODE"
                    ControlSource ="PCODE"
                    FontName ="Titr"
                    InputMask ="#\"  \"#\"  \"#\"  \"#\"  \"#\"  \"#\"  \"#\"  \"#\"  \"#\"  \"#\"  \"#\"  \"#"
                End
                Begin Label
                    TextFontCharSet =2
                    TextAlign =2
                    TextFontFamily =18
                    ReadingOrder =1
                    NumeralShapes =2
                    Left =3285
                    Top =690
                    Width =3750
                    Height =450
                    FontSize =20
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label316"
                    Caption ="*********"
                    FontName ="Webdings"
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =345
            OnFormat ="[Event Procedure]"
            Name ="Detail"
            Begin
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =9930
                    Width =1343
                    Height =340
                    FontSize =10
                    FontWeight =700
                    BorderColor =4210752
                    Name ="CODE11"
                    ControlSource ="N_FANI"
                    Format ="Fixed"
                    FontName ="Titr"
                End
                Begin TextBox
                    DecimalPlaces =0
                    RunningSum =1
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =11395
                    Width =510
                    Height =340
                    FontSize =12
                    FontWeight =700
                    TabIndex =1
                    BorderColor =4210752
                    Name ="RADIFwe"
                    ControlSource ="=1"
                    FontName ="Titr"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3435
                    Width =6431
                    Height =340
                    FontSize =12
                    FontWeight =700
                    TabIndex =2
                    BorderColor =4210752
                    Name ="CODE1"
                    ControlSource ="=IIf([Forms]![baseknow]![codeview]=-1,[kala] & \" \" & [code],[kala])"
                    FontName ="Titr"
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =1
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Width =1191
                    Height =340
                    FontSize =12
                    FontWeight =700
                    TabIndex =3
                    BorderColor =4210752
                    Name ="MEGHk"
                    ControlSource ="MEGHk"
                    FontName ="Titr"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =1
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1165
                    Width =1247
                    Height =340
                    FontSize =12
                    FontWeight =700
                    TabIndex =4
                    BorderColor =4210752
                    Name ="MEGH"
                    ControlSource ="MEGH"
                    FontName ="Titr"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =2
                    RunningSum =1
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =7345
                    Width =1021
                    Height =340
                    FontSize =12
                    FontWeight =700
                    TabIndex =5
                    BorderColor =4210752
                    Name ="MEGHks"
                    ControlSource ="MEGHk"
                    FontName ="Titr"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    RunningSum =1
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =6070
                    Width =1247
                    Height =340
                    FontSize =12
                    FontWeight =700
                    TabIndex =6
                    BorderColor =4210752
                    Name ="MEGHs"
                    ControlSource ="MEGH"
                    FontName ="Titr"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =2385
                    Width =960
                    Height =345
                    FontSize =11
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="vahed"
                    Caption ="كارتن"
                    FontName ="Titr"
                End
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Left =6916
                    Top =56
                    TabIndex =7
                    Name ="VAHED_K"
                    ControlSource ="VAHED_K"
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =0
            Name ="GroupFooter0"
        End
        Begin PageFooter
            Height =5952
            Name ="PageFooter1"
            Begin
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =1
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1190
                    Width =1247
                    Height =340
                    FontSize =12
                    FontWeight =700
                    BorderColor =4210752
                    Name ="Text310"
                    ControlSource ="=[MEGHs]"
                    FontName ="Titr"
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =1
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Width =1192
                    Height =340
                    FontSize =12
                    FontWeight =700
                    TabIndex =1
                    BorderColor =4210752
                    Name ="Text311"
                    ControlSource ="=[MEGHks]"
                    FontName ="Titr"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =2
                    Left =5465
                    Top =907
                    Width =6240
                    Height =405
                    FontSize =10
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label206"
                    Caption ="طبق صورت به تعداد صحيح و سالم تحويل گرفتم.شماره ماشين: "
                    FontName ="Traffic"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =2
                    Left =1190
                    Top =1327
                    Width =10515
                    Height =345
                    ForeColor =0
                    Name ="Label207"
                    Caption ="درصورت هرگونه مغايرت در محموله ارسالي در زمان تخليه و باحضور راننده اعلام شود در"
                        " غير اين صورت شركت هيچ تعهدي ندارد"
                    FontName ="Traffic"
                End
            End
        End
    End
End
CodeBehindForm
' See "HAVALAH_ANBAR_84744.cls"
