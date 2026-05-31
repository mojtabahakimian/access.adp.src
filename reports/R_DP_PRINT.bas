Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    OrderByOn = NotDefault
    DefaultView =0
    AllowUpdating =4
    TabularCharSet =178
    TabularFamily =42
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =10275
    DatasheetFontHeight =10
    ItemSuffix =102
    Left =600
    Top =330
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xa8d9514133ede540
    End
    RecordSource ="SELECT     dbo.PGET_HED.DATE AS DT, dbo.PGET_HED.MOLAH, dbo.PGET_HED.N_S, dbo.PG"
        "ET_LST.DATE, dbo.TCOD_DPS.NAMES, dbo.PGET_LST.NO_AM,                        dbo."
        "PGET_LST.NAHVA, dbo.PGET_LST.FHES_K, dbo.PGET_LST.FHES_M, dbo.PGET_LST.FHES_T, d"
        "bo.PGET_LST.THES_K, dbo.PGET_LST.THES_M,                        dbo.PGET_LST.THE"
        "S_T, dbo.PGET_LST.SHARH, dbo.PGET_LST.MABL, dbo.PGET_LST.N_SERI, dbo.PGET_LST.BA"
        "NK, '' AS kk, dbo.PGET_LST.ID,                        dbo.TCOD_DPSKIND.NAMES AS "
        "nonames, dbo.TOTA_HES.NAME AS FKNAME, dbo.DETA_HES.NAME AS FMNAME, dbo.TDETA_HES"
        ".NAME AS FTNAME,                        TOTA_HES_1.NAME AS TKNAME, DETA_HES_1.NA"
        "ME AS TMNAME, TDETA_HES_1.NAME AS TTNAME, dbo.DEED_HED.base,                    "
        "    dbo.SALA_DTL.EMZA AS emza1, SALA_DTL_1.EMZA AS emza2, SALA_DTL_2.EMZA AS emz"
        "a3, dbo.PGET_HED.sgn1usid, dbo.PGET_HED.sgn2usid,                        dbo.PGE"
        "T_HED.sgn3usid, dbo.PGET_HED.SGN1, dbo.PGET_HED.SGN2, dbo.PGET_HED.SGN3 FROM    "
        "     dbo.SALA_DTL SALA_DTL_1 RIGHT OUTER JOIN                       dbo.SALA_DTL"
        " RIGHT OUTER JOIN                       dbo.TDETA_HES TDETA_HES_1 INNER JOIN    "
        "                   dbo.DETA_HES DETA_HES_1 ON TDETA_HES_1.N_KOL = DETA_HES_1.N_K"
        "OL AND TDETA_HES_1.NUMBER = DETA_HES_1.NUMBER INNER JOIN                       d"
        "bo.PGET_LST INNER JOIN                       dbo.TCOD_DPS ON dbo.PGET_LST.NO_AM "
        "= dbo.TCOD_DPS.CODE INNER JOIN                       dbo.TCOD_DPSKIND ON dbo.PGE"
        "T_LST.NAHVA = dbo.TCOD_DPSKIND.CODE INNER JOIN                       dbo.PGET_HE"
        "D ON dbo.PGET_LST.ID = dbo.PGET_HED.ID INNER JOIN                       dbo.TDET"
        "A_HES ON dbo.PGET_LST.FHES_K = dbo.TDETA_HES.N_KOL AND dbo.PGET_LST.FHES_M = dbo"
        ".TDETA_HES.NUMBER AND                        dbo.PGET_LST.FHES_T = dbo.TDETA_HES"
        ".TNUMBER INNER JOIN                       dbo.DETA_HES ON dbo.TDETA_HES.N_KOL = "
        "dbo.DETA_HES.N_KOL AND dbo.TDETA_HES.NUMBER = dbo.DETA_HES.NUMBER INNER JOIN    "
        "                   dbo.TOTA_HES ON dbo.DETA_HES.N_KOL = dbo.TOTA_HES.NUMBER ON T"
        "DETA_HES_1.N_KOL = dbo.PGET_LST.THES_K AND                        TDETA_HES_1.NU"
        "MBER = dbo.PGET_LST.THES_M AND TDETA_HES_1.TNUMBER = dbo.PGET_LST.THES_T INNER J"
        "OIN                       dbo.TOTA_HES TOTA_HES_1 ON DETA_HES_1.N_KOL = TOTA_HES"
        "_1.NUMBER INNER JOIN                       dbo.DEED_HED ON dbo.PGET_HED.N_S = db"
        "o.DEED_HED.N_S LEFT OUTER JOIN                       dbo.SALA_DTL SALA_DTL_2 ON "
        "dbo.PGET_HED.sgn3usid = SALA_DTL_2.IDD ON dbo.SALA_DTL.IDD = dbo.PGET_HED.sgn1us"
        "id ON                        SALA_DTL_1.IDD = dbo.PGET_HED.sgn2usid"
    Caption ="دريافت پرداخت روزانه"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial (Arabic)"
    OnError ="[Event Procedure]"
    RecordSource ="SELECT     dbo.PGET_HED.DATE AS DT, dbo.PGET_HED.MOLAH, dbo.PGET_HED.N_S, dbo.PG"
        "ET_LST.DATE, dbo.TCOD_DPS.NAMES, dbo.PGET_LST.NO_AM,                        dbo."
        "PGET_LST.NAHVA, dbo.PGET_LST.FHES_K, dbo.PGET_LST.FHES_M, dbo.PGET_LST.FHES_T, d"
        "bo.PGET_LST.THES_K, dbo.PGET_LST.THES_M,                        dbo.PGET_LST.THE"
        "S_T, dbo.PGET_LST.SHARH, dbo.PGET_LST.MABL, dbo.PGET_LST.N_SERI, dbo.PGET_LST.BA"
        "NK, '' AS kk, dbo.PGET_LST.ID,                        dbo.TCOD_DPSKIND.NAMES AS "
        "nonames, dbo.TOTA_HES.NAME AS FKNAME, dbo.DETA_HES.NAME AS FMNAME, dbo.TDETA_HES"
        ".NAME AS FTNAME,                        TOTA_HES_1.NAME AS TKNAME, DETA_HES_1.NA"
        "ME AS TMNAME, TDETA_HES_1.NAME AS TTNAME, dbo.DEED_HED.base,                    "
        "    dbo.SALA_DTL.EMZA AS emza1, SALA_DTL_1.EMZA AS emza2, SALA_DTL_2.EMZA AS emz"
        "a3, dbo.PGET_HED.sgn1usid, dbo.PGET_HED.sgn2usid,                        dbo.PGE"
        "T_HED.sgn3usid, dbo.PGET_HED.SGN1, dbo.PGET_HED.SGN2, dbo.PGET_HED.SGN3 FROM    "
        "     dbo.SALA_DTL SALA_DTL_1 RIGHT OUTER JOIN                       dbo.SALA_DTL"
        " RIGHT OUTER JOIN                       dbo.TDETA_HES TDETA_HES_1 INNER JOIN    "
        "                   dbo.DETA_HES DETA_HES_1 ON TDETA_HES_1.N_KOL = DETA_HES_1.N_K"
        "OL AND TDETA_HES_1.NUMBER = DETA_HES_1.NUMBER INNER JOIN                       d"
        "bo.PGET_LST INNER JOIN                       dbo.TCOD_DPS ON dbo.PGET_LST.NO_AM "
        "= dbo.TCOD_DPS.CODE INNER JOIN                       dbo.TCOD_DPSKIND ON dbo.PGE"
        "T_LST.NAHVA = dbo.TCOD_DPSKIND.CODE INNER JOIN                       dbo.PGET_HE"
        "D ON dbo.PGET_LST.ID = dbo.PGET_HED.ID INNER JOIN                       dbo.TDET"
        "A_HES ON dbo.PGET_LST.FHES_K = dbo.TDETA_HES.N_KOL AND dbo.PGET_LST.FHES_M = dbo"
        ".TDETA_HES.NUMBER AND                        dbo.PGET_LST.FHES_T = dbo.TDETA_HES"
        ".TNUMBER INNER JOIN                       dbo.DETA_HES ON dbo.TDETA_HES.N_KOL = "
        "dbo.DETA_HES.N_KOL AND dbo.TDETA_HES.NUMBER = dbo.DETA_HES.NUMBER INNER JOIN    "
        "                   dbo.TOTA_HES ON dbo.DETA_HES.N_KOL = dbo.TOTA_HES.NUMBER ON T"
        "DETA_HES_1.N_KOL = dbo.PGET_LST.THES_K AND                        TDETA_HES_1.NU"
        "MBER = dbo.PGET_LST.THES_M AND TDETA_HES_1.TNUMBER = dbo.PGET_LST.THES_T INNER J"
        "OIN                       dbo.TOTA_HES TOTA_HES_1 ON DETA_HES_1.N_KOL = TOTA_HES"
        "_1.NUMBER INNER JOIN                       dbo.DEED_HED ON dbo.PGET_HED.N_S = db"
        "o.DEED_HED.N_S LEFT OUTER JOIN                       dbo.SALA_DTL SALA_DTL_2 ON "
        "dbo.PGET_HED.sgn3usid = SALA_DTL_2.IDD ON dbo.SALA_DTL.IDD = dbo.PGET_HED.sgn1us"
        "id ON                        SALA_DTL_1.IDD = dbo.PGET_HED.sgn2usid"
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
        Begin BreakLevel
            GroupFooter = NotDefault
            ControlSource ="DT"
        End
        Begin PageHeader
            Height =1755
            Name ="PageHeader"
            Begin
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =9525
                    Top =917
                    Width =555
                    Height =765
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="N_S Label"
                    Caption ="شماره رديف"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="N_S_Label"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =8409
                    Top =920
                    Width =960
                    Height =405
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="DATE_S Label"
                    Caption ="نوع عمليات"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="DATE_S_Label"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =6270
                    Top =900
                    Width =810
                    Height =405
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="SHARH Label"
                    Caption ="از حساب "
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="SHARH_Label"
                End
                Begin Line
                    BorderWidth =2
                    Left =15
                    Top =902
                    Width =10165
                    Name ="Line26"
                End
                Begin Line
                    BorderWidth =2
                    Left =15
                    Top =1727
                    Width =10165
                    Name ="Line27"
                End
                Begin Line
                    Left =8327
                    Top =923
                    Width =0
                    Height =794
                    Name ="Line29"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =4725
                    Top =917
                    Width =0
                    Height =794
                    Name ="Line32"
                End
                Begin Line
                    BorderWidth =2
                    Top =920
                    Width =0
                    Height =794
                    Name ="Line36"
                End
                Begin Line
                    BorderWidth =2
                    Left =10155
                    Top =917
                    Width =0
                    Height =794
                    Name ="Line37"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =4065
                    Width =2235
                    Height =390
                    FontSize =10
                    FontWeight =900
                    ForeColor =0
                    Name ="Label22"
                    Caption ="چاپ صورتحساب عملكرد خزانه"
                    FontName ="Titr"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =1486
                    Top =917
                    Width =3225
                    Height =765
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="Label72"
                    Caption ="شرح"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =45
                    Top =910
                    Width =1356
                    Height =765
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="Label80"
                    Caption ="مبلغ"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    Left =1456
                    Top =917
                    Width =0
                    Height =794
                    Name ="Line81"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =8402
                    Top =1322
                    Width =975
                    Height =405
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="Label86"
                    Caption ="نحوه "
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    Left =4740
                    Top =1307
                    Width =4734
                    Name ="Line87"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =6255
                    Top =1350
                    Width =840
                    Height =405
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="Label90"
                    Caption ="به حساب "
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =9694
                    Top =397
                    Width =480
                    Height =405
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="Label92"
                    Caption ="تاريخ"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    Left =9465
                    Top =920
                    Width =0
                    Height =794
                    Name ="Line94"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =8320
                    Top =397
                    Width =1311
                    Height =375
                    FontSize =8
                    Name ="Text93"
                    ControlSource ="DT"
                    FontName ="WeblogmaYekan"
                    InputMask ="####/##/##"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =1343
                    Top =450
                    Width =615
                    Height =405
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="Label97"
                    Caption ="شماره:"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =56
                    Top =453
                    Width =1281
                    Height =360
                    FontSize =8
                    TabIndex =1
                    Name ="ID"
                    ControlSource ="ID"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2430
                    Top =390
                    Width =5433
                    Height =375
                    FontSize =10
                    TabIndex =2
                    BorderColor =32768
                    Name ="Text90"
                    ControlSource ="=[Forms]![Baseknow]![WIDTH_D]"
                    FontName ="WeblogmaYekan"
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            CanGrow = NotDefault
            Height =779
            Name ="Detail"
            Begin
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4770
                    Width =3525
                    Height =375
                    ColumnWidth =4650
                    FontSize =8
                    Name ="SHARH"
                    ControlSource ="=[FKName] & \"-\" & [FMName] & \"-\" & [FTName]"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    RunningSum =1
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =9495
                    Width =621
                    Height =375
                    FontSize =8
                    TabIndex =1
                    Name ="N_S"
                    ControlSource ="=1"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    Left =8327
                    Width =0
                    Height =779
                    Name ="Line39"
                End
                Begin Line
                    Left =4725
                    Width =0
                    Height =779
                    Name ="Line42"
                End
                Begin Line
                    BorderWidth =2
                    Width =0
                    Height =779
                    Name ="Line46"
                End
                Begin Line
                    BorderWidth =2
                    Left =10155
                    Width =0
                    Height =779
                    Name ="Line47"
                End
                Begin Line
                    LineSlant = NotDefault
                    Top =779
                    Width =10178
                    BorderColor =4210752
                    Name ="Line48"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =8342
                    Width =1071
                    Height =375
                    FontSize =8
                    TabIndex =2
                    Name ="Text71"
                    ControlSource ="NAMES"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1485
                    Width =3231
                    Height =750
                    FontSize =8
                    TabIndex =3
                    Name ="Text75"
                    ControlSource ="SHARH"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =47
                    Width =1356
                    Height =375
                    FontSize =8
                    TabIndex =4
                    Name ="MABL"
                    ControlSource ="MABL"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    Left =1456
                    Width =0
                    Height =779
                    Name ="Line83"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =8325
                    Top =390
                    Width =1101
                    Height =375
                    FontSize =8
                    TabIndex =5
                    Name ="nonames"
                    ControlSource ="nonames"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    Left =4740
                    Top =375
                    Width =4749
                    Name ="Line89"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4770
                    Top =390
                    Width =3525
                    Height =375
                    FontSize =8
                    TabIndex =6
                    Name ="Text91"
                    ControlSource ="=[TKName] & \"-\" & [TMName] & \"-\" & [TTName]"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    Left =9465
                    Width =0
                    Height =779
                    Name ="Line95"
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =2740
            OnFormat ="[Event Procedure]"
            Name ="GroupFooter1"
            Begin
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5328
                    Top =623
                    Width =4421
                    Height =300
                    FontSize =8
                    Name ="Text23"
                    ControlSource ="=FARSIDATE(Now())"
                    FontName ="WeblogmaYekan"
                    InputMask ="####/##/##"
                End
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =113
                    Top =566
                    Width =4421
                    Height =300
                    FontSize =8
                    TabIndex =1
                    Name ="Text24"
                    ControlSource ="=\"صفحه \" & [Page] & \" از \" & [Pages]"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    BorderWidth =2
                    Width =0
                    Height =419
                    Name ="Line64"
                End
                Begin Line
                    LineSlant = NotDefault
                    Top =419
                    Width =10178
                    BorderColor =4210752
                    Name ="Line65"
                End
                Begin Line
                    BorderWidth =2
                    Left =10141
                    Width =0
                    Height =419
                    Name ="Line67"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =1644
                    Width =645
                    Height =375
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="Label68"
                    Caption ="جمع كل:"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =61
                    Width =1371
                    Height =375
                    FontSize =8
                    TabIndex =2
                    Name ="Text84"
                    ControlSource ="=Sum([MABL])"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =1457
                    Width =0
                    Height =419
                    Name ="Line85"
                End
                Begin Line
                    BorderWidth =2
                    Width =10165
                    Name ="Line96"
                End
                Begin BoundObjectFrame
                    SizeMode =3
                    OldBorderStyle =0
                    Left =6840
                    Top =934
                    Width =1701
                    Height =1701
                    TabIndex =3
                    Name ="EMZA1"
                    ControlSource ="EMZA1"
                End
                Begin BoundObjectFrame
                    SizeMode =3
                    OldBorderStyle =0
                    Left =3435
                    Top =919
                    Width =1701
                    Height =1701
                    TabIndex =4
                    Name ="EMZA2"
                    ControlSource ="EMZA2"
                End
                Begin BoundObjectFrame
                    SizeMode =3
                    OldBorderStyle =0
                    Top =915
                    Width =1701
                    Height =1701
                    TabIndex =5
                    Name ="EMZA3"
                    ControlSource ="EMZA3"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =8550
                    Top =1354
                    Width =1725
                    Height =285
                    FontSize =8
                    FontWeight =400
                    BackColor =128
                    ForeColor =0
                    Name ="semat1"
                    Caption ="."
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =5128
                    Top =1354
                    Width =1725
                    Height =285
                    FontSize =8
                    FontWeight =400
                    BackColor =128
                    ForeColor =0
                    Name ="semat2"
                    Caption ="."
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =1710
                    Top =1354
                    Width =1725
                    Height =285
                    FontSize =8
                    FontWeight =400
                    BackColor =128
                    ForeColor =0
                    Name ="semat3"
                    Caption ="."
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =8547
                    Top =1021
                    Width =1725
                    Height =285
                    FontSize =8
                    FontWeight =400
                    BackColor =128
                    ForeColor =0
                    Name ="nemz1"
                    Caption ="."
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =5130
                    Top =1021
                    Width =1725
                    Height =285
                    FontSize =8
                    FontWeight =400
                    BackColor =128
                    ForeColor =0
                    Name ="nemz2"
                    Caption ="."
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =1695
                    Top =1021
                    Width =1725
                    Height =285
                    FontSize =8
                    FontWeight =400
                    BackColor =128
                    ForeColor =0
                    Name ="nemz3"
                    Caption ="."
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Left =5952
                    Top =1530
                    TabIndex =6
                    Name ="sgn1usid"
                    ControlSource ="sgn1usid"
                    Begin
                        Begin Label
                            TextAlign =0
                            Left =4251
                            Top =1530
                            Width =930
                            Height =255
                            Name ="Label99"
                            Caption ="sgn1usid:"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Left =5952
                    Top =1870
                    TabIndex =7
                    Name ="sgn2usid"
                    ControlSource ="sgn2usid"
                    Begin
                        Begin Label
                            TextAlign =0
                            Left =4251
                            Top =1870
                            Width =930
                            Height =255
                            Name ="Label100"
                            Caption ="sgn2usid:"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Left =5952
                    Top =2211
                    TabIndex =8
                    Name ="sgn3usid"
                    ControlSource ="sgn3usid"
                    Begin
                        Begin Label
                            TextAlign =0
                            Left =4251
                            Top =2211
                            Width =930
                            Height =255
                            Name ="Label101"
                            Caption ="sgn3usid:"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Left =1984
                    Top =1757
                    TabIndex =9
                    Name ="SGN1"
                    ControlSource ="SGN1"
                End
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Left =1984
                    Top =2076
                    TabIndex =10
                    Name ="SGN2"
                    ControlSource ="SGN2"
                End
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Left =1984
                    Top =2395
                    TabIndex =11
                    Name ="SGN3"
                    ControlSource ="SGN3"
                End
            End
        End
        Begin PageFooter
            Height =0
            Name ="PageFooter"
        End
    End
End
CodeBehindForm
' See "R_DP_PRINT.cls"
