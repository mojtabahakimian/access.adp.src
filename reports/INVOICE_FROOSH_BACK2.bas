Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    DefaultView =0
    AllowUpdating =4
    TabularCharSet =178
    TabularFamily =42
    DateGrouping =1
    Orientation =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =10619
    DatasheetFontHeight =10
    ItemSuffix =168
    Left =930
    Top =1830
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x91bf526914f9e540
    End
    RecordSource ="SELECT        dbo.HEAD_BACK.NUMBER, dbo.HEAD_BACK.HTAG, dbo.HEAD_BACK.DATE_N, db"
        "o.HEAD_BACK.ANBAR, dbo.HEAD_BACK.CUST_NO, \015\012                         dbo.H"
        "EAD_BACK.NUMBER1, dbo.HEAD_BACK.TAH, dbo.HEAD_BACK.MAS, dbo.HEAD_BACK.VAS, dbo.H"
        "EAD_BACK.N_S, dbo.HEAD_BACK.MOLAH, \015\012                         dbo.HEAD_BAC"
        "K.M_NAGHD, dbo.HEAD_BACK.MABL_VAR, dbo.HEAD_BACK.MOIN_VAR, dbo.HEAD_BACK.MABL_HA"
        "V, dbo.HEAD_BACK.MOIN_HAV, \015\012                         dbo.HEAD_BACK.MABL_H"
        "AZ, dbo.HEAD_BACK.MOIN_HAZ, dbo.HEAD_BACK.TAKHFIF, dbo.HEAD_BACK.MOIN_KHF, dbo.H"
        "EAD_BACK.ANBARF, \015\012                         dbo.HEAD_BACK.FNUMCO, dbo.INVO"
        "_LST.NUMBER AS INUMBER, dbo.INVO_LST.TAG, dbo.INVO_LST.ANBAR AS IANBAR, dbo.INVO"
        "_LST.RADIF, \015\012                         dbo.INVO_LST.CODE, dbo.INVO_LST.MEG"
        "H, dbo.INVO_LST.MEGHk, dbo.INVO_LST.MEGH_MAR, dbo.INVO_LST.MANDAH, dbo.INVO_LST."
        "MABL, \015\012                         dbo.INVO_LST.MABL_K, dbo.INVO_LST.FROM_A,"
        " dbo.INVO_LST.N_RASID, dbo.INVO_LST.MEGH_R, dbo.INVO_LST.RADAH, dbo.INVO_LST.SAN"
        "AD_NO, \015\012                         dbo.INVO_LST.CUST_NO AS ICUST_NO, dbo.IN"
        "VO_LST.ANBARF AS IANBARF, dbo.INVO_LST.VAHED_K, dbo.CUST_HESAB.NAME, \015\012   "
        "                      dbo.CUST_HESAB.ADDRESS, dbo.CUST_HESAB.TEL, dbo.TCOD_ANBAR"
        ".NAMES, dbo.STUF_DEF.NAME AS SNAME, dbo.CUST_HESAB.hes, \015\012                "
        "         dbo.TCOD_VAHEDS.NAMES AS VNAMES, dbo.STUF_DEF.N_FANI, dbo.CUST_HESAB.CO"
        "DE_E, dbo.CUST_HESAB.ECODE, dbo.CUST_HESAB.PCODE, \015\012                      "
        "   dbo.CUST_HESAB.IYALAT, dbo.CUST_HESAB.CITY, dbo.CUST_HESAB.MCODEM, dbo.CUST_H"
        "ESAB.TOZIH, dbo.CUST_HESAB.CUST_COD, \015\012                         dbo.CUST_H"
        "ESAB.MOBILE, dbo.CUST_HESAB.Longitude, dbo.CUST_HESAB.Latitude, dbo.CUST_HESAB.R"
        "OUTE_NAME, dbo.HEAD_BACK.DEPATMAN, \015\012                         dbo.HEAD_BAC"
        "K.MBAA, dbo.HEAD_BACK.SGN1, dbo.HEAD_BACK.SGN2, dbo.HEAD_BACK.SGN3, dbo.HEAD_BAC"
        "K.SGN1usid, dbo.HEAD_BACK.sgn2usid, \015\012                         dbo.HEAD_BA"
        "CK.sgn3usid, dbo.SALA_DTL.EMZA AS EMZA1, SALA_DTL_1.EMZA AS EMZA2, SALA_DTL_2.EM"
        "ZA AS EMZA3\015\012FROM            dbo.SALA_DTL AS SALA_DTL_2 RIGHT OUTER JOIN\015"
        "\012                         dbo.CUST_HESAB INNER JOIN\015\012                  "
        "       dbo.STUF_DEF INNER JOIN\015\012                         dbo.HEAD_BACK INN"
        "ER JOIN\015\012                         dbo.INVO_LST INNER JOIN\015\012         "
        "                dbo.TCOD_VAHEDS ON dbo.INVO_LST.VAHED_K = dbo.TCOD_VAHEDS.CODE I"
        "NNER JOIN\015\012                         dbo.TCOD_ANBAR ON dbo.INVO_LST.ANBAR ="
        " dbo.TCOD_ANBAR.CODE ON dbo.HEAD_BACK.NUMBER1 = dbo.INVO_LST.NUMBER AND \015\012"
        "                         dbo.HEAD_BACK.HTAG = dbo.INVO_LST.TAG ON dbo.STUF_DEF.C"
        "ODE = dbo.INVO_LST.CODE ON dbo.CUST_HESAB.hes = dbo.HEAD_BACK.CUST_NO ON \015\012"
        "                         SALA_DTL_2.IDD = dbo.HEAD_BACK.sgn3usid LEFT OUTER JOIN"
        "\015\012                         dbo.SALA_DTL AS SALA_DTL_1 ON dbo.HEAD_BACK.sgn"
        "2usid = SALA_DTL_1.IDD LEFT OUTER JOIN\015\012                         dbo.SALA_"
        "DTL ON dbo.HEAD_BACK.SGN1usid = dbo.SALA_DTL.IDD\015\012WHERE        (dbo.INVO_L"
        "ST.MEGH_MAR <> 0)"
    Caption ="فاكتور برگشت فروش"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial (Arabic)"
    OnError ="[Event Procedure]"
    RecordSource ="SELECT        dbo.HEAD_BACK.NUMBER, dbo.HEAD_BACK.HTAG, dbo.HEAD_BACK.DATE_N, db"
        "o.HEAD_BACK.ANBAR, dbo.HEAD_BACK.CUST_NO, \015\012                         dbo.H"
        "EAD_BACK.NUMBER1, dbo.HEAD_BACK.TAH, dbo.HEAD_BACK.MAS, dbo.HEAD_BACK.VAS, dbo.H"
        "EAD_BACK.N_S, dbo.HEAD_BACK.MOLAH, \015\012                         dbo.HEAD_BAC"
        "K.M_NAGHD, dbo.HEAD_BACK.MABL_VAR, dbo.HEAD_BACK.MOIN_VAR, dbo.HEAD_BACK.MABL_HA"
        "V, dbo.HEAD_BACK.MOIN_HAV, \015\012                         dbo.HEAD_BACK.MABL_H"
        "AZ, dbo.HEAD_BACK.MOIN_HAZ, dbo.HEAD_BACK.TAKHFIF, dbo.HEAD_BACK.MOIN_KHF, dbo.H"
        "EAD_BACK.ANBARF, \015\012                         dbo.HEAD_BACK.FNUMCO, dbo.INVO"
        "_LST.NUMBER AS INUMBER, dbo.INVO_LST.TAG, dbo.INVO_LST.ANBAR AS IANBAR, dbo.INVO"
        "_LST.RADIF, \015\012                         dbo.INVO_LST.CODE, dbo.INVO_LST.MEG"
        "H, dbo.INVO_LST.MEGHk, dbo.INVO_LST.MEGH_MAR, dbo.INVO_LST.MANDAH, dbo.INVO_LST."
        "MABL, \015\012                         dbo.INVO_LST.MABL_K, dbo.INVO_LST.FROM_A,"
        " dbo.INVO_LST.N_RASID, dbo.INVO_LST.MEGH_R, dbo.INVO_LST.RADAH, dbo.INVO_LST.SAN"
        "AD_NO, \015\012                         dbo.INVO_LST.CUST_NO AS ICUST_NO, dbo.IN"
        "VO_LST.ANBARF AS IANBARF, dbo.INVO_LST.VAHED_K, dbo.CUST_HESAB.NAME, \015\012   "
        "                      dbo.CUST_HESAB.ADDRESS, dbo.CUST_HESAB.TEL, dbo.TCOD_ANBAR"
        ".NAMES, dbo.STUF_DEF.NAME AS SNAME, dbo.CUST_HESAB.hes, \015\012                "
        "         dbo.TCOD_VAHEDS.NAMES AS VNAMES, dbo.STUF_DEF.N_FANI, dbo.CUST_HESAB.CO"
        "DE_E, dbo.CUST_HESAB.ECODE, dbo.CUST_HESAB.PCODE, \015\012                      "
        "   dbo.CUST_HESAB.IYALAT, dbo.CUST_HESAB.CITY, dbo.CUST_HESAB.MCODEM, dbo.CUST_H"
        "ESAB.TOZIH, dbo.CUST_HESAB.CUST_COD, \015\012                         dbo.CUST_H"
        "ESAB.MOBILE, dbo.CUST_HESAB.Longitude, dbo.CUST_HESAB.Latitude, dbo.CUST_HESAB.R"
        "OUTE_NAME, dbo.HEAD_BACK.DEPATMAN, \015\012                         dbo.HEAD_BAC"
        "K.MBAA, dbo.HEAD_BACK.SGN1, dbo.HEAD_BACK.SGN2, dbo.HEAD_BACK.SGN3, dbo.HEAD_BAC"
        "K.SGN1usid, dbo.HEAD_BACK.sgn2usid, \015\012                         dbo.HEAD_BA"
        "CK.sgn3usid, dbo.SALA_DTL.EMZA AS EMZA1, SALA_DTL_1.EMZA AS EMZA2, SALA_DTL_2.EM"
        "ZA AS EMZA3\015\012FROM            dbo.SALA_DTL AS SALA_DTL_2 RIGHT OUTER JOIN\015"
        "\012                         dbo.CUST_HESAB INNER JOIN\015\012                  "
        "       dbo.STUF_DEF INNER JOIN\015\012                         dbo.HEAD_BACK INN"
        "ER JOIN\015\012                         dbo.INVO_LST INNER JOIN\015\012         "
        "                dbo.TCOD_VAHEDS ON dbo.INVO_LST.VAHED_K = dbo.TCOD_VAHEDS.CODE I"
        "NNER JOIN\015\012                         dbo.TCOD_ANBAR ON dbo.INVO_LST.ANBAR ="
        " dbo.TCOD_ANBAR.CODE ON dbo.HEAD_BACK.NUMBER1 = dbo.INVO_LST.NUMBER AND \015\012"
        "                         dbo.HEAD_BACK.HTAG = dbo.INVO_LST.TAG ON dbo.STUF_DEF.C"
        "ODE = dbo.INVO_LST.CODE ON dbo.CUST_HESAB.hes = dbo.HEAD_BACK.CUST_NO ON \015\012"
        "                         SALA_DTL_2.IDD = dbo.HEAD_BACK.sgn3usid LEFT OUTER JOIN"
        "\015\012                         dbo.SALA_DTL AS SALA_DTL_1 ON dbo.HEAD_BACK.sgn"
        "2usid = SALA_DTL_1.IDD LEFT OUTER JOIN\015\012                         dbo.SALA_"
        "DTL ON dbo.HEAD_BACK.SGN1usid = dbo.SALA_DTL.IDD\015\012WHERE        (dbo.INVO_L"
        "ST.MEGH_MAR <> 0)"
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
            GroupHeader = NotDefault
            GroupFooter = NotDefault
            ControlSource ="[NUMBER1]"
        End
        Begin BreakLevel
            ControlSource ="RADIF"
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =3300
            Name ="GroupHeader2"
            Begin
                Begin Label
                    TextFontFamily =2
                    Left =3945
                    Width =2130
                    Height =540
                    FontSize =14
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label88"
                    Caption ="فاكتور برگشت فروش"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2205
                    Top =510
                    Width =5613
                    Height =540
                    FontSize =14
                    BorderColor =32768
                    Name ="Text90"
                    ControlSource ="=[Forms]![Baseknow]![WIDTH_D]"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =60
                    Top =540
                    Width =1116
                    Height =450
                    FontSize =12
                    TabIndex =1
                    BorderColor =32768
                    Name ="DATE_N"
                    ControlSource ="DATE_N"
                    FontName ="WeblogmaYekan"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =1213
                            Top =540
                            Width =660
                            Height =480
                            FontSize =12
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label92"
                            Caption ="تاريخ :"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =7815
                    Top =510
                    Width =1356
                    Height =435
                    FontSize =12
                    TabIndex =2
                    BorderColor =32768
                    Name ="NUMBER"
                    ControlSource ="NUMBER"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =9222
                            Top =510
                            Width =1275
                            Height =480
                            FontSize =12
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label93"
                            Caption ="شماره فاكتور:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin Line
                    BorderWidth =3
                    Top =1650
                    Width =10611
                    BorderColor =0
                    Name ="Line94"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1875
                    Top =1662
                    Width =7671
                    Height =450
                    FontSize =12
                    TabIndex =3
                    BorderColor =32768
                    Name ="CUST_NO"
                    ControlSource ="NAME"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =9525
                            Top =1665
                            Width =1005
                            Height =480
                            FontSize =12
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label95"
                            Caption ="مشتري:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Top =1650
                    Width =1101
                    Height =450
                    FontSize =12
                    TabIndex =4
                    BorderColor =32768
                    Name ="hes"
                    ControlSource ="hes"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =1092
                            Top =1650
                            Width =405
                            Height =480
                            FontSize =12
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label96"
                            Caption ="كد:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1875
                    Top =2229
                    Width =7821
                    Height =450
                    FontSize =12
                    TabIndex =5
                    BorderColor =32768
                    Name ="ADDRESS"
                    ControlSource ="ADDRESS"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =9732
                            Top =2229
                            Width =765
                            Height =480
                            FontSize =12
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label97"
                            Caption ="آدرس:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =60
                    Top =2220
                    Width =1011
                    Height =450
                    FontSize =12
                    TabIndex =6
                    BorderColor =32768
                    Name ="TEL"
                    ControlSource ="TEL"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =1092
                            Top =2220
                            Width =555
                            Height =480
                            FontSize =12
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label98"
                            Caption ="تلفن:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin Line
                    BorderWidth =3
                    Top =2775
                    Width =10611
                    BorderColor =0
                    Name ="Line99"
                End
                Begin Line
                    BorderWidth =3
                    Left =10575
                    Top =2790
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line130"
                End
                Begin Line
                    Left =10005
                    Top =2790
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line131"
                End
                Begin Line
                    Left =9435
                    Top =2790
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line132"
                End
                Begin Line
                    Left =5625
                    Top =2790
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line133"
                End
                Begin Line
                    Left =4485
                    Top =2790
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line134"
                End
                Begin Line
                    Left =2715
                    Top =2790
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line136"
                End
                Begin Line
                    Left =1215
                    Top =2790
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line137"
                End
                Begin Line
                    BorderWidth =3
                    Left =30
                    Top =2790
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line138"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =345
                    Top =2835
                    Width =585
                    Height =405
                    FontSize =10
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label139"
                    Caption ="مبــلغ"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =1689
                    Top =2831
                    Width =585
                    Height =405
                    FontSize =10
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label140"
                    Caption ="فـــي"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =3645
                    Top =2835
                    Width =810
                    Height =405
                    FontSize =10
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label141"
                    Caption ="مقدار كل"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =4687
                    Top =2835
                    Width =750
                    Height =405
                    FontSize =10
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label143"
                    Caption ="واحد كالا"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =6727
                    Top =2835
                    Width =1710
                    Height =405
                    FontSize =10
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label144"
                    Caption ="شــــــــــرح  كالا"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =9510
                    Top =2835
                    Width =375
                    Height =405
                    FontSize =10
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label145"
                    Caption ="انبار"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =10011
                    Top =2831
                    Width =540
                    Height =405
                    FontSize =10
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label146"
                    Caption ="رديف"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    BorderWidth =3
                    Top =3300
                    Width =10611
                    BorderColor =0
                    Name ="Line147"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =2782
                    Top =2835
                    Width =720
                    Height =405
                    FontSize =10
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label153"
                    Caption ="مرجوعي"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    Left =3600
                    Top =2790
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line155"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =7592
                    Top =1077
                    Width =1641
                    Height =435
                    FontSize =12
                    TabIndex =7
                    BorderColor =32768
                    Name ="NUMBER1"
                    ControlSource ="NUMBER1"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =9330
                            Top =1080
                            Width =1230
                            Height =450
                            FontSize =12
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label161"
                            Caption ="شماره حواله :"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =7313
                    Top =113
                    TabIndex =8
                    Name ="HTAG"
                    ControlSource ="HTAG"
                    FontName ="WeblogmaYekan"
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =450
            Name ="Detail"
            Begin
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =9467
                    Top =16
                    Width =504
                    Height =405
                    FontSize =8
                    BorderColor =32768
                    Name ="IANBAR"
                    ControlSource ="IANBAR"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =0
                    RunningSum =1
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =9977
                    Width =504
                    Height =405
                    FontSize =8
                    TabIndex =1
                    BorderColor =32768
                    Name ="RADIF123"
                    ControlSource ="=1"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5637
                    Top =16
                    Width =3774
                    Height =405
                    FontSize =8
                    TabIndex =2
                    BorderColor =32768
                    Name ="CODE"
                    ControlSource ="SNAME"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3645
                    Top =15
                    Width =804
                    Height =405
                    FontSize =8
                    TabIndex =3
                    BorderColor =32768
                    Name ="MEGHk"
                    ControlSource ="MEGHk"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1234
                    Top =16
                    Width =1449
                    Height =405
                    FontSize =8
                    TabIndex =4
                    BorderColor =32768
                    Name ="MABL"
                    ControlSource ="MABL"
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =60
                    Top =16
                    Width =1119
                    Height =405
                    FontSize =8
                    TabIndex =5
                    BorderColor =32768
                    Name ="MABL_K"
                    ControlSource ="=[MABL]*[MEGH_MAR]"
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4485
                    Top =15
                    Width =1110
                    Height =405
                    FontSize =8
                    TabIndex =6
                    BorderColor =32768
                    Name ="VAHED_K"
                    ControlSource ="VNAMES"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Top =450
                    Width =10596
                    BorderColor =0
                    Name ="Line100"
                End
                Begin Line
                    BorderWidth =3
                    Left =10575
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line101"
                End
                Begin Line
                    BorderWidth =1
                    Left =10005
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line102"
                End
                Begin Line
                    BorderWidth =1
                    Left =9435
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line103"
                End
                Begin Line
                    BorderWidth =1
                    Left =5625
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line104"
                End
                Begin Line
                    BorderWidth =1
                    Left =4485
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line105"
                End
                Begin Line
                    BorderWidth =1
                    Left =2715
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line107"
                End
                Begin Line
                    BorderWidth =1
                    Left =1215
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line108"
                End
                Begin Line
                    BorderWidth =2
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line109"
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2745
                    Top =15
                    Width =804
                    Height =405
                    FontSize =8
                    TabIndex =7
                    BorderColor =32768
                    Name ="Text154"
                    ControlSource ="MEGH_MAR"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    BorderWidth =1
                    Left =3600
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line156"
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            CanGrow = NotDefault
            Height =5896
            OnFormat ="[Event Procedure]"
            Name ="GroupFooter3"
            Begin
                Begin Line
                    BorderWidth =3
                    Top =30
                    Width =10555
                    BorderColor =0
                    Name ="Line110"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3518
                    Top =95
                    Width =5796
                    Height =750
                    FontSize =10
                    BorderColor =32768
                    Name ="Text111"
                    ControlSource ="=ALPHANUM([GHABEL]) & \" \" & \"ريال\""
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =9297
                            Top =83
                            Width =1200
                            Height =405
                            FontSize =10
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label112"
                            Caption ="مبلغ به حروف:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    BackStyle =1
                    NumeralShapes =2
                    Left =60
                    Top =2235
                    Width =1761
                    Height =435
                    FontSize =8
                    TabIndex =1
                    Name ="NPAR"
                    ControlSource ="=[Forms]![HEAD_LST_FROOSH_BACK2]![NPAR]"
                    Format ="#,##0.00;-#,##0.00"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextAlign =0
                            TextFontFamily =2
                            Left =1876
                            Top =2242
                            Width =1815
                            Height =435
                            FontSize =8
                            FontWeight =400
                            ForeColor =0
                            Name ="Label86"
                            Caption ="جمع مبالغ پرداختي:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    BackStyle =1
                    NumeralShapes =2
                    Left =61
                    Top =90
                    Width =1761
                    Height =435
                    FontSize =8
                    TabIndex =2
                    Name ="JF"
                    ControlSource ="=[Forms]![HEAD_LST_FROOSH_BACK2]![JF]"
                    Format ="#,##0.00;-#,##0.00"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextAlign =0
                            TextFontFamily =2
                            Left =1881
                            Top =90
                            Width =1470
                            Height =435
                            FontSize =8
                            FontWeight =400
                            ForeColor =0
                            Name ="Label115"
                            Caption ="جمع كل فاكتور:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    BackStyle =1
                    NumeralShapes =2
                    Left =56
                    Top =850
                    Width =1761
                    Height =435
                    FontSize =8
                    TabIndex =3
                    Name ="HKH"
                    ControlSource ="=[Forms]![HEAD_LST_FROOSH_BACK2]![HKH]"
                    Format ="#,##0.00;-#,##0.00"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextAlign =0
                            TextFontFamily =2
                            Left =1871
                            Top =850
                            Width =1590
                            Height =435
                            FontSize =8
                            FontWeight =400
                            ForeColor =0
                            Name ="Label116"
                            Caption ="خدمات:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    BackStyle =1
                    NumeralShapes =2
                    Left =60
                    Top =480
                    Width =1761
                    Height =435
                    FontSize =8
                    TabIndex =4
                    Name ="NTKHFIF"
                    ControlSource ="=[Forms]![HEAD_LST_FROOSH_BACK2]![NTKHFIF]"
                    Format ="#,##0.00;-#,##0.00"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextAlign =0
                            TextFontFamily =2
                            Left =1875
                            Top =480
                            Width =780
                            Height =435
                            FontSize =8
                            FontWeight =400
                            ForeColor =0
                            Name ="Label94"
                            Caption ="تخفيفات:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =60
                    Top =1815
                    Width =1761
                    Height =435
                    FontSize =8
                    TabIndex =5
                    BackColor =16777164
                    Name ="GHABEL"
                    ControlSource ="=[Forms]![HEAD_LST_FROOSH_BACK2]![GHABEL]"
                    Format ="#,##0.00;-#,##0.00"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextAlign =0
                            TextFontFamily =2
                            Left =1875
                            Top =1822
                            Width =1695
                            Height =435
                            FontSize =8
                            FontWeight =400
                            ForeColor =0
                            Name ="Label117"
                            Caption ="مبلغ قابل پرداخت:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =60
                    Top =2640
                    Width =1761
                    Height =435
                    FontSize =8
                    TabIndex =6
                    BackColor =16777164
                    Name ="Text98"
                    ControlSource ="=[Forms]![HEAD_LST_FROOSH_BACK2]![MAN]"
                    Format ="#,##0.00;#,##0.00-"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextAlign =0
                            TextFontFamily =2
                            Left =1875
                            Top =2640
                            Width =1305
                            Height =435
                            FontSize =8
                            FontWeight =400
                            ForeColor =0
                            Name ="Label99"
                            Caption ="مانده فاكتور:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =2
                    Left =3465
                    Top =30
                    Width =0
                    Height =3091
                    BorderColor =0
                    Name ="Line118"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =2
                    Width =0
                    Height =3151
                    BorderColor =0
                    Name ="Line119"
                End
                Begin Line
                    BorderWidth =3
                    Top =3120
                    Width =10540
                    BorderColor =0
                    Name ="Line120"
                End
                Begin Line
                    Top =480
                    Width =3450
                    BorderColor =0
                    Name ="Line121"
                End
                Begin Line
                    Top =900
                    Width =3450
                    BorderColor =0
                    Name ="Line122"
                End
                Begin Line
                    Top =1320
                    Width =3450
                    BorderColor =0
                    Name ="Line123"
                End
                Begin Line
                    Top =2265
                    Width =3450
                    BorderColor =0
                    Name ="Line124"
                End
                Begin Line
                    Top =2655
                    Width =3450
                    BorderColor =0
                    Name ="Line125"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =1830
                    Top =30
                    Width =0
                    Height =3106
                    BorderColor =0
                    Name ="Line126"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =3
                    Left =10560
                    Width =0
                    Height =3166
                    BorderColor =0
                    Name ="Line127"
                End
                Begin Line
                    Left =3405
                    Top =900
                    Width =7143
                    BorderColor =0
                    Name ="Line128"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =2
                    Left =9660
                    Top =945
                    Width =840
                    Height =405
                    FontSize =10
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label148"
                    Caption ="توضيحات:"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =2
                    Left =3495
                    Top =945
                    Width =6150
                    Height =1605
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="COMM"
                    Caption ="."
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5430
                    Top =2610
                    Width =3966
                    Height =435
                    FontSize =14
                    TabIndex =7
                    BackColor =128
                    Name ="MANDAH"
                    ControlSource ="=[Forms]![HEAD_LST_FROOSH_BACK2]![MANDAH]"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextAlign =0
                            TextFontFamily =2
                            Left =9405
                            Top =2610
                            Width =1185
                            Height =435
                            FontSize =11
                            FontWeight =400
                            BackColor =128
                            ForeColor =0
                            Name ="Label69"
                            Caption ="مانده حساب:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    BackStyle =1
                    NumeralShapes =2
                    Left =56
                    Top =1350
                    Width =1761
                    Height =435
                    FontSize =8
                    TabIndex =8
                    Name ="Text165"
                    ControlSource ="=[Forms]![HEAD_LST_FROOSH_BACK2]![mbaa]"
                    Format ="#,##0.00;-#,##0.00"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextAlign =0
                            TextFontFamily =2
                            Left =1871
                            Top =1350
                            Width =1590
                            Height =435
                            FontSize =8
                            FontWeight =400
                            ForeColor =0
                            Name ="Label224"
                            Caption ="3% ماليات و عوارض:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin Line
                    Top =1820
                    Width =3450
                    BorderColor =0
                    Name ="Line167"
                End
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =16
                    Top =3118
                    Width =2529
                    Height =330
                    FontSize =10
                    TabIndex =9
                    BorderColor =32768
                    Name ="Text211"
                    ControlSource ="=Format(Farsidate(Now()),\"00\\/00\\/00\") & \" - \" & Format(Now(),\"hh:nn:ss a"
                        "mpm\")"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Top =3458
                    Width =10593
                    Height =540
                    FontSize =12
                    TabIndex =10
                    BorderColor =32768
                    Name ="Text153"
                    ControlSource ="=[Forms]![Baseknow]![HIGH_D]"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Top =3968
                    TabIndex =11
                    Name ="SGN1"
                    ControlSource ="SGN1"
                End
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Top =4287
                    TabIndex =12
                    Name ="SGN2"
                    ControlSource ="SGN2"
                End
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Top =4606
                    TabIndex =13
                    Name ="SGN3"
                    ControlSource ="SGN3"
                End
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Top =4761
                    TabIndex =14
                    Name ="SGN1usid"
                    ControlSource ="SGN1usid"
                End
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Top =4881
                    TabIndex =15
                    Name ="sgn2usid"
                    ControlSource ="sgn2usid"
                End
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Top =5106
                    TabIndex =16
                    Name ="sgn3usid"
                    ControlSource ="sgn3usid"
                End
                Begin BoundObjectFrame
                    OldBorderStyle =0
                    Left =8059
                    Top =4091
                    Width =1701
                    Height =1701
                    TabIndex =17
                    Name ="EMZA1"
                    ControlSource ="EMZA1"
                End
                Begin BoundObjectFrame
                    OldBorderStyle =0
                    Left =5402
                    Top =4088
                    Width =1701
                    Height =1701
                    TabIndex =18
                    Name ="EMZA2"
                    ControlSource ="EMZA2"
                End
                Begin BoundObjectFrame
                    OldBorderStyle =0
                    Left =2777
                    Top =4088
                    Width =1701
                    Height =1701
                    TabIndex =19
                    Name ="EMZA3"
                    ControlSource ="EMZA3"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =8894
                    Top =4414
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
                    Left =6302
                    Top =4421
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
                    Left =3647
                    Top =4422
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
                    Left =8891
                    Top =4081
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
                    Left =6304
                    Top =4088
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
                    Left =3632
                    Top =4089
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
            End
        End
    End
End
CodeBehindForm
' See "INVOICE_FROOSH_BACK2.cls"
