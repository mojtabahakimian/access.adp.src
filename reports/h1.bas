Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    FilterOn = NotDefault
    DefaultView =0
    AllowUpdating =4
    TabularCharSet =178
    TabularFamily =42
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =11001
    DatasheetFontHeight =10
    ItemSuffix =197
    Left =930
    Top =480
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xb3c2f33533ede540
    End
    RecordSource ="SELECT     dbo.HEAD_LST.NUMBER, dbo.HEAD_LST.TAG, dbo.HEAD_LST.DATE_N, dbo.HEAD_"
        "LST.CUST_NO, dbo.INVO_LST.ANBAR, dbo.HEAD_LST.NUMBER1, \015\012                 "
        "     dbo.HEAD_LST.TAH, dbo.HEAD_LST.MAS, dbo.HEAD_LST.VAS, dbo.HEAD_LST.N_S, dbo"
        ".HEAD_LST.M_NAGHD, dbo.HEAD_LST.MABL_VAR, \015\012                      dbo.HEAD"
        "_LST.MOIN_VAR, dbo.HEAD_LST.MABL_HAV, dbo.HEAD_LST.MOIN_HAV, dbo.HEAD_LST.MABL_H"
        "AZ, dbo.HEAD_LST.MOIN_HAZ, \015\012                      dbo.HEAD_LST.TAKHFIF, d"
        "bo.HEAD_LST.MOIN_KHF, dbo.HEAD_LST.ANBARF, dbo.HEAD_LST.FNUMCO, dbo.INVO_LST.NUM"
        "BER AS INUMBER, \015\012                      dbo.INVO_LST.TAG AS ITAG, dbo.INVO"
        "_LST.ANBAR AS IANBAR, dbo.INVO_LST.RADIF, dbo.INVO_LST.CODE, dbo.INVO_LST.MEGH, "
        "\015\012                      dbo.INVO_LST.MEGHk, dbo.INVO_LST.MEGH_MAR, dbo.INV"
        "O_LST.MANDAH, dbo.INVO_LST.MABL, dbo.INVO_LST.MABL_K, dbo.INVO_LST.FROM_A, \015\012"
        "                      dbo.INVO_LST.N_RASID, dbo.INVO_LST.MEGH_R, dbo.INVO_LST.CU"
        "ST_NO AS ICUST_NO, dbo.CUST_HESAB.NAME, dbo.CUST_HESAB.ADDRESS, \015\012        "
        "              dbo.CUST_HESAB.TEL, dbo.TCOD_ANBAR.NAMES, ISNULL(dbo.STUF_DEF.NAME"
        ", N' ') + N'-' + ISNULL(dbo.STUF_DEF.CODE, N' ') \015\012                      +"
        " N' ' + ISNULL(dbo.HIDVISIT(dbo.INVO_LST.MANDAH), N' ') AS KALA, dbo.TCOD_VAHEDS"
        ".NAMES AS VNAMES, dbo.HEAD_LST.MOLAH, \015\012                      dbo.CUST_HES"
        "AB.hes, dbo.DEPART.DEPNAME, dbo.HEAD_LST.USER_NAME, dbo.SHIFT.SHNAME, ISNULL(dbo"
        ".STUF_DEF.NAME, N' ') \015\012                      + N'-' + ISNULL(dbo.HIDVISIT"
        "(dbo.INVO_LST.MANDAH), N' ') AS kala2, dbo.STUF_DEF.RADAH, dbo.INVO_LST.N_KOL, d"
        "bo.HEAD_LST.HMBAA, \015\012                      dbo.HEAD_LST.MBAA, dbo.CUST_HES"
        "AB.CODE_E, dbo.CUST_HESAB.ECODE, dbo.CUST_HESAB.PCODE, dbo.CUST_HESAB.IYALAT, \015"
        "\012                      dbo.CUST_HESAB.CITY, dbo.CUST_HESAB.MCODEM, dbo.INVO_L"
        "ST.N_MOIN, dbo.INVO_LST.IMBAA, \015\012                      dbo.INVO_LST.MABL_K"
        " - dbo.INVO_LST.N_MOIN AS mabkbt, dbo.INVO_LST.MABL_K - dbo.INVO_LST.N_MOIN + db"
        "o.INVO_LST.IMBAA AS mabkn, \015\012                      dbo.CUST_HESAB.NAME AS "
        "HESAB, dbo.STUF_DEF.N_FANI, dbo.OTHER_DTL.REQUEST_NO, dbo.OTHER_DTL.BARNAMEH, db"
        "o.OTHER_DTL.DRIVER, \015\012                      dbo.OTHER_DTL.DRIVER_MOB, dbo."
        "OTHER_DTL.CAMIUN_NUM, dbo.OTHER_DTL.MAGHSAD, dbo.OTHER_DTL.CAM_KHALY, \015\012  "
        "                    dbo.OTHER_DTL.CAM_POOR, dbo.OTHER_DTL.TOZIH, CAST(dbo.STUF_D"
        "EF.CODE AS bigint) AS CODEn, dbo.INVO_LST.SANAD_NO, \015\012                    "
        "  dbo.CUST_HESAB.MOBILE, dbo.INVO_LST.TKHN, dbo.HEAD_LST.SHARAYET, dbo.STUF_FSK."
        "POSITION, dbo.HEAD_LST.sgn1usid, \015\012                      dbo.HEAD_LST.sgn2"
        "usid, dbo.HEAD_LST.sgn3usid, dbo.SALA_DTL.EMZA AS EMZA1, dbo.SALA_DTL.EMZA AS EM"
        "ZA2, SALA_DTL_1.EMZA AS EMZA3, \015\012                      dbo.HEAD_LST.SGN1, "
        "dbo.HEAD_LST.SGN2, dbo.HEAD_LST.SGN3\015\012FROM         dbo.STUF_DEF INNER JOIN"
        "\015\012                      dbo.TCOD_VAHEDS INNER JOIN\015\012                "
        "      dbo.DEPART INNER JOIN\015\012                      dbo.HEAD_LST INNER JOIN"
        "\015\012                      dbo.INVO_LST ON dbo.HEAD_LST.NUMBER = dbo.INVO_LST"
        ".NUMBER AND dbo.HEAD_LST.TAG = dbo.INVO_LST.TAG ON \015\012                     "
        " dbo.DEPART.DEPATMAN = dbo.HEAD_LST.DEPATMAN ON dbo.TCOD_VAHEDS.CODE = dbo.INVO_"
        "LST.VAHED_K ON \015\012                      dbo.STUF_DEF.CODE = dbo.INVO_LST.CO"
        "DE INNER JOIN\015\012                      dbo.TCOD_ANBAR ON dbo.INVO_LST.ANBAR "
        "= dbo.TCOD_ANBAR.CODE INNER JOIN\015\012                      dbo.CUST_HESAB ON "
        "dbo.HEAD_LST.CUST_NO = dbo.CUST_HESAB.hes INNER JOIN\015\012                    "
        "  dbo.SHIFT ON dbo.HEAD_LST.SHIFT = dbo.SHIFT.SHIFT_ID INNER JOIN\015\012       "
        "               dbo.STUF_FSK ON dbo.INVO_LST.CODE = dbo.STUF_FSK.CODE AND dbo.INV"
        "O_LST.ANBAR = dbo.STUF_FSK.ANBAR LEFT OUTER JOIN\015\012                      db"
        "o.SALA_DTL SALA_DTL_2 ON dbo.HEAD_LST.sgn3usid = SALA_DTL_2.IDD LEFT OUTER JOIN\015"
        "\012                      dbo.SALA_DTL SALA_DTL_1 ON dbo.HEAD_LST.sgn2usid = SAL"
        "A_DTL_1.IDD LEFT OUTER JOIN\015\012                      dbo.SALA_DTL ON dbo.HEA"
        "D_LST.sgn1usid = dbo.SALA_DTL.IDD LEFT OUTER JOIN\015\012                      d"
        "bo.OTHER_DTL ON dbo.HEAD_LST.NUMBER = dbo.OTHER_DTL.NUMBER AND dbo.HEAD_LST.TAG "
        "= dbo.OTHER_DTL.TAG"
    Caption ="حواله خروج كالا از انبار"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial (Arabic)"
    OnError ="[Event Procedure]"
    RecordSource ="SELECT     dbo.HEAD_LST.NUMBER, dbo.HEAD_LST.TAG, dbo.HEAD_LST.DATE_N, dbo.HEAD_"
        "LST.CUST_NO, dbo.INVO_LST.ANBAR, dbo.HEAD_LST.NUMBER1, \015\012                 "
        "     dbo.HEAD_LST.TAH, dbo.HEAD_LST.MAS, dbo.HEAD_LST.VAS, dbo.HEAD_LST.N_S, dbo"
        ".HEAD_LST.M_NAGHD, dbo.HEAD_LST.MABL_VAR, \015\012                      dbo.HEAD"
        "_LST.MOIN_VAR, dbo.HEAD_LST.MABL_HAV, dbo.HEAD_LST.MOIN_HAV, dbo.HEAD_LST.MABL_H"
        "AZ, dbo.HEAD_LST.MOIN_HAZ, \015\012                      dbo.HEAD_LST.TAKHFIF, d"
        "bo.HEAD_LST.MOIN_KHF, dbo.HEAD_LST.ANBARF, dbo.HEAD_LST.FNUMCO, dbo.INVO_LST.NUM"
        "BER AS INUMBER, \015\012                      dbo.INVO_LST.TAG AS ITAG, dbo.INVO"
        "_LST.ANBAR AS IANBAR, dbo.INVO_LST.RADIF, dbo.INVO_LST.CODE, dbo.INVO_LST.MEGH, "
        "\015\012                      dbo.INVO_LST.MEGHk, dbo.INVO_LST.MEGH_MAR, dbo.INV"
        "O_LST.MANDAH, dbo.INVO_LST.MABL, dbo.INVO_LST.MABL_K, dbo.INVO_LST.FROM_A, \015\012"
        "                      dbo.INVO_LST.N_RASID, dbo.INVO_LST.MEGH_R, dbo.INVO_LST.CU"
        "ST_NO AS ICUST_NO, dbo.CUST_HESAB.NAME, dbo.CUST_HESAB.ADDRESS, \015\012        "
        "              dbo.CUST_HESAB.TEL, dbo.TCOD_ANBAR.NAMES, ISNULL(dbo.STUF_DEF.NAME"
        ", N' ') + N'-' + ISNULL(dbo.STUF_DEF.CODE, N' ') \015\012                      +"
        " N' ' + ISNULL(dbo.HIDVISIT(dbo.INVO_LST.MANDAH), N' ') AS KALA, dbo.TCOD_VAHEDS"
        ".NAMES AS VNAMES, dbo.HEAD_LST.MOLAH, \015\012                      dbo.CUST_HES"
        "AB.hes, dbo.DEPART.DEPNAME, dbo.HEAD_LST.USER_NAME, dbo.SHIFT.SHNAME, ISNULL(dbo"
        ".STUF_DEF.NAME, N' ') \015\012                      + N'-' + ISNULL(dbo.HIDVISIT"
        "(dbo.INVO_LST.MANDAH), N' ') AS kala2, dbo.STUF_DEF.RADAH, dbo.INVO_LST.N_KOL, d"
        "bo.HEAD_LST.HMBAA, \015\012                      dbo.HEAD_LST.MBAA, dbo.CUST_HES"
        "AB.CODE_E, dbo.CUST_HESAB.ECODE, dbo.CUST_HESAB.PCODE, dbo.CUST_HESAB.IYALAT, \015"
        "\012                      dbo.CUST_HESAB.CITY, dbo.CUST_HESAB.MCODEM, dbo.INVO_L"
        "ST.N_MOIN, dbo.INVO_LST.IMBAA, \015\012                      dbo.INVO_LST.MABL_K"
        " - dbo.INVO_LST.N_MOIN AS mabkbt, dbo.INVO_LST.MABL_K - dbo.INVO_LST.N_MOIN + db"
        "o.INVO_LST.IMBAA AS mabkn, \015\012                      dbo.CUST_HESAB.NAME AS "
        "HESAB, dbo.STUF_DEF.N_FANI, dbo.OTHER_DTL.REQUEST_NO, dbo.OTHER_DTL.BARNAMEH, db"
        "o.OTHER_DTL.DRIVER, \015\012                      dbo.OTHER_DTL.DRIVER_MOB, dbo."
        "OTHER_DTL.CAMIUN_NUM, dbo.OTHER_DTL.MAGHSAD, dbo.OTHER_DTL.CAM_KHALY, \015\012  "
        "                    dbo.OTHER_DTL.CAM_POOR, dbo.OTHER_DTL.TOZIH, CAST(dbo.STUF_D"
        "EF.CODE AS bigint) AS CODEn, dbo.INVO_LST.SANAD_NO, \015\012                    "
        "  dbo.CUST_HESAB.MOBILE, dbo.INVO_LST.TKHN, dbo.HEAD_LST.SHARAYET, dbo.STUF_FSK."
        "POSITION, dbo.HEAD_LST.sgn1usid, \015\012                      dbo.HEAD_LST.sgn2"
        "usid, dbo.HEAD_LST.sgn3usid, dbo.SALA_DTL.EMZA AS EMZA1, dbo.SALA_DTL.EMZA AS EM"
        "ZA2, SALA_DTL_1.EMZA AS EMZA3, \015\012                      dbo.HEAD_LST.SGN1, "
        "dbo.HEAD_LST.SGN2, dbo.HEAD_LST.SGN3\015\012FROM         dbo.STUF_DEF INNER JOIN"
        "\015\012                      dbo.TCOD_VAHEDS INNER JOIN\015\012                "
        "      dbo.DEPART INNER JOIN\015\012                      dbo.HEAD_LST INNER JOIN"
        "\015\012                      dbo.INVO_LST ON dbo.HEAD_LST.NUMBER = dbo.INVO_LST"
        ".NUMBER AND dbo.HEAD_LST.TAG = dbo.INVO_LST.TAG ON \015\012                     "
        " dbo.DEPART.DEPATMAN = dbo.HEAD_LST.DEPATMAN ON dbo.TCOD_VAHEDS.CODE = dbo.INVO_"
        "LST.VAHED_K ON \015\012                      dbo.STUF_DEF.CODE = dbo.INVO_LST.CO"
        "DE INNER JOIN\015\012                      dbo.TCOD_ANBAR ON dbo.INVO_LST.ANBAR "
        "= dbo.TCOD_ANBAR.CODE INNER JOIN\015\012                      dbo.CUST_HESAB ON "
        "dbo.HEAD_LST.CUST_NO = dbo.CUST_HESAB.hes INNER JOIN\015\012                    "
        "  dbo.SHIFT ON dbo.HEAD_LST.SHIFT = dbo.SHIFT.SHIFT_ID INNER JOIN\015\012       "
        "               dbo.STUF_FSK ON dbo.INVO_LST.CODE = dbo.STUF_FSK.CODE AND dbo.INV"
        "O_LST.ANBAR = dbo.STUF_FSK.ANBAR LEFT OUTER JOIN\015\012                      db"
        "o.SALA_DTL SALA_DTL_2 ON dbo.HEAD_LST.sgn3usid = SALA_DTL_2.IDD LEFT OUTER JOIN\015"
        "\012                      dbo.SALA_DTL SALA_DTL_1 ON dbo.HEAD_LST.sgn2usid = SAL"
        "A_DTL_1.IDD LEFT OUTER JOIN\015\012                      dbo.SALA_DTL ON dbo.HEA"
        "D_LST.sgn1usid = dbo.SALA_DTL.IDD LEFT OUTER JOIN\015\012                      d"
        "bo.OTHER_DTL ON dbo.HEAD_LST.NUMBER = dbo.OTHER_DTL.NUMBER AND dbo.HEAD_LST.TAG "
        "= dbo.OTHER_DTL.TAG"
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
            SortOrder = NotDefault
            GroupFooter = NotDefault
            KeepTogether =1
            ControlSource ="NUMBER"
        End
        Begin PageHeader
            Height =3045
            Name ="PageHeader"
            Begin
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =3855
                    Width =2550
                    Height =540
                    FontSize =14
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label88"
                    Caption ="حواله خروج كالا از انبار"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2265
                    Top =495
                    Width =5613
                    Height =540
                    FontSize =14
                    BorderColor =32768
                    Name ="Text90"
                    ControlSource ="=Forms!Baseknow!WIDTH_D"
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
                    Top =540
                    Width =1506
                    Height =450
                    FontSize =10
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
                            Left =1603
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
                    Left =7875
                    Top =510
                    Width =1686
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
                            Left =9630
                            Top =510
                            Width =1260
                            Height =480
                            FontSize =12
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label93"
                            Caption ="شماره حواله:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin Line
                    BorderWidth =3
                    Top =1050
                    Width =11001
                    BorderColor =0
                    Name ="Line94"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2265
                    Top =1107
                    Width =7836
                    Height =450
                    FontSize =12
                    TabIndex =3
                    BorderColor =32768
                    Name ="moshtary"
                    ControlSource ="NAME"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =10137
                            Top =1107
                            Width =780
                            Height =480
                            FontSize =12
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label95"
                            Caption ="خريدار:"
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
                    Left =120
                    Top =1095
                    Width =1371
                    Height =450
                    FontSize =12
                    TabIndex =4
                    BorderColor =32768
                    Name ="CUST_NO"
                    ControlSource ="CUST_NO"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =1482
                            Top =1095
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
                    Left =2265
                    Top =1569
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
                            Left =10122
                            Top =1569
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
                    Top =1560
                    Width =1401
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
                            Left =1482
                            Top =1560
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
                    Top =2520
                    Width =11001
                    BorderColor =0
                    Name ="Line99"
                End
                Begin Line
                    BorderWidth =2
                    Left =10965
                    Top =2535
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line130"
                End
                Begin Line
                    Left =10395
                    Top =2535
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line131"
                End
                Begin Line
                    Left =8205
                    Top =2535
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line132"
                End
                Begin Line
                    Left =3585
                    Top =2535
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line133"
                End
                Begin Line
                    Left =2755
                    Top =2535
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line134"
                End
                Begin Line
                    Left =1035
                    Top =2535
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line135"
                End
                Begin Line
                    BorderWidth =2
                    Top =2535
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line138"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =180
                    Top =2580
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
                    Left =2932
                    Top =2580
                    Width =525
                    Height =405
                    FontSize =10
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label142"
                    Caption ="مقدار"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =1260
                    Top =2610
                    Width =1455
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
                    Left =4687
                    Top =2580
                    Width =2070
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
                    Left =8280
                    Top =2580
                    Width =1980
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
                    Left =10401
                    Top =2576
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
                    Top =3045
                    Width =11001
                    BorderColor =0
                    Name ="Line147"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =60
                    Top =2040
                    Width =9816
                    Height =450
                    FontSize =12
                    TabIndex =7
                    BorderColor =32768
                    Name ="MOLAH1"
                    ControlSource ="=IIf(IsNull([MOLAH]),[MOLAH],HIDVISIT([MOLAH]))"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =9900
                            Top =2040
                            Width =1035
                            Height =480
                            FontSize =12
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label155"
                            Caption ="ملاحظات:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =8955
                    Width =921
                    Height =450
                    FontSize =12
                    TabIndex =8
                    BorderColor =32768
                    Name ="MOLAH"
                    ControlSource ="MOLAH"
                    FontName ="WeblogmaYekan"
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            CanGrow = NotDefault
            Height =450
            OnFormat ="[Event Procedure]"
            Name ="Detail"
            Begin
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =8237
                    Top =16
                    Width =2109
                    Height =405
                    FontSize =12
                    BorderColor =32768
                    Name ="ANBAR"
                    ControlSource ="NAMES"
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
                    Left =10411
                    Top =16
                    Width =504
                    Height =405
                    FontSize =12
                    TabIndex =1
                    BorderColor =32768
                    Name ="RADIF"
                    ControlSource ="=1"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3597
                    Top =16
                    Width =4569
                    Height =405
                    FontSize =12
                    TabIndex =2
                    BorderColor =32768
                    Name ="CODE"
                    ControlSource ="=IIf([Forms]![baseknow]![TFTPAGE]=1,[kala],[kala2])"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2785
                    Top =16
                    Width =789
                    Height =405
                    FontSize =12
                    TabIndex =3
                    BorderColor =32768
                    Name ="MEGH"
                    ControlSource ="MEGH"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =105
                    Top =16
                    Width =894
                    Height =405
                    FontSize =12
                    TabIndex =4
                    BorderColor =32768
                    Name ="MEGHk"
                    ControlSource ="MEGHk"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1080
                    Top =15
                    Width =1650
                    Height =405
                    FontSize =12
                    TabIndex =5
                    BorderColor =32768
                    Name ="VAHED_K"
                    ControlSource ="VNAMES"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Top =450
                    Width =10986
                    BorderColor =0
                    Name ="Line100"
                End
                Begin Line
                    BorderWidth =2
                    Left =10965
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line101"
                End
                Begin Line
                    Left =10395
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line102"
                End
                Begin Line
                    Left =8205
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line103"
                End
                Begin Line
                    Left =3585
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line104"
                End
                Begin Line
                    Left =2755
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line105"
                End
                Begin Line
                    Left =1035
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line106"
                End
                Begin Line
                    BorderWidth =2
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line109"
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =2680
            OnFormat ="[Event Procedure]"
            Name ="GroupFooter3"
            Begin
                Begin Line
                    BorderWidth =3
                    Top =30
                    Width =10990
                    BorderColor =0
                    Name ="Line110"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =1643
                    Top =566
                    Width =2475
                    Height =405
                    FontSize =10
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label191"
                    Caption ="نام ونام خانوادگي انبار دار:"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =6402
                    Top =566
                    Width =3015
                    Height =405
                    FontSize =10
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label192"
                    Caption ="نام و نام خانوادگي تحويل گيرنده:"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =5782
                    Top =566
                    Width =570
                    Height =405
                    FontSize =10
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label193"
                    Caption ="امضاء "
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =907
                    Top =566
                    Width =570
                    Height =405
                    FontSize =10
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label194"
                    Caption ="امضاء "
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =1
                    BorderWidth =2
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =15
                    Width =1000
                    Height =390
                    FontSize =11
                    Name ="Text202"
                    ControlSource ="=Sum([MEGHk])"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =9312
                    Top =64
                    Width =1086
                    Height =345
                    TabIndex =1
                    Name ="USERNAME"
                    ControlSource ="=ucurrentuser()"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextAlign =0
                            TextFontFamily =2
                            Left =10439
                            Top =56
                            Width =480
                            Height =375
                            FontWeight =400
                            Name ="Label205"
                            Caption ="كاربر:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =7540
                    Top =64
                    Width =1206
                    Height =345
                    TabIndex =2
                    Name ="Text206"
                    ControlSource ="=Now()"
                    Format ="Short Time"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextAlign =0
                            TextFontFamily =2
                            Left =8782
                            Top =59
                            Width =510
                            Height =375
                            FontWeight =400
                            Name ="Label207"
                            Caption ="ساعت:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =2
                    Left =3004
                    Top =113
                    Width =3285
                    Height =345
                    FontSize =10
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="vazn"
                    FontName ="WeblogmaYekan"
                End
                Begin BoundObjectFrame
                    SizeMode =3
                    OldBorderStyle =0
                    Left =6840
                    Top =979
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
                    Top =964
                    Width =1701
                    Height =1701
                    TabIndex =4
                    Name ="EMZA2"
                    ControlSource ="EMZA2"
                End
                Begin BoundObjectFrame
                    SizeMode =3
                    OldBorderStyle =0
                    Top =960
                    Width =1701
                    Height =1701
                    TabIndex =5
                    Name ="EMZA3"
                    ControlSource ="EMZA3"
                End
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Left =8220
                    Top =963
                    TabIndex =6
                    Name ="SGN1usid"
                    ControlSource ="SGN1usid"
                End
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Left =8220
                    Top =1282
                    TabIndex =7
                    Name ="sgn2usid"
                    ControlSource ="sgn2usid"
                End
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Left =7361
                    Top =1318
                    TabIndex =8
                    Name ="sgn3usid"
                    ControlSource ="sgn3usid"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =8550
                    Top =1399
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
                    Top =1399
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
                    Top =1399
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
                    Top =1066
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
                    Top =1066
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
                    Top =1066
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
                    Left =1814
                    Top =1757
                    TabIndex =9
                    Name ="SGN1"
                    ControlSource ="SGN1"
                End
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Left =1814
                    Top =2076
                    TabIndex =10
                    Name ="SGN2"
                    ControlSource ="SGN2"
                End
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Left =1814
                    Top =2395
                    TabIndex =11
                    Name ="SGN3"
                    ControlSource ="SGN3"
                End
            End
        End
        Begin PageFooter
            Visible = NotDefault
            Height =0
            Name ="PageFooter"
        End
    End
End
CodeBehindForm
' See "h1.cls"
