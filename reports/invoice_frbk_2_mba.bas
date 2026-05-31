Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    FilterOn = NotDefault
    DefaultView =0
    AllowUpdating =4
    TabularCharSet =178
    TabularFamily =42
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =15212
    DatasheetFontHeight =10
    ItemSuffix =298
    Left =600
    Top =390
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x5cdaf9e513f9e540
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
    DatasheetFontName ="Arial (Arabic)"
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
        Begin PageBreak
            Width =283
        End
        Begin BreakLevel
            GroupFooter = NotDefault
            KeepTogether =1
            ControlSource ="NUMBER"
        End
        Begin BreakLevel
            ControlSource ="RADIF"
        End
        Begin PageHeader
            Height =4158
            Name ="PageHeader"
            Begin
                Begin TextBox
                    TextFontCharSet =0
                    TextAlign =3
                    TextFontFamily =34
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =75
                    Top =405
                    Width =1191
                    Height =375
                    FontSize =8
                    BorderColor =32768
                    Name ="DATE_N"
                    ControlSource ="DATE_N"
                    FontName ="Tahoma"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            TextFontCharSet =0
                            TextAlign =3
                            TextFontFamily =34
                            Left =1318
                            Top =405
                            Width =570
                            Height =390
                            FontSize =8
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label92"
                            Caption ="تاريخ :"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextFontCharSet =0
                    TextAlign =3
                    TextFontFamily =34
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =75
                    Top =60
                    Width =1191
                    Height =375
                    FontSize =8
                    TabIndex =1
                    BorderColor =32768
                    Name ="NUMBER1"
                    ControlSource ="NUMBER"
                    FontName ="Tahoma"
                    Begin
                        Begin Label
                            TextFontCharSet =0
                            TextAlign =3
                            TextFontFamily =34
                            Left =1302
                            Top =60
                            Width =1065
                            Height =390
                            FontSize =8
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label93"
                            Caption ="شماره سريال:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    TextFontCharSet =0
                    TextAlign =3
                    TextFontFamily =34
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =6600
                    Top =2220
                    Width =6771
                    Height =270
                    FontSize =7
                    TabIndex =2
                    BorderColor =32768
                    Name ="HESAB"
                    ControlSource ="NAME"
                    FontName ="Tahoma"
                    Begin
                        Begin Label
                            TextFontCharSet =0
                            TextAlign =3
                            TextFontFamily =34
                            Left =13455
                            Top =2220
                            Width =1650
                            Height =270
                            FontSize =7
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label95"
                            Caption ="نام شخص حقيقي/حقوقي:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    TextFontCharSet =0
                    TextAlign =3
                    TextFontFamily =34
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3480
                    Top =2835
                    Width =11001
                    Height =375
                    FontSize =7
                    TabIndex =3
                    BorderColor =32768
                    Name ="ADDRESS"
                    ControlSource ="ADDRESS"
                    FontName ="Tahoma"
                    Begin
                        Begin Label
                            TextFontCharSet =0
                            TextAlign =3
                            TextFontFamily =34
                            Left =14517
                            Top =2835
                            Width =630
                            Height =390
                            FontSize =7
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label97"
                            Caption ="نشاني:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    TextFontCharSet =0
                    TextAlign =3
                    TextFontFamily =34
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =75
                    Top =2822
                    Width =1911
                    Height =375
                    FontSize =7
                    TabIndex =4
                    BorderColor =32768
                    Name ="TEL"
                    ControlSource ="TEL"
                    FontName ="Tahoma"
                    Begin
                        Begin Label
                            TextFontCharSet =0
                            TextAlign =3
                            TextFontFamily =34
                            Left =2069
                            Top =2820
                            Width =1260
                            Height =390
                            FontSize =7
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label98"
                            Caption ="شماره تلفن/نمابر:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin Line
                    Left =14520
                    Top =3495
                    Width =0
                    Height =663
                    BorderColor =0
                    Name ="Line131"
                End
                Begin Line
                    Left =13095
                    Top =3495
                    Width =0
                    Height =663
                    BorderColor =0
                    Name ="Line132"
                End
                Begin Line
                    Left =6120
                    Top =3495
                    Width =0
                    Height =663
                    BorderColor =0
                    Name ="Line133"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =4815
                    Top =3495
                    Width =0
                    Height =663
                    BorderColor =0
                    Name ="Line134"
                End
                Begin Line
                    Left =3180
                    Top =3495
                    Width =0
                    Height =663
                    BorderColor =0
                    Name ="Line135"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =1590
                    Top =3495
                    Width =0
                    Height =663
                    BorderColor =0
                    Name ="Line137"
                End
                Begin Label
                    TextFontCharSet =0
                    TextAlign =2
                    TextFontFamily =34
                    Left =330
                    Top =3555
                    Width =855
                    Height =405
                    FontSize =8
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label139"
                    Caption ="مبــلغ كل"
                    FontName ="Tahoma"
                End
                Begin Label
                    TextFontCharSet =0
                    TextAlign =2
                    TextFontFamily =34
                    Left =2124
                    Top =3465
                    Width =615
                    Height =405
                    FontSize =8
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label140"
                    Caption ="مبلغ"
                    FontName ="Tahoma"
                End
                Begin Label
                    TextFontCharSet =0
                    TextAlign =2
                    TextFontFamily =34
                    Left =5246
                    Top =3465
                    Width =690
                    Height =375
                    FontSize =8
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label142"
                    Caption ="تعداد/"
                    FontName ="Tahoma"
                End
                Begin Label
                    TextFontCharSet =0
                    TextAlign =2
                    TextFontFamily =34
                    Left =3265
                    Top =3690
                    Width =1530
                    Height =285
                    FontSize =8
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label143"
                    Caption ="اندازه گيري"
                    FontName ="Tahoma"
                End
                Begin Label
                    TextFontCharSet =0
                    TextAlign =2
                    TextFontFamily =34
                    Left =7880
                    Top =3555
                    Width =4485
                    Height =405
                    FontSize =8
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label144"
                    Caption ="شــــــــــرح  كالا يا خدمات"
                    FontName ="Tahoma"
                End
                Begin Label
                    TextFontCharSet =0
                    TextAlign =2
                    TextFontFamily =34
                    Left =13140
                    Top =3555
                    Width =1320
                    Height =405
                    FontSize =8
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label145"
                    Caption ="كد كالا"
                    FontName ="Tahoma"
                End
                Begin Label
                    TextFontCharSet =0
                    TextAlign =2
                    TextFontFamily =34
                    Left =14571
                    Top =3551
                    Width =615
                    Height =405
                    FontSize =8
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label146"
                    Caption ="رديف"
                    FontName ="Tahoma"
                End
                Begin TextBox
                    TextFontCharSet =0
                    TextAlign =3
                    TextFontFamily =34
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =6598
                    Top =802
                    Width =6516
                    Height =375
                    FontSize =7
                    TabIndex =5
                    BorderColor =32768
                    Name ="Text156"
                    ControlSource ="=[Forms]![Baseknow]![NAME]"
                    FontName ="Tahoma"
                    Begin
                        Begin Label
                            TextFontCharSet =0
                            TextAlign =3
                            TextFontFamily =34
                            Left =13171
                            Top =810
                            Width =1890
                            Height =390
                            FontSize =7
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label157"
                            Caption ="نام شخص حقيقي/حقوقي:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    TextFontCharSet =0
                    TextAlign =3
                    TextFontFamily =34
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3525
                    Top =1515
                    Width =10956
                    Height =375
                    FontSize =7
                    TabIndex =6
                    BorderColor =32768
                    Name ="Text160"
                    ControlSource ="=[Forms]![Baseknow]![TFADDRESS]"
                    FontName ="Tahoma"
                    Begin
                        Begin Label
                            TextFontCharSet =0
                            TextAlign =3
                            TextFontFamily =34
                            Left =14502
                            Top =1515
                            Width =570
                            Height =390
                            FontSize =7
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label161"
                            Caption ="نشاني:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    TextFontCharSet =0
                    TextAlign =3
                    TextFontFamily =34
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =75
                    Top =1517
                    Width =1926
                    Height =375
                    FontSize =7
                    TabIndex =7
                    BorderColor =32768
                    Name ="Text162"
                    ControlSource ="=[Forms]![Baseknow]![TFTEL]"
                    FontName ="Tahoma"
                    Begin
                        Begin Label
                            TextFontCharSet =0
                            TextAlign =3
                            TextFontFamily =34
                            Left =2074
                            Top =1515
                            Width =1350
                            Height =390
                            FontSize =7
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label163"
                            Caption ="شماره تلفن/نمابر:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    TextFontCharSet =0
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =11181
                    Top =623
                    Width =666
                    Height =255
                    FontSize =8
                    TabIndex =8
                    Name ="HTAG"
                    ControlSource ="HTAG"
                    FontName ="Tahoma"
                End
                Begin TextBox
                    Visible = NotDefault
                    TextFontCharSet =0
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =6363
                    Top =113
                    FontSize =8
                    TabIndex =9
                    Name ="CODE"
                    ControlSource ="CODE"
                    FontName ="Tahoma"
                End
                Begin Label
                    TextFontCharSet =0
                    TextAlign =2
                    TextFontFamily =34
                    Left =6045
                    Width =2790
                    Height =465
                    FontSize =8
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label88"
                    Caption ="صورتحساب برگشت فروش كالا و خدمات"
                    FontName ="Tahoma"
                End
                Begin Label
                    TextFontCharSet =0
                    TextFontFamily =34
                    Left =6834
                    Top =465
                    Width =1485
                    Height =330
                    FontSize =8
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label204"
                    Caption ="مشخصات فروشنده"
                    FontName ="Tahoma"
                End
                Begin TextBox
                    TextFontCharSet =0
                    TextAlign =3
                    TextFontFamily =34
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3570
                    Top =802
                    Height =375
                    FontSize =7
                    TabIndex =10
                    BorderColor =32768
                    Name ="Text205"
                    ControlSource ="=[Forms]![Baseknow]![ECODE]"
                    FontName ="Tahoma"
                    Begin
                        Begin Label
                            TextFontCharSet =0
                            TextAlign =3
                            TextFontFamily =34
                            Left =5311
                            Top =795
                            Width =1185
                            Height =390
                            FontSize =7
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label207"
                            Caption ="شماره اقتصادي:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    TextFontCharSet =0
                    TextAlign =3
                    TextFontFamily =34
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =68
                    Top =800
                    Width =1476
                    Height =375
                    FontSize =7
                    TabIndex =11
                    BorderColor =32768
                    Name ="Text208"
                    ControlSource ="=[Forms]![Baseknow]![MCODEM]"
                    FontName ="Tahoma"
                    Begin
                        Begin Label
                            TextFontCharSet =0
                            TextAlign =3
                            TextFontFamily =34
                            Left =1635
                            Top =795
                            Width =1500
                            Height =390
                            FontSize =7
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label209"
                            Caption ="شماره ثبت/كد ملي:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    TextFontCharSet =0
                    TextAlign =3
                    TextFontFamily =34
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =12150
                    Top =1134
                    Width =1521
                    Height =375
                    FontSize =7
                    TabIndex =12
                    BorderColor =32768
                    Name ="Text211"
                    ControlSource ="=[Forms]![Baseknow]![IYALAT]"
                    FontName ="Tahoma"
                    Begin
                        Begin Label
                            TextFontCharSet =0
                            TextAlign =3
                            TextFontFamily =34
                            Left =13675
                            Top =1125
                            Width =1395
                            Height =390
                            FontSize =7
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label212"
                            Caption ="نشاني كامل:استان:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    TextFontCharSet =0
                    TextAlign =3
                    TextFontFamily =34
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =9932
                    Top =1124
                    Width =1416
                    Height =375
                    FontSize =7
                    TabIndex =13
                    BorderColor =32768
                    Name ="Text216"
                    ControlSource ="=[Forms]![Baseknow]![CITY]"
                    FontName ="Tahoma"
                    Begin
                        Begin Label
                            TextFontCharSet =0
                            TextAlign =3
                            TextFontFamily =34
                            Left =11352
                            Top =1115
                            Width =765
                            Height =390
                            FontSize =7
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label217"
                            Caption ="شهرستان:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    TextFontCharSet =0
                    TextAlign =3
                    TextFontFamily =34
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3585
                    Top =1119
                    Width =1686
                    Height =375
                    FontSize =7
                    TabIndex =14
                    BorderColor =32768
                    Name ="Text218"
                    ControlSource ="=[Forms]![Baseknow]![PCODE]"
                    FontName ="Tahoma"
                    Begin
                        Begin Label
                            TextFontCharSet =0
                            TextAlign =3
                            TextFontFamily =34
                            Left =5329
                            Top =1110
                            Width =1335
                            Height =390
                            FontSize =7
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label219"
                            Caption ="كد پستي 10رقمي:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin Rectangle
                    BorderWidth =2
                    Left =15
                    Top =737
                    Width =15192
                    Height =1187
                    Name ="Box220"
                End
                Begin Rectangle
                    BorderWidth =2
                    Left =15
                    Top =2160
                    Width =15192
                    Height =1067
                    Name ="Box222"
                End
                Begin Label
                    TextFontCharSet =0
                    TextFontFamily =34
                    Left =6819
                    Top =1875
                    Width =1125
                    Height =330
                    FontSize =7
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label223"
                    Caption ="مشخصات خريدار"
                    FontName ="Tahoma"
                End
                Begin TextBox
                    TextFontCharSet =0
                    TextAlign =3
                    TextFontFamily =34
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3571
                    Top =2220
                    Width =1686
                    Height =270
                    FontSize =7
                    TabIndex =15
                    BorderColor =32768
                    Name ="ECODE"
                    ControlSource ="ECODE"
                    FontName ="Tahoma"
                    Begin
                        Begin Label
                            TextFontCharSet =0
                            TextAlign =3
                            TextFontFamily =34
                            Left =5477
                            Top =2220
                            Width =1005
                            Height =270
                            FontSize =7
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label225"
                            Caption ="شماره اقتصادي:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    TextFontCharSet =0
                    TextAlign =3
                    TextFontFamily =34
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =75
                    Top =2220
                    Width =1461
                    Height =270
                    FontSize =7
                    TabIndex =16
                    BorderColor =32768
                    Name ="MCODEM"
                    ControlSource ="MCODEM"
                    FontName ="Tahoma"
                    Begin
                        Begin Label
                            TextFontCharSet =0
                            TextAlign =3
                            TextFontFamily =34
                            Left =1882
                            Top =2220
                            Width =1245
                            Height =270
                            FontSize =7
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label227"
                            Caption ="شماره ثبت/كد ملي:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    TextFontCharSet =0
                    TextAlign =3
                    TextFontFamily =34
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3593
                    Top =2481
                    Width =1686
                    Height =375
                    FontSize =7
                    TabIndex =17
                    BorderColor =32768
                    Name ="PCODE"
                    ControlSource ="PCODE"
                    FontName ="Tahoma"
                    Begin
                        Begin Label
                            TextFontCharSet =0
                            TextAlign =3
                            TextFontFamily =34
                            Left =5337
                            Top =2472
                            Width =1335
                            Height =390
                            FontSize =7
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label229"
                            Caption ="كد پستي 10رقمي:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    TextFontCharSet =0
                    TextAlign =3
                    TextFontFamily =34
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =12195
                    Top =2544
                    Width =1521
                    Height =375
                    FontSize =7
                    TabIndex =18
                    BorderColor =32768
                    Name ="Text230"
                    ControlSource ="IYALAT"
                    FontName ="Tahoma"
                    Begin
                        Begin Label
                            TextFontCharSet =0
                            TextAlign =3
                            TextFontFamily =34
                            Left =13720
                            Top =2535
                            Width =1395
                            Height =390
                            FontSize =7
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label231"
                            Caption ="نشاني كامل:استان:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    TextFontCharSet =0
                    TextAlign =3
                    TextFontFamily =34
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =9990
                    Top =2544
                    Width =1416
                    Height =375
                    FontSize =7
                    TabIndex =19
                    BorderColor =32768
                    Name ="Text232"
                    ControlSource ="CITY"
                    FontName ="Tahoma"
                    Begin
                        Begin Label
                            TextFontCharSet =0
                            TextAlign =3
                            TextFontFamily =34
                            Left =11410
                            Top =2535
                            Width =765
                            Height =390
                            FontSize =7
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label233"
                            Caption ="شهرستان:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin Label
                    TextFontCharSet =0
                    TextFontFamily =34
                    Left =6240
                    Top =3180
                    Width =2220
                    Height =330
                    FontSize =8
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label236"
                    Caption ="مشخصات كالا يا خدمات مورد معامله"
                    FontName ="Tahoma"
                End
                Begin Rectangle
                    BorderWidth =2
                    Left =15
                    Top =3480
                    Width =15177
                    Height =662
                    Name ="Box237"
                End
                Begin Label
                    TextFontCharSet =0
                    TextAlign =2
                    TextFontFamily =34
                    Left =5220
                    Top =3711
                    Width =735
                    Height =285
                    FontSize =8
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label238"
                    Caption =" مقدار"
                    FontName ="Tahoma"
                End
                Begin Label
                    TextFontCharSet =0
                    TextAlign =2
                    TextFontFamily =34
                    Left =3690
                    Top =3465
                    Width =810
                    Height =375
                    FontSize =8
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label239"
                    Caption ="واحد"
                    FontName ="Tahoma"
                End
                Begin Label
                    TextFontCharSet =0
                    TextAlign =2
                    TextFontFamily =34
                    Left =2094
                    Top =3675
                    Width =645
                    Height =405
                    FontSize =8
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label240"
                    Caption ="واحد"
                    FontName ="Tahoma"
                End
                Begin Label
                    TextFontCharSet =0
                    TextAlign =2
                    TextFontFamily =34
                    Left =2070
                    Top =3870
                    Width =720
                    Height =285
                    FontSize =8
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label241"
                    Caption ="(ريال)"
                    FontName ="Tahoma"
                End
                Begin Label
                    TextFontCharSet =0
                    TextAlign =2
                    TextFontFamily =34
                    Left =420
                    Top =3780
                    Width =705
                    Height =330
                    FontSize =8
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label242"
                    Caption ="(ريال)"
                    FontName ="Tahoma"
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =465
            OnFormat ="[Event Procedure]"
            Name ="Detail"
            Begin
                Begin TextBox
                    DecimalPlaces =0
                    TextFontCharSet =0
                    TextAlign =2
                    TextFontFamily =34
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =13157
                    Top =16
                    Width =1314
                    Height =405
                    FontSize =8
                    BorderColor =32768
                    Name ="CODE11"
                    ControlSource ="CODE"
                    Format ="Fixed"
                    FontName ="Tahoma"
                End
                Begin TextBox
                    DecimalPlaces =0
                    RunningSum =1
                    TextFontCharSet =0
                    TextAlign =2
                    TextFontFamily =34
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =14566
                    Top =16
                    Width =594
                    Height =405
                    FontSize =8
                    TabIndex =1
                    BorderColor =32768
                    Name ="RADIFwe"
                    ControlSource ="=1"
                    FontName ="Tahoma"
                End
                Begin TextBox
                    TextFontCharSet =0
                    TextAlign =3
                    TextFontFamily =34
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =6180
                    Top =16
                    Width =6849
                    Height =405
                    FontSize =8
                    TabIndex =2
                    BorderColor =32768
                    Name ="CODE1"
                    ControlSource ="SNAME"
                    FontName ="Tahoma"
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextFontCharSet =0
                    TextAlign =2
                    TextFontFamily =34
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4875
                    Top =16
                    Width =1179
                    Height =405
                    FontSize =8
                    TabIndex =3
                    BorderColor =32768
                    Name ="MEGH_MAR"
                    ControlSource ="MEGH_MAR"
                    FontName ="Tahoma"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextFontCharSet =0
                    TextAlign =3
                    TextFontFamily =34
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1620
                    Width =1524
                    Height =450
                    FontSize =8
                    TabIndex =4
                    BorderColor =32768
                    Name ="MABL"
                    ControlSource ="MABL"
                    Format ="#,###"
                    FontName ="Tahoma"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextFontCharSet =0
                    TextAlign =3
                    TextFontFamily =34
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =60
                    Top =16
                    Width =1464
                    Height =405
                    FontSize =8
                    TabIndex =5
                    BorderColor =32768
                    Name ="MABL_K"
                    ControlSource ="=[MABL]*[MEGH_MAR]"
                    Format ="#,###"
                    FontName ="Tahoma"
                End
                Begin TextBox
                    TextFontCharSet =0
                    TextAlign =2
                    TextFontFamily =34
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3195
                    Top =15
                    Width =1575
                    Height =405
                    FontSize =8
                    TabIndex =6
                    BorderColor =32768
                    Name ="VAHED_K"
                    ControlSource ="VNAMES"
                    FontName ="Tahoma"
                End
                Begin Line
                    BorderWidth =1
                    Left =15
                    Top =450
                    Width =15180
                    BorderColor =0
                    Name ="Line100"
                End
                Begin Line
                    BorderWidth =3
                    Left =15210
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line101"
                End
                Begin Line
                    BorderWidth =1
                    Left =14520
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line102"
                End
                Begin Line
                    BorderWidth =1
                    Left =13095
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line103"
                End
                Begin Line
                    BorderWidth =1
                    Left =6120
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line104"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Left =4815
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line105"
                End
                Begin Line
                    BorderWidth =1
                    Left =3180
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line106"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Left =1590
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line108"
                End
                Begin Line
                    BorderWidth =2
                    Top =15
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line270"
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =3344
            OnFormat ="[Event Procedure]"
            Name ="GroupFooter3"
            Begin
                Begin TextBox
                    DecimalPlaces =0
                    TextFontCharSet =0
                    TextAlign =3
                    TextFontFamily =34
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Width =1539
                    Height =405
                    FontSize =8
                    BorderColor =32768
                    Name ="Text271"
                    ControlSource ="=[Forms]![HEAD_LST_FROOSH_BACK2]![JF]"
                    Format ="#,###"
                    FontName ="Tahoma"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextFontCharSet =0
                    TextAlign =3
                    TextFontFamily =34
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3
                    Top =453
                    Width =1539
                    Height =405
                    FontSize =8
                    TabIndex =1
                    BorderColor =32768
                    Name ="TAKHFIF"
                    ControlSource ="=[Forms]![HEAD_LST_FROOSH_BACK2]![NTKHFIF]"
                    Format ="#,###"
                    FontName ="Tahoma"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextFontCharSet =0
                    TextAlign =3
                    TextFontFamily =34
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =-5
                    Top =907
                    Width =1554
                    Height =405
                    FontSize =8
                    TabIndex =2
                    BorderColor =32768
                    Name ="MBAA"
                    ControlSource ="=[Forms]![HEAD_LST_FROOSH_BACK2]![mbaa]"
                    Format ="#,###"
                    FontName ="Tahoma"
                End
                Begin Line
                    BorderWidth =1
                    Left =1590
                    Width =0
                    Height =1845
                    BorderColor =0
                    Name ="Line278"
                End
                Begin Label
                    OldBorderStyle =1
                    BorderWidth =1
                    TextFontCharSet =0
                    TextAlign =3
                    TextFontFamily =34
                    Left =2895
                    Top =1095
                    Width =4485
                    Height =390
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="Label283"
                    Caption ="مهر و امضاء خريدار :"
                    FontName ="Tahoma"
                End
                Begin Label
                    OldBorderStyle =1
                    BorderWidth =1
                    TextFontCharSet =0
                    TextAlign =3
                    TextFontFamily =34
                    Left =7380
                    Top =1095
                    Width =7830
                    Height =390
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="Label284"
                    Caption ="مهر و امضاء فروشنده :"
                    FontName ="Tahoma"
                End
                Begin TextBox
                    TextFontCharSet =0
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =7365
                    Top =795
                    Width =7086
                    Height =285
                    FontSize =8
                    TabIndex =3
                    Name ="MOLAH"
                    ControlSource ="MOLAH"
                    FontName ="Tahoma"
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            BorderWidth =1
                            TextFontCharSet =0
                            TextAlign =3
                            TextFontFamily =34
                            Left =7367
                            Top =799
                            Width =7830
                            Height =300
                            FontSize =8
                            FontWeight =400
                            ForeColor =0
                            Name ="Label285"
                            Caption ="توضيحات:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin Label
                    OldBorderStyle =1
                    BorderWidth =1
                    TextFontCharSet =0
                    TextAlign =3
                    TextFontFamily =34
                    Left =7380
                    Top =495
                    Width =7830
                    Height =285
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="Label286"
                    Caption ="شرايط و نحوه فروش:        نقدي                                   غير نقدي  "
                    FontName ="Tahoma"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextFontCharSet =0
                    TextAlign =3
                    TextFontFamily =34
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Top =1365
                    Width =1554
                    Height =405
                    FontSize =8
                    TabIndex =4
                    BorderColor =32768
                    Name ="Text287"
                    ControlSource ="=Sum([MABL_K])-[TAKHFIF]+[MBAA]"
                    Format ="#,###"
                    FontName ="Tahoma"
                End
                Begin Line
                    BorderWidth =1
                    Width =0
                    Height =1845
                    BorderColor =0
                    Name ="Line288"
                End
                Begin Line
                    BorderWidth =1
                    Top =1830
                    Width =2724
                    BorderColor =0
                    Name ="Line289"
                End
                Begin Label
                    BorderWidth =1
                    TextFontCharSet =0
                    TextFontFamily =34
                    Left =1650
                    Top =105
                    Width =780
                    Height =240
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="Label290"
                    Caption ="جمع مبلغ:"
                    FontName ="Tahoma"
                End
                Begin Label
                    BorderWidth =1
                    TextFontCharSet =0
                    TextFontFamily =34
                    Left =1650
                    Top =540
                    Width =780
                    Height =240
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="Label291"
                    Caption ="تخفيف:"
                    FontName ="Tahoma"
                End
                Begin Label
                    BorderWidth =1
                    TextFontCharSet =0
                    TextFontFamily =34
                    Left =1650
                    Top =1005
                    Width =960
                    Height =240
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="Label292"
                    Caption ="ارزش افزوده:"
                    FontName ="Tahoma"
                End
                Begin Label
                    BorderWidth =1
                    TextFontCharSet =0
                    TextFontFamily =34
                    Left =1650
                    Top =1485
                    Width =960
                    Height =240
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="Label293"
                    Caption ="قابل پرداخت:"
                    FontName ="Tahoma"
                End
                Begin Line
                    BorderWidth =1
                    Left =2715
                    Width =0
                    Height =1845
                    BorderColor =0
                    Name ="Line294"
                End
                Begin Line
                    BorderWidth =1
                    Top =1335
                    Width =2724
                    BorderColor =0
                    Name ="Line295"
                End
                Begin Line
                    BorderWidth =1
                    Top =885
                    Width =2724
                    BorderColor =0
                    Name ="Line296"
                End
                Begin Line
                    BorderWidth =1
                    Top =435
                    Width =2724
                    BorderColor =0
                    Name ="Line297"
                End
                Begin BoundObjectFrame
                    OldBorderStyle =0
                    Left =12595
                    Top =1540
                    Width =1701
                    Height =1701
                    TabIndex =5
                    Name ="EMZA1"
                    ControlSource ="EMZA1"
                End
                Begin BoundObjectFrame
                    OldBorderStyle =0
                    Left =9938
                    Top =1537
                    Width =1701
                    Height =1701
                    TabIndex =6
                    Name ="EMZA2"
                    ControlSource ="EMZA2"
                End
                Begin BoundObjectFrame
                    OldBorderStyle =0
                    Left =7313
                    Top =1537
                    Width =1701
                    Height =1701
                    TabIndex =7
                    Name ="EMZA3"
                    ControlSource ="EMZA3"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =13430
                    Top =1863
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
                    Left =10838
                    Top =1870
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
                    Left =8183
                    Top =1871
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
                    Left =13427
                    Top =1530
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
                    Left =10840
                    Top =1537
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
                    Left =8168
                    Top =1538
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
                    Left =3741
                    Top =1757
                    TabIndex =8
                    Name ="SGN1"
                    ControlSource ="SGN1"
                End
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Left =3741
                    Top =2076
                    TabIndex =9
                    Name ="SGN2"
                    ControlSource ="SGN2"
                End
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Left =3741
                    Top =2395
                    TabIndex =10
                    Name ="SGN3"
                    ControlSource ="SGN3"
                End
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Left =453
                    Top =2437
                    TabIndex =11
                    Name ="SGN1usid"
                    ControlSource ="SGN1usid"
                End
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Left =453
                    Top =2557
                    TabIndex =12
                    Name ="sgn2usid"
                    ControlSource ="sgn2usid"
                End
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Left =453
                    Top =2782
                    TabIndex =13
                    Name ="sgn3usid"
                    ControlSource ="sgn3usid"
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
' See "invoice_frbk_2_mba.cls"
