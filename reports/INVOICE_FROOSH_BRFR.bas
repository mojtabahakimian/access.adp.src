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
    Width =10601
    DatasheetFontHeight =10
    ItemSuffix =160
    Left =600
    Top =2100
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xf4e752ed14f9e540
    End
    RecordSource ="SELECT     dbo.HEAD_LST.NUMBER1, dbo.HEAD_LST.NUMBER, dbo.HEAD_LST.DATE_N, dbo.I"
        "NVO_LST.NUMBER AS INUMBER, \015\012                      dbo.HEAD_LST.TAG - 1 AS"
        " HTAG, dbo.INVO_LST.ANBAR, dbo.INVO_LST.RADIF, dbo.INVO_LST.CODE, dbo.INVO_LST.M"
        "EGH, dbo.INVO_LST.MEGHk, \015\012                      dbo.INVO_LST.MEGH_MAR, db"
        "o.INVO_LST.MANDAH, dbo.INVO_LST.MABL, dbo.INVO_LST.MABL_K, dbo.INVO_LST.FROM_A, "
        "dbo.INVO_LST.N_RASID, \015\012                      dbo.INVO_LST.MEGH_R, dbo.INV"
        "O_LST.RADAH, dbo.INVO_LST.SANAD_NO, dbo.INVO_LST.ANBARF, dbo.INVO_LST.VAHED_K, d"
        "bo.STUF_DEF.NAME, \015\012                      dbo.TCOD_ANBAR.NAMES, dbo.TCOD_V"
        "AHEDS.NAMES AS VNAMES, dbo.HEAD_LST.TAH, dbo.HEAD_LST.MOLAH, dbo.CUSTKIND.CUSTKN"
        "AME, \015\012                      dbo.DEPART.DEPNAME, dbo.SHIFT.SHNAME, dbo.CUS"
        "T_HESAB.NAME AS HESAB, dbo.CUST_HESAB.ADDRESS, dbo.CUST_HESAB.TEL, \015\012     "
        "                 ISNULL(dbo.STUF_DEF.NAME, N' ') + N' ' + ISNULL(dbo.INVO_LST.MA"
        "NDAH, N' ') AS KALA, dbo.HEAD_LST.CUST_NO, dbo.INVO_LST.N_KOL, \015\012         "
        "             dbo.INVO_LST.N_MOIN, dbo.HEAD_LST.FNUMCO, dbo.CUST_HESAB.ECODE, dbo"
        ".CUST_HESAB.PCODE, dbo.CUST_HESAB.IYALAT, \015\012                      dbo.CUST"
        "_HESAB.MCODEM, dbo.CUST_HESAB.CITY, dbo.INVO_LST.MABL_K - dbo.INVO_LST.N_MOIN AS"
        " mabkbt, dbo.INVO_LST.IMBAA, \015\012                      dbo.INVO_LST.MABL_K -"
        " dbo.INVO_LST.N_MOIN + dbo.INVO_LST.IMBAA AS mabkn, dbo.CUST_HESAB.CODE_E, dbo.H"
        "EAD_LST.TAKHFIF, \015\012                      dbo.HEAD_LST.MBAA, dbo.STUF_DEF.N"
        "_FANI, dbo.HEAD_LST.SADER, dbo.HEAD_LST.ANBARF AS ANBARFF, dbo.HEAD_LST.SGN1, \015"
        "\012                      dbo.HEAD_LST.SGN2, dbo.HEAD_LST.SGN3, dbo.HEAD_LST.sgn"
        "1usid, dbo.HEAD_LST.sgn2usid, dbo.HEAD_LST.sgn3usid, \015\012                   "
        "   dbo.SALA_DTL.EMZA AS EMZA1, SALA_DTL_1.EMZA AS EMZA2, SALA_DTL_2.EMZA AS EMZA"
        "3\015\012FROM         dbo.STUF_DEF RIGHT OUTER JOIN\015\012                     "
        " dbo.CUSTKIND RIGHT OUTER JOIN\015\012                      dbo.INVO_LST INNER J"
        "OIN\015\012                      dbo.TCOD_ANBAR ON dbo.INVO_LST.ANBAR = dbo.TCOD"
        "_ANBAR.CODE INNER JOIN\015\012                      dbo.HEAD_LST ON dbo.INVO_LST"
        ".NUMBER = dbo.HEAD_LST.NUMBER AND dbo.INVO_LST.TAG = dbo.HEAD_LST.TAG - 1 LEFT O"
        "UTER JOIN\015\012                      dbo.SALA_DTL SALA_DTL_2 ON dbo.HEAD_LST.s"
        "gn3usid = SALA_DTL_2.IDD LEFT OUTER JOIN\015\012                      dbo.SALA_D"
        "TL SALA_DTL_1 ON dbo.HEAD_LST.sgn2usid = SALA_DTL_1.IDD LEFT OUTER JOIN\015\012 "
        "                     dbo.SALA_DTL ON dbo.HEAD_LST.sgn1usid = dbo.SALA_DTL.IDD LE"
        "FT OUTER JOIN\015\012                      dbo.CUST_HESAB ON dbo.HEAD_LST.CUST_N"
        "O = dbo.CUST_HESAB.hes LEFT OUTER JOIN\015\012                      dbo.SHIFT ON"
        " dbo.HEAD_LST.SHIFT = dbo.SHIFT.SHIFT_ID LEFT OUTER JOIN\015\012                "
        "      dbo.DEPART ON dbo.HEAD_LST.DEPATMAN = dbo.DEPART.DEPATMAN ON dbo.CUSTKIND."
        "CUST_COD = dbo.HEAD_LST.CUST_KIND LEFT OUTER JOIN\015\012                      d"
        "bo.TCOD_VAHEDS ON dbo.INVO_LST.VAHED_K = dbo.TCOD_VAHEDS.CODE ON dbo.STUF_DEF.CO"
        "DE = dbo.INVO_LST.CODE"
    Caption ="فاكتور برگشت فروش"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial (Arabic)"
    OnError ="[Event Procedure]"
    RecordSource ="SELECT     dbo.HEAD_LST.NUMBER1, dbo.HEAD_LST.NUMBER, dbo.HEAD_LST.DATE_N, dbo.I"
        "NVO_LST.NUMBER AS INUMBER, \015\012                      dbo.HEAD_LST.TAG - 1 AS"
        " HTAG, dbo.INVO_LST.ANBAR, dbo.INVO_LST.RADIF, dbo.INVO_LST.CODE, dbo.INVO_LST.M"
        "EGH, dbo.INVO_LST.MEGHk, \015\012                      dbo.INVO_LST.MEGH_MAR, db"
        "o.INVO_LST.MANDAH, dbo.INVO_LST.MABL, dbo.INVO_LST.MABL_K, dbo.INVO_LST.FROM_A, "
        "dbo.INVO_LST.N_RASID, \015\012                      dbo.INVO_LST.MEGH_R, dbo.INV"
        "O_LST.RADAH, dbo.INVO_LST.SANAD_NO, dbo.INVO_LST.ANBARF, dbo.INVO_LST.VAHED_K, d"
        "bo.STUF_DEF.NAME, \015\012                      dbo.TCOD_ANBAR.NAMES, dbo.TCOD_V"
        "AHEDS.NAMES AS VNAMES, dbo.HEAD_LST.TAH, dbo.HEAD_LST.MOLAH, dbo.CUSTKIND.CUSTKN"
        "AME, \015\012                      dbo.DEPART.DEPNAME, dbo.SHIFT.SHNAME, dbo.CUS"
        "T_HESAB.NAME AS HESAB, dbo.CUST_HESAB.ADDRESS, dbo.CUST_HESAB.TEL, \015\012     "
        "                 ISNULL(dbo.STUF_DEF.NAME, N' ') + N' ' + ISNULL(dbo.INVO_LST.MA"
        "NDAH, N' ') AS KALA, dbo.HEAD_LST.CUST_NO, dbo.INVO_LST.N_KOL, \015\012         "
        "             dbo.INVO_LST.N_MOIN, dbo.HEAD_LST.FNUMCO, dbo.CUST_HESAB.ECODE, dbo"
        ".CUST_HESAB.PCODE, dbo.CUST_HESAB.IYALAT, \015\012                      dbo.CUST"
        "_HESAB.MCODEM, dbo.CUST_HESAB.CITY, dbo.INVO_LST.MABL_K - dbo.INVO_LST.N_MOIN AS"
        " mabkbt, dbo.INVO_LST.IMBAA, \015\012                      dbo.INVO_LST.MABL_K -"
        " dbo.INVO_LST.N_MOIN + dbo.INVO_LST.IMBAA AS mabkn, dbo.CUST_HESAB.CODE_E, dbo.H"
        "EAD_LST.TAKHFIF, \015\012                      dbo.HEAD_LST.MBAA, dbo.STUF_DEF.N"
        "_FANI, dbo.HEAD_LST.SADER, dbo.HEAD_LST.ANBARF AS ANBARFF, dbo.HEAD_LST.SGN1, \015"
        "\012                      dbo.HEAD_LST.SGN2, dbo.HEAD_LST.SGN3, dbo.HEAD_LST.sgn"
        "1usid, dbo.HEAD_LST.sgn2usid, dbo.HEAD_LST.sgn3usid, \015\012                   "
        "   dbo.SALA_DTL.EMZA AS EMZA1, SALA_DTL_1.EMZA AS EMZA2, SALA_DTL_2.EMZA AS EMZA"
        "3\015\012FROM         dbo.STUF_DEF RIGHT OUTER JOIN\015\012                     "
        " dbo.CUSTKIND RIGHT OUTER JOIN\015\012                      dbo.INVO_LST INNER J"
        "OIN\015\012                      dbo.TCOD_ANBAR ON dbo.INVO_LST.ANBAR = dbo.TCOD"
        "_ANBAR.CODE INNER JOIN\015\012                      dbo.HEAD_LST ON dbo.INVO_LST"
        ".NUMBER = dbo.HEAD_LST.NUMBER AND dbo.INVO_LST.TAG = dbo.HEAD_LST.TAG - 1 LEFT O"
        "UTER JOIN\015\012                      dbo.SALA_DTL SALA_DTL_2 ON dbo.HEAD_LST.s"
        "gn3usid = SALA_DTL_2.IDD LEFT OUTER JOIN\015\012                      dbo.SALA_D"
        "TL SALA_DTL_1 ON dbo.HEAD_LST.sgn2usid = SALA_DTL_1.IDD LEFT OUTER JOIN\015\012 "
        "                     dbo.SALA_DTL ON dbo.HEAD_LST.sgn1usid = dbo.SALA_DTL.IDD LE"
        "FT OUTER JOIN\015\012                      dbo.CUST_HESAB ON dbo.HEAD_LST.CUST_N"
        "O = dbo.CUST_HESAB.hes LEFT OUTER JOIN\015\012                      dbo.SHIFT ON"
        " dbo.HEAD_LST.SHIFT = dbo.SHIFT.SHIFT_ID LEFT OUTER JOIN\015\012                "
        "      dbo.DEPART ON dbo.HEAD_LST.DEPATMAN = dbo.DEPART.DEPATMAN ON dbo.CUSTKIND."
        "CUST_COD = dbo.HEAD_LST.CUST_KIND LEFT OUTER JOIN\015\012                      d"
        "bo.TCOD_VAHEDS ON dbo.INVO_LST.VAHED_K = dbo.TCOD_VAHEDS.CODE ON dbo.STUF_DEF.CO"
        "DE = dbo.INVO_LST.CODE"
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
            ControlSource ="[NUMBER]"
        End
        Begin BreakLevel
            ControlSource ="RADIF"
        End
        Begin PageHeader
            Height =3075
            Name ="PageHeader"
            Begin
                Begin Label
                    TextFontFamily =2
                    Left =3960
                    Width =2160
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
                    Top =480
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
                    Left =105
                    Top =540
                    Width =1266
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
                            Left =1408
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
                    Left =7800
                    Top =510
                    Width =1281
                    Height =435
                    FontSize =12
                    TabIndex =2
                    BorderColor =32768
                    Name ="NUMBER1"
                    ControlSource ="NUMBER1"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =9132
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
                    Top =1050
                    Width =10506
                    BorderColor =0
                    Name ="Line94"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1785
                    Top =1062
                    Width =7671
                    Height =450
                    FontSize =12
                    TabIndex =3
                    BorderColor =32768
                    Name ="HESAB"
                    ControlSource ="HESAB"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =9435
                            Top =1065
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
                    Left =75
                    Top =1050
                    Width =1131
                    Height =450
                    FontSize =12
                    TabIndex =4
                    BorderColor =32768
                    Name ="hes"
                    ControlSource ="CUST_NO"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =1197
                            Top =1050
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
                    Left =1950
                    Top =1539
                    Width =7656
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
                            Left =9642
                            Top =1539
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
                    Left =75
                    Top =1530
                    Width =1161
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
                            Left =1257
                            Top =1530
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
                    Top =2550
                    Width =10506
                    BorderColor =0
                    Name ="Line99"
                End
                Begin Line
                    BorderWidth =3
                    Left =10485
                    Top =2565
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line130"
                End
                Begin Line
                    Left =9915
                    Top =2565
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line131"
                End
                Begin Line
                    Left =9345
                    Top =2565
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line132"
                End
                Begin Line
                    Left =5535
                    Top =2565
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line133"
                End
                Begin Line
                    Left =4710
                    Top =2565
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line134"
                End
                Begin Line
                    Left =3570
                    Top =2565
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line135"
                End
                Begin Line
                    Left =2625
                    Top =2565
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line136"
                End
                Begin Line
                    Left =1365
                    Top =2565
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line137"
                End
                Begin Line
                    BorderWidth =2
                    Top =2565
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line138"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =390
                    Top =2610
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
                    Left =1599
                    Top =2606
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
                    Left =2715
                    Top =2610
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
                    Left =4882
                    Top =2610
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
                    Left =3747
                    Top =2610
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
                    Left =6637
                    Top =2610
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
                    Left =9420
                    Top =2610
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
                    Left =9921
                    Top =2606
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
                    Top =3075
                    Width =10506
                    BorderColor =0
                    Name ="Line147"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =90
                    Top =2055
                    Width =9306
                    Height =450
                    FontSize =12
                    TabIndex =7
                    BorderColor =32768
                    Name ="MOLAH"
                    ControlSource ="MOLAH"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =9420
                            Top =2055
                            Width =990
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
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =339
                    Top =170
                    TabIndex =8
                    Name ="HTAG"
                    ControlSource ="HTAG"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =6870
                    Width =1131
                    Height =435
                    FontSize =12
                    TabIndex =9
                    BorderColor =32768
                    Name ="Text155"
                    ControlSource ="FNUMCO"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =8055
                            Width =2370
                            Height =480
                            FontSize =12
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label156"
                            Caption ="شماره فاكتور فروشنده:"
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
                    Left =90
                    Width =1281
                    Height =435
                    FontSize =12
                    TabIndex =10
                    BorderColor =32768
                    Name ="Text157"
                    ControlSource ="NUMBER"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =1422
                            Width =1275
                            Height =480
                            FontSize =12
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label158"
                            Caption ="شماره رسيد:"
                            FontName ="WeblogmaYekan"
                        End
                    End
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
                    Left =9377
                    Top =16
                    Width =504
                    Height =405
                    FontSize =10
                    BorderColor =32768
                    Name ="IANBAR"
                    ControlSource ="ANBAR"
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
                    Left =9931
                    Top =16
                    Width =504
                    Height =405
                    FontSize =10
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
                    Left =5547
                    Top =16
                    Width =3774
                    Height =405
                    FontSize =10
                    TabIndex =2
                    BorderColor =32768
                    Name ="CODE0"
                    ControlSource ="=IIf([Forms]![baseknow]![codeview]=-1,[kala] & \" \" & [code],[kala])"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4735
                    Top =16
                    Width =789
                    Height =405
                    FontSize =10
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
                    Left =2640
                    Top =16
                    Width =894
                    Height =405
                    FontSize =10
                    TabIndex =4
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
                    Left =1399
                    Top =16
                    Width =1194
                    Height =405
                    FontSize =10
                    TabIndex =5
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
                    Width =1269
                    Height =405
                    FontSize =10
                    TabIndex =6
                    BorderColor =32768
                    Name ="MABL_K"
                    ControlSource ="MABL_K"
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3585
                    Top =15
                    Width =1110
                    Height =405
                    FontSize =10
                    TabIndex =7
                    BorderColor =32768
                    Name ="VNAMES"
                    ControlSource ="VNAMES"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Top =450
                    Width =10491
                    BorderColor =0
                    Name ="Line100"
                End
                Begin Line
                    BorderWidth =3
                    Left =10485
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line101"
                End
                Begin Line
                    BorderWidth =1
                    Left =9915
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line102"
                End
                Begin Line
                    BorderWidth =1
                    Left =9345
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line103"
                End
                Begin Line
                    BorderWidth =1
                    Left =5535
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line104"
                End
                Begin Line
                    BorderWidth =1
                    Left =4710
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line105"
                End
                Begin Line
                    BorderWidth =1
                    Left =3570
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line106"
                End
                Begin Line
                    BorderWidth =1
                    Left =2625
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line107"
                End
                Begin Line
                    BorderWidth =1
                    Left =1365
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
                    Visible = NotDefault
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =1020
                    Top =56
                    TabIndex =8
                    Name ="CODE"
                    ControlSource ="CODE"
                    FontName ="WeblogmaYekan"
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =5272
            OnFormat ="[Event Procedure]"
            Name ="GroupFooter3"
            Begin
                Begin Line
                    BorderWidth =3
                    Left =60
                    Top =30
                    Width =10405
                    BorderColor =0
                    Name ="Line110"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =2
                    Left =3450
                    Top =30
                    Width =0
                    Height =3196
                    BorderColor =0
                    Name ="Line118"
                End
                Begin Line
                    BorderWidth =2
                    Width =0
                    Height =3226
                    BorderColor =0
                    Name ="Line119"
                End
                Begin Line
                    BorderWidth =3
                    Left =30
                    Top =3240
                    Width =10450
                    BorderColor =0
                    Name ="Line120"
                End
                Begin Line
                    Left =30
                    Top =465
                    Width =3435
                    BorderColor =0
                    Name ="Line121"
                End
                Begin Line
                    Left =30
                    Top =885
                    Width =3435
                    BorderColor =0
                    Name ="Line122"
                End
                Begin Line
                    Left =30
                    Top =1305
                    Width =3435
                    BorderColor =0
                    Name ="Line123"
                End
                Begin Line
                    Left =30
                    Top =2370
                    Width =3435
                    BorderColor =0
                    Name ="Line124"
                End
                Begin Line
                    Left =30
                    Top =2760
                    Width =3435
                    BorderColor =0
                    Name ="Line125"
                End
                Begin Line
                    Left =1755
                    Top =30
                    Width =0
                    Height =3196
                    BorderColor =0
                    Name ="Line126"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =3
                    Left =10485
                    Width =0
                    Height =3271
                    BorderColor =0
                    Name ="Line127"
                End
                Begin Line
                    Left =3315
                    Top =885
                    Width =7143
                    BorderColor =0
                    Name ="Line128"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =2
                    Left =9570
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
                    Left =3525
                    Top =945
                    Width =6030
                    Height =1710
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
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3571
                    Top =2664
                    Width =5871
                    Height =435
                    FontSize =14
                    BackColor =128
                    Name ="MANDAH"
                    ControlSource ="=[Forms]![HEAD_LST_BRFR]![MANDAH]"
                    Format ="#,##0.00;#,##0.00-"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextAlign =0
                            TextFontFamily =2
                            Left =9451
                            Top =2664
                            Width =1020
                            Height =435
                            FontSize =12
                            FontWeight =400
                            BackColor =128
                            ForeColor =0
                            Name ="Label69"
                            Caption ="مانده حساب:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin Label
                    TextAlign =0
                    TextFontFamily =2
                    Left =1788
                    Top =2347
                    Width =1665
                    Height =435
                    FontSize =11
                    FontWeight =400
                    ForeColor =0
                    Name ="Label177"
                    Caption ="جمع مبالغ پرداختي:"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =0
                    TextFontFamily =2
                    Left =1793
                    Top =75
                    Width =1395
                    Height =435
                    FontSize =11
                    FontWeight =400
                    ForeColor =0
                    Name ="Label178"
                    Caption ="جمع كل فاكتور:"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =0
                    TextFontFamily =2
                    Left =1787
                    Top =480
                    Width =1320
                    Height =435
                    FontSize =11
                    FontWeight =400
                    ForeColor =0
                    Name ="Label179"
                    Caption ="هزينه خدمات:"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =0
                    TextFontFamily =2
                    Left =1787
                    Top =915
                    Width =1335
                    Height =435
                    FontSize =11
                    FontWeight =400
                    ForeColor =0
                    Name ="Label180"
                    Caption ="تخفيفات:"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =0
                    TextFontFamily =2
                    Left =1787
                    Top =1905
                    Width =1740
                    Height =435
                    FontSize =11
                    FontWeight =400
                    ForeColor =0
                    Name ="Label181"
                    Caption ="مبلغ قابل پرداخت:"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =0
                    TextFontFamily =2
                    Left =1787
                    Top =2745
                    Width =1170
                    Height =435
                    FontSize =11
                    FontWeight =400
                    ForeColor =0
                    Name ="Label182"
                    Caption ="مانده فاكتور:"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    NumeralShapes =2
                    Left =1
                    Top =2347
                    Width =1726
                    Height =435
                    FontSize =11
                    FontWeight =400
                    ForeColor =0
                    Name ="JPAY"
                    Caption ="جمع مبالغ پرداختي:"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    NumeralShapes =2
                    Left =66
                    Top =75
                    Width =1666
                    Height =435
                    FontSize =11
                    FontWeight =400
                    ForeColor =0
                    Name ="JF"
                    Caption ="جمع كل فاكتور:"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    NumeralShapes =2
                    Left =60
                    Top =480
                    Width =1666
                    Height =435
                    FontSize =11
                    FontWeight =400
                    ForeColor =0
                    Name ="HKH"
                    Caption ="هزينه خدمات:"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    NumeralShapes =2
                    Left =60
                    Top =915
                    Width =1666
                    Height =435
                    FontSize =11
                    FontWeight =400
                    ForeColor =0
                    Name ="TF"
                    Caption ="تخفيفات:"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    NumeralShapes =2
                    Left =60
                    Top =1905
                    Width =1666
                    Height =435
                    FontSize =11
                    FontWeight =400
                    ForeColor =0
                    Name ="GABEL"
                    Caption ="مبلغ قابل پرداخت:"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    NumeralShapes =2
                    Left =60
                    Top =2745
                    Width =1666
                    Height =435
                    FontSize =11
                    FontWeight =400
                    ForeColor =0
                    Name ="MAN"
                    Caption ="مانده فاكتور:"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =2
                    Left =9203
                    Top =56
                    Width =1200
                    Height =405
                    FontSize =10
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label189"
                    Caption ="مبلغ به حروف:"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =2
                    Left =3375
                    Top =56
                    Width =5790
                    Height =795
                    FontSize =10
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="HR"
                    Caption ="مبلغ به حروف:"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    BorderWidth =1
                    TextAlign =0
                    TextFontFamily =2
                    Left =1785
                    Top =1365
                    Width =1620
                    Height =435
                    FontSize =10
                    FontWeight =400
                    BorderColor =4210752
                    ForeColor =0
                    Name ="Label202"
                    Caption ="3% ماليات و عوارض:"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    BorderWidth =1
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    NumeralShapes =2
                    Left =61
                    Top =1395
                    Width =1635
                    Height =435
                    FontSize =11
                    FontWeight =400
                    BorderColor =4210752
                    ForeColor =0
                    Name ="MBAA"
                    Caption ="ماليات"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    Left =30
                    Top =1845
                    Width =3435
                    BorderColor =0
                    Name ="Line159"
                End
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Top =3288
                    TabIndex =1
                    Name ="SGN1"
                    ControlSource ="SGN1"
                End
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Top =3607
                    TabIndex =2
                    Name ="SGN2"
                    ControlSource ="SGN2"
                End
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Top =3926
                    TabIndex =3
                    Name ="SGN3"
                    ControlSource ="SGN3"
                End
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Top =4081
                    TabIndex =4
                    Name ="SGN1usid"
                    ControlSource ="SGN1usid"
                End
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Top =4201
                    TabIndex =5
                    Name ="sgn2usid"
                    ControlSource ="sgn2usid"
                End
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Top =4426
                    TabIndex =6
                    Name ="sgn3usid"
                    ControlSource ="sgn3usid"
                End
                Begin BoundObjectFrame
                    OldBorderStyle =0
                    Left =7922
                    Top =3415
                    Width =1701
                    Height =1701
                    TabIndex =7
                    Name ="EMZA1"
                    ControlSource ="EMZA1"
                End
                Begin BoundObjectFrame
                    OldBorderStyle =0
                    Left =5265
                    Top =3412
                    Width =1701
                    Height =1701
                    TabIndex =8
                    Name ="EMZA2"
                    ControlSource ="EMZA2"
                End
                Begin BoundObjectFrame
                    OldBorderStyle =0
                    Left =2640
                    Top =3412
                    Width =1701
                    Height =1701
                    TabIndex =9
                    Name ="EMZA3"
                    ControlSource ="EMZA3"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =8757
                    Top =3738
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
                    Left =6165
                    Top =3745
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
                    Left =3510
                    Top =3746
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
                    Left =8754
                    Top =3405
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
                    Left =6167
                    Top =3412
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
                    Left =3495
                    Top =3413
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
        Begin PageFooter
            Visible = NotDefault
            Height =0
            Name ="PageFooter"
        End
    End
End
CodeBehindForm
' See "INVOICE_FROOSH_BRFR.cls"
