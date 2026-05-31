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
    GridX =64
    GridY =64
    Width =10335
    DatasheetFontHeight =10
    ItemSuffix =228
    Left =1560
    Top =690
    DatasheetGridlinesColor =12632256
    Tag ="QSL_HAVLAH_ANBAR"
    RecSrcDt = Begin
        0x844b510a33ede540
    End
    RecordSource ="SELECT HEAD_BACK_ANBAR.NUMBER1, HEAD_BACK_ANBAR.NUMBER, HEAD_BACK_ANBAR.DATE_N, "
        "INVO_LST.NUMBER AS INUMBER, HEAD_BACK_ANBAR.HTAG, INVO_LST.ANBAR, INVO_LST.RADIF"
        ", INVO_LST.CODE, INVO_LST.MEGH, INVO_LST.MEGHk, INVO_LST.MEGH_MAR, INVO_LST.MAND"
        "AH, INVO_LST.MABL, INVO_LST.MABL_K, INVO_LST.FROM_A, INVO_LST.N_RASID, INVO_LST."
        "MEGH_R, INVO_LST.RADAH, INVO_LST.SANAD_NO, INVO_LST.ANBARF, INVO_LST.VAHED_K, ST"
        "UF_DEF.NAME, TCOD_ANBAR.NAMES, TCOD_VAHEDS.NAMES AS VNAMES, HEAD_BACK_ANBAR.TAH,"
        " HEAD_BACK_ANBAR.MOLAH, CUSTKIND.CUSTKNAME, DEPART.DEPNAME, SHIFT.SHNAME, CUST_H"
        "ESAB.NAME AS HESAB, CUST_HESAB.ADDRESS, CUST_HESAB.TEL, ISNULL(STUF_DEF.NAME, N'"
        " ') + N' ' + ISNULL(INVO_LST.MANDAH, N' ') AS KALA, HEAD_BACK_ANBAR.CUST_NO, INV"
        "O_LST.N_KOL, INVO_LST.N_MOIN, HEAD_BACK_ANBAR.FNUMCO, CUST_HESAB.ECODE, CUST_HES"
        "AB.PCODE, CUST_HESAB.IYALAT, CUST_HESAB.MCODEM, CUST_HESAB.CITY, INVO_LST.MABL_K"
        " - INVO_LST.N_MOIN AS mabkbt, INVO_LST.IMBAA, INVO_LST.MABL_K - INVO_LST.N_MOIN "
        "+ INVO_LST.IMBAA AS mabkn, CUST_HESAB.CODE_E, HEAD_BACK_ANBAR.TAKHFIF, HEAD_BACK"
        "_ANBAR.MBAA, STUF_DEF.N_FANI, HEAD_BACK_ANBAR.SADER, HEAD_BACK_ANBAR.ANBARF AS A"
        "NBARFF, OTHER_DTL.REQUEST_NO, OTHER_DTL.BARNAMEH, OTHER_DTL.DRIVER, OTHER_DTL.DR"
        "IVER_MOB, OTHER_DTL.CAMIUN_NUM, OTHER_DTL.MAGHSAD, OTHER_DTL.CAM_KHALY, OTHER_DT"
        "L.CAM_POOR, OTHER_DTL.TOZIH, OTHER_DTL.CAMIUN, OTHER_DTL_SUB.CODE AS CODEG, OTHE"
        "R_DTL_SUB.CAM_KHALY AS CAM_KHALYG, OTHER_DTL_SUB.CAM_POOR AS CAM_POORG, OTHER_DT"
        "L_SUB.MEGHk AS MEGHkG, OTHER_DTL_SUB.TOZIH AS TOZIHG, OTHER_DTL_SUB.VAZNH, STUF_"
        "DEF_1.NAME AS NAMEG, HEAD_BACK_ANBAR.SHARAYET, DEPART.DEPART, INVO_LST.TKHN, HEA"
        "D_BACK_ANBAR.MAS, HEAD_BACK_ANBAR.HTAG AS TAG, SALA_DTL.EMZA AS EMZA1, SALA_DTL_"
        "1.EMZA AS EMZA2, SALA_DTL_2.EMZA AS EMZA3, HEAD_BACK_ANBAR.SGN1usid, HEAD_BACK_A"
        "NBAR.sgn2usid, HEAD_BACK_ANBAR.sgn3usid, HEAD_LST.SGN1, HEAD_LST.SGN2, HEAD_LST."
        "SGN3, HEAD_LST.SGN4 FROM STUF_DEF RIGHT OUTER JOIN SHIFT RIGHT OUTER JOIN DEPART"
        " RIGHT OUTER JOIN HEAD_LST INNER JOIN INVO_LST INNER JOIN TCOD_ANBAR ON INVO_LST"
        ".ANBAR = TCOD_ANBAR.CODE INNER JOIN HEAD_BACK_ANBAR ON INVO_LST.NUMBER = HEAD_BA"
        "CK_ANBAR.NUMBER AND INVO_LST.TAG = HEAD_BACK_ANBAR.HTAG ON HEAD_LST.NUMBER = HEA"
        "D_BACK_ANBAR.NUMBER AND HEAD_LST.TAG = HEAD_BACK_ANBAR.TAG ON DEPART.DEPATMAN = "
        "HEAD_BACK_ANBAR.DEPATMAN LEFT OUTER JOIN CUSTKIND ON HEAD_BACK_ANBAR.CUST_KIND ="
        " CUSTKIND.CUST_COD ON SHIFT.SHIFT_ID = HEAD_BACK_ANBAR.SHIFT LEFT OUTER JOIN OTH"
        "ER_DTL ON HEAD_BACK_ANBAR.NUMBER = OTHER_DTL.NUMBER AND HEAD_BACK_ANBAR.HTAG = O"
        "THER_DTL.TAG LEFT OUTER JOIN SALA_DTL SALA_DTL_2 ON HEAD_BACK_ANBAR.sgn3usid = S"
        "ALA_DTL_2.IDD LEFT OUTER JOIN SALA_DTL SALA_DTL_1 ON HEAD_BACK_ANBAR.sgn2usid = "
        "SALA_DTL_1.IDD LEFT OUTER JOIN SALA_DTL ON HEAD_BACK_ANBAR.SGN1usid = SALA_DTL.I"
        "DD LEFT OUTER JOIN STUF_DEF STUF_DEF_1 INNER JOIN OTHER_DTL_SUB ON STUF_DEF_1.CO"
        "DE = OTHER_DTL_SUB.CODE ON INVO_LST.CODE = OTHER_DTL_SUB.CODE AND INVO_LST.NUMBE"
        "R = OTHER_DTL_SUB.NUMBER AND INVO_LST.TAG = OTHER_DTL_SUB.TAGG LEFT OUTER JOIN C"
        "UST_HESAB ON HEAD_BACK_ANBAR.CUST_NO = CUST_HESAB.hes LEFT OUTER JOIN TCOD_VAHED"
        "S ON INVO_LST.VAHED_K = TCOD_VAHEDS.CODE ON STUF_DEF.CODE = INVO_LST.CODE"
    Caption ="فاكتور فروش"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial (Arabic)"
    OnError ="[Event Procedure]"
    RecordSource ="SELECT HEAD_BACK_ANBAR.NUMBER1, HEAD_BACK_ANBAR.NUMBER, HEAD_BACK_ANBAR.DATE_N, "
        "INVO_LST.NUMBER AS INUMBER, HEAD_BACK_ANBAR.HTAG, INVO_LST.ANBAR, INVO_LST.RADIF"
        ", INVO_LST.CODE, INVO_LST.MEGH, INVO_LST.MEGHk, INVO_LST.MEGH_MAR, INVO_LST.MAND"
        "AH, INVO_LST.MABL, INVO_LST.MABL_K, INVO_LST.FROM_A, INVO_LST.N_RASID, INVO_LST."
        "MEGH_R, INVO_LST.RADAH, INVO_LST.SANAD_NO, INVO_LST.ANBARF, INVO_LST.VAHED_K, ST"
        "UF_DEF.NAME, TCOD_ANBAR.NAMES, TCOD_VAHEDS.NAMES AS VNAMES, HEAD_BACK_ANBAR.TAH,"
        " HEAD_BACK_ANBAR.MOLAH, CUSTKIND.CUSTKNAME, DEPART.DEPNAME, SHIFT.SHNAME, CUST_H"
        "ESAB.NAME AS HESAB, CUST_HESAB.ADDRESS, CUST_HESAB.TEL, ISNULL(STUF_DEF.NAME, N'"
        " ') + N' ' + ISNULL(INVO_LST.MANDAH, N' ') AS KALA, HEAD_BACK_ANBAR.CUST_NO, INV"
        "O_LST.N_KOL, INVO_LST.N_MOIN, HEAD_BACK_ANBAR.FNUMCO, CUST_HESAB.ECODE, CUST_HES"
        "AB.PCODE, CUST_HESAB.IYALAT, CUST_HESAB.MCODEM, CUST_HESAB.CITY, INVO_LST.MABL_K"
        " - INVO_LST.N_MOIN AS mabkbt, INVO_LST.IMBAA, INVO_LST.MABL_K - INVO_LST.N_MOIN "
        "+ INVO_LST.IMBAA AS mabkn, CUST_HESAB.CODE_E, HEAD_BACK_ANBAR.TAKHFIF, HEAD_BACK"
        "_ANBAR.MBAA, STUF_DEF.N_FANI, HEAD_BACK_ANBAR.SADER, HEAD_BACK_ANBAR.ANBARF AS A"
        "NBARFF, OTHER_DTL.REQUEST_NO, OTHER_DTL.BARNAMEH, OTHER_DTL.DRIVER, OTHER_DTL.DR"
        "IVER_MOB, OTHER_DTL.CAMIUN_NUM, OTHER_DTL.MAGHSAD, OTHER_DTL.CAM_KHALY, OTHER_DT"
        "L.CAM_POOR, OTHER_DTL.TOZIH, OTHER_DTL.CAMIUN, OTHER_DTL_SUB.CODE AS CODEG, OTHE"
        "R_DTL_SUB.CAM_KHALY AS CAM_KHALYG, OTHER_DTL_SUB.CAM_POOR AS CAM_POORG, OTHER_DT"
        "L_SUB.MEGHk AS MEGHkG, OTHER_DTL_SUB.TOZIH AS TOZIHG, OTHER_DTL_SUB.VAZNH, STUF_"
        "DEF_1.NAME AS NAMEG, HEAD_BACK_ANBAR.SHARAYET, DEPART.DEPART, INVO_LST.TKHN, HEA"
        "D_BACK_ANBAR.MAS, HEAD_BACK_ANBAR.HTAG AS TAG, SALA_DTL.EMZA AS EMZA1, SALA_DTL_"
        "1.EMZA AS EMZA2, SALA_DTL_2.EMZA AS EMZA3, HEAD_BACK_ANBAR.SGN1usid, HEAD_BACK_A"
        "NBAR.sgn2usid, HEAD_BACK_ANBAR.sgn3usid, HEAD_LST.SGN1, HEAD_LST.SGN2, HEAD_LST."
        "SGN3, HEAD_LST.SGN4 FROM STUF_DEF RIGHT OUTER JOIN SHIFT RIGHT OUTER JOIN DEPART"
        " RIGHT OUTER JOIN HEAD_LST INNER JOIN INVO_LST INNER JOIN TCOD_ANBAR ON INVO_LST"
        ".ANBAR = TCOD_ANBAR.CODE INNER JOIN HEAD_BACK_ANBAR ON INVO_LST.NUMBER = HEAD_BA"
        "CK_ANBAR.NUMBER AND INVO_LST.TAG = HEAD_BACK_ANBAR.HTAG ON HEAD_LST.NUMBER = HEA"
        "D_BACK_ANBAR.NUMBER AND HEAD_LST.TAG = HEAD_BACK_ANBAR.TAG ON DEPART.DEPATMAN = "
        "HEAD_BACK_ANBAR.DEPATMAN LEFT OUTER JOIN CUSTKIND ON HEAD_BACK_ANBAR.CUST_KIND ="
        " CUSTKIND.CUST_COD ON SHIFT.SHIFT_ID = HEAD_BACK_ANBAR.SHIFT LEFT OUTER JOIN OTH"
        "ER_DTL ON HEAD_BACK_ANBAR.NUMBER = OTHER_DTL.NUMBER AND HEAD_BACK_ANBAR.HTAG = O"
        "THER_DTL.TAG LEFT OUTER JOIN SALA_DTL SALA_DTL_2 ON HEAD_BACK_ANBAR.sgn3usid = S"
        "ALA_DTL_2.IDD LEFT OUTER JOIN SALA_DTL SALA_DTL_1 ON HEAD_BACK_ANBAR.sgn2usid = "
        "SALA_DTL_1.IDD LEFT OUTER JOIN SALA_DTL ON HEAD_BACK_ANBAR.SGN1usid = SALA_DTL.I"
        "DD LEFT OUTER JOIN STUF_DEF STUF_DEF_1 INNER JOIN OTHER_DTL_SUB ON STUF_DEF_1.CO"
        "DE = OTHER_DTL_SUB.CODE ON INVO_LST.CODE = OTHER_DTL_SUB.CODE AND INVO_LST.NUMBE"
        "R = OTHER_DTL_SUB.NUMBER AND INVO_LST.TAG = OTHER_DTL_SUB.TAGG LEFT OUTER JOIN C"
        "UST_HESAB ON HEAD_BACK_ANBAR.CUST_NO = CUST_HESAB.hes LEFT OUTER JOIN TCOD_VAHED"
        "S ON INVO_LST.VAHED_K = TCOD_VAHEDS.CODE ON STUF_DEF.CODE = INVO_LST.CODE"
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
        Begin BreakLevel
            ControlSource ="KALA"
        End
        Begin PageHeader
            Height =4020
            Name ="PageHeader"
            Begin
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =7892
                    Top =850
                    Width =1311
                    Height =375
                    FontSize =10
                    BorderColor =32768
                    Name ="DATE_N"
                    ControlSource ="DATE_N"
                    FontName ="WeblogmaYekan"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =9255
                            Top =850
                            Width =570
                            Height =390
                            FontSize =10
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
                    Left =7740
                    Top =465
                    Width =1476
                    Height =375
                    FontSize =10
                    TabIndex =1
                    BorderColor =32768
                    Name ="NUMBER"
                    ControlSource ="NUMBER"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =9252
                            Top =465
                            Width =1065
                            Height =390
                            FontSize =10
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
                    Top =1290
                    Width =10335
                    BorderColor =0
                    Name ="Line94"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2535
                    Top =2232
                    Width =7056
                    Height =375
                    FontSize =10
                    TabIndex =2
                    BorderColor =32768
                    Name ="HESAB"
                    ControlSource ="HESAB"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =9612
                            Top =2232
                            Width =675
                            Height =390
                            FontSize =10
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
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =75
                    Top =2220
                    Width =1791
                    Height =375
                    FontSize =10
                    TabIndex =3
                    BorderColor =32768
                    Name ="CUST_NO"
                    ControlSource ="CUST_NO"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =1992
                            Top =2220
                            Width =360
                            Height =390
                            FontSize =10
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
                    Left =3240
                    Top =2655
                    Width =6381
                    Height =375
                    FontSize =10
                    TabIndex =4
                    BorderColor =32768
                    Name ="ADDRESS"
                    ControlSource ="ADDRESS"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =9657
                            Top =2655
                            Width =630
                            Height =390
                            FontSize =10
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
                    Top =2625
                    Width =2736
                    Height =375
                    FontSize =10
                    TabIndex =5
                    BorderColor =32768
                    Name ="TEL"
                    ControlSource ="TEL"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =2772
                            Top =2625
                            Width =465
                            Height =390
                            FontSize =10
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
                    LineSlant = NotDefault
                    BorderWidth =3
                    Top =3495
                    Width =10335
                    BorderColor =0
                    Name ="Line99"
                End
                Begin Line
                    BorderWidth =3
                    Left =10335
                    Top =3510
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line130"
                End
                Begin Line
                    Left =9885
                    Top =3510
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line131"
                End
                Begin Line
                    Left =6030
                    Top =3510
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line133"
                End
                Begin Line
                    Left =5175
                    Top =3510
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line134"
                End
                Begin Line
                    Left =4065
                    Top =3510
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line135"
                End
                Begin Line
                    Left =3120
                    Top =3510
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line136"
                End
                Begin Line
                    Left =1500
                    Top =3510
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line137"
                End
                Begin Line
                    BorderWidth =3
                    Top =3510
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line138"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =405
                    Top =3555
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
                    Left =2370
                    Top =3555
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
                    Left =3210
                    Top =3555
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
                    Left =5352
                    Top =3555
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
                    Left =4242
                    Top =3555
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
                    Left =6787
                    Top =3555
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
                    Visible = NotDefault
                    TextAlign =2
                    TextFontFamily =2
                    Left =476
                    Top =3555
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
                    Vertical = NotDefault
                    TextAlign =2
                    TextFontFamily =2
                    Left =9936
                    Top =3551
                    Width =375
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
                    LineSlant = NotDefault
                    BorderWidth =3
                    Top =4020
                    Width =10335
                    BorderColor =0
                    Name ="Line147"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1335
                    Top =3060
                    Width =8136
                    Height =375
                    FontSize =10
                    TabIndex =6
                    BorderColor =32768
                    Name ="MOLAH1"
                    ControlSource ="=IIf(IsNull([MOLAH]),[MOLAH],HIDVISIT([MOLAH]))"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =9495
                            Top =3060
                            Width =810
                            Height =390
                            FontSize =10
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
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3300
                    Top =1302
                    Width =6171
                    Height =375
                    FontSize =10
                    TabIndex =7
                    BorderColor =32768
                    Name ="Text156"
                    ControlSource ="=[Forms]![Baseknow]![NAME]"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =9525
                            Top =1305
                            Width =765
                            Height =390
                            FontSize =10
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label157"
                            Caption ="فروشنده:"
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
                    Left =2320
                    Top =1710
                    Width =7266
                    Height =375
                    FontSize =10
                    TabIndex =8
                    BorderColor =32768
                    Name ="Text160"
                    ControlSource ="=[Forms]![Baseknow]![TFADDRESS]"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =9622
                            Top =1710
                            Width =630
                            Height =390
                            FontSize =10
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label161"
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
                    Left =105
                    Top =1320
                    Width =2631
                    Height =375
                    FontSize =10
                    TabIndex =9
                    BorderColor =32768
                    Name ="Text162"
                    ControlSource ="=[Forms]![Baseknow]![TFTEL]"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =2742
                            Top =1320
                            Width =465
                            Height =390
                            FontSize =10
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label163"
                            Caption ="تلفن:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin Line
                    BorderWidth =3
                    Top =2175
                    Width =10335
                    BorderColor =0
                    Name ="Line164"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =60
                    Top =1716
                    Width =1581
                    Height =375
                    FontSize =10
                    TabIndex =10
                    BorderColor =32768
                    Name ="DEPNAME"
                    ControlSource ="DEPNAME"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =1740
                            Top =1716
                            Width =525
                            Height =390
                            FontSize =10
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label206"
                            Caption ="واحد:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =1417
                    Top =56
                    Width =666
                    Height =255
                    FontSize =10
                    TabIndex =11
                    Name ="HTAG"
                    ControlSource ="HTAG"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    Left =2006
                    Top =3510
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line194"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =1590
                    Top =3525
                    Width =405
                    Height =435
                    FontSize =10
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label195"
                    Caption ="%ت"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =8340
                    Width =1131
                    Height =375
                    FontSize =10
                    TabIndex =12
                    BorderColor =32768
                    Name ="MOLAH"
                    ControlSource ="MOLAH"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =4335
                    Top =495
                    Width =1860
                    Height =615
                    FontSize =10
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label88"
                    Caption ="فاكتور فروش"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =15
                    Width =10293
                    Height =570
                    FontSize =10
                    TabIndex =13
                    BorderColor =32768
                    Name ="Text90"
                    ControlSource ="=[Forms]![Baseknow]![WIDTH_D]"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    Visible = NotDefault
                    TextAlign =3
                    TextFontFamily =2
                    Left =3
                    Top =863
                    Width =3105
                    Height =390
                    FontSize =10
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Ltvis"
                    Caption ="تلفن ويزيتور :"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    Visible = NotDefault
                    TextAlign =3
                    TextFontFamily =2
                    Top =478
                    Width =3105
                    Height =390
                    FontSize =10
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="lvisit"
                    Caption ="ويزيتور:"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =30
                    Top =3062
                    Width =771
                    Height =375
                    FontSize =10
                    TabIndex =14
                    BorderColor =32768
                    Name ="MAS"
                    ControlSource ="MAS"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =811
                            Top =3060
                            Width =480
                            Height =390
                            FontSize =10
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label211"
                            Caption ="مدت:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =421
            OnFormat ="[Event Procedure]"
            Name ="Detail"
            Begin
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =433
                    Top =16
                    Width =504
                    Height =405
                    FontSize =8
                    BorderColor =32768
                    Name ="ANBAR"
                    ControlSource ="ANBAR"
                    Format ="Fixed"
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
                    Left =9871
                    Top =16
                    Width =414
                    Height =405
                    FontSize =8
                    TabIndex =1
                    BorderColor =32768
                    Name ="RADIFwe"
                    ControlSource ="=1"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =6072
                    Top =16
                    Width =3759
                    Height =405
                    FontSize =8
                    TabIndex =2
                    BorderColor =32768
                    Name ="CODE1"
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
                    Left =5190
                    Top =15
                    Width =819
                    Height =405
                    FontSize =12
                    TabIndex =3
                    BorderColor =32768
                    Name ="MEGH"
                    ControlSource ="MEGH"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3135
                    Top =16
                    Width =894
                    Height =405
                    FontSize =12
                    TabIndex =4
                    BorderColor =32768
                    Name ="MEGHk"
                    ControlSource ="MEGHk"
                    Format ="Fixed"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2029
                    Top =16
                    Width =1059
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
                    Top =16
                    Width =1464
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
                    Left =4040
                    Top =15
                    Width =1110
                    Height =405
                    FontSize =8
                    TabIndex =7
                    BorderColor =32768
                    Name ="VAHED_K"
                    ControlSource ="VNAMES"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    BorderWidth =1
                    Top =420
                    Width =10335
                    BorderColor =0
                    Name ="Line100"
                End
                Begin Line
                    BorderWidth =3
                    Left =10335
                    Width =0
                    Height =421
                    BorderColor =0
                    Name ="Line101"
                End
                Begin Line
                    BorderWidth =1
                    Left =9885
                    Width =0
                    Height =421
                    BorderColor =0
                    Name ="Line102"
                End
                Begin Line
                    BorderWidth =1
                    Left =6030
                    Width =0
                    Height =421
                    BorderColor =0
                    Name ="Line104"
                End
                Begin Line
                    BorderWidth =1
                    Left =5175
                    Width =0
                    Height =421
                    BorderColor =0
                    Name ="Line105"
                End
                Begin Line
                    BorderWidth =1
                    Left =4065
                    Width =0
                    Height =421
                    BorderColor =0
                    Name ="Line106"
                End
                Begin Line
                    BorderWidth =1
                    Left =3120
                    Width =0
                    Height =421
                    BorderColor =0
                    Name ="Line107"
                End
                Begin Line
                    BorderWidth =1
                    Left =1500
                    Width =0
                    Height =421
                    BorderColor =0
                    Name ="Line108"
                End
                Begin Line
                    BorderWidth =2
                    Width =0
                    Height =421
                    BorderColor =0
                    Name ="Line109"
                End
                Begin Line
                    BorderWidth =1
                    Left =2006
                    Width =0
                    Height =421
                    BorderColor =0
                    Name ="Line197"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1530
                    Width =459
                    Height =225
                    FontSize =8
                    TabIndex =8
                    BorderColor =32768
                    Name ="N_KOL"
                    ControlSource ="N_KOL"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    Left =1515
                    Top =240
                    Width =491
                    Name ="Line219"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1530
                    Top =180
                    Width =459
                    Height =225
                    FontSize =8
                    TabIndex =9
                    BorderColor =32768
                    Name ="TKHN"
                    ControlSource ="TKHN"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Left =3118
                    Top =150
                    TabIndex =10
                    Name ="KALA"
                    ControlSource ="KALA"
                End
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Left =5084
                    Top =124
                    TabIndex =11
                    Name ="CODE"
                    ControlSource ="CODE"
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =5731
            OnFormat ="[Event Procedure]"
            Name ="GroupFooter3"
            Begin
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =3
                    Top =30
                    Width =10335
                    BorderColor =0
                    Name ="Line110"
                End
                Begin Line
                    BorderWidth =2
                    Left =3450
                    Top =30
                    Width =0
                    Height =2761
                    BorderColor =0
                    Name ="Line118"
                End
                Begin Line
                    BorderWidth =2
                    Width =0
                    Height =2776
                    BorderColor =0
                    Name ="Line119"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =3
                    Top =2790
                    Width =10335
                    BorderColor =0
                    Name ="Line120"
                End
                Begin Line
                    Top =420
                    Width =3435
                    BorderColor =0
                    Name ="Line121"
                End
                Begin Line
                    Left =15
                    Top =810
                    Width =3435
                    BorderColor =0
                    Name ="Line122"
                End
                Begin Line
                    Left =15
                    Top =1155
                    Width =3435
                    BorderColor =0
                    Name ="Line123"
                End
                Begin Line
                    Left =15
                    Top =1935
                    Width =3435
                    BorderColor =0
                    Name ="Line124"
                End
                Begin Line
                    Left =15
                    Top =2310
                    Width =3435
                    BorderColor =0
                    Name ="Line125"
                End
                Begin Line
                    Left =1770
                    Width =0
                    Height =2746
                    BorderColor =0
                    Name ="Line126"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =3
                    Left =10335
                    Width =0
                    Height =2836
                    BorderColor =0
                    Name ="Line127"
                End
                Begin Line
                    LineSlant = NotDefault
                    Top =810
                    Width =3436
                    BorderColor =0
                    Name ="Line128"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =2
                    Left =9345
                    Top =1257
                    Width =900
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
                    Left =3401
                    Top =1278
                    Width =5865
                    Height =930
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="COMM"
                    Caption ="."
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
                    Left =45
                    Top =3540
                    Width =10278
                    Height =540
                    FontSize =10
                    BorderColor =32768
                    Name ="Text153"
                    ControlSource ="=[Forms]![Baseknow]![HIGH_D]"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    BorderWidth =1
                    TextAlign =0
                    TextFontFamily =2
                    Left =1817
                    Top =1920
                    Width =1725
                    Height =375
                    FontSize =10
                    FontWeight =400
                    BorderColor =4210752
                    ForeColor =0
                    Name ="Label177"
                    Caption ="جمع مبالغ پرداختي:"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    BorderWidth =1
                    TextAlign =0
                    TextFontFamily =2
                    Left =1823
                    Top =56
                    Width =1455
                    Height =345
                    FontSize =10
                    FontWeight =400
                    BorderColor =4210752
                    ForeColor =0
                    Name ="Label178"
                    Caption ="جمع كل فاكتور:"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    BorderWidth =1
                    TextAlign =0
                    TextFontFamily =2
                    Left =1815
                    Top =795
                    Width =1635
                    Height =420
                    FontSize =10
                    FontWeight =400
                    BorderColor =4210752
                    ForeColor =0
                    Name ="Label179"
                    Caption ="خدمات:"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    BorderWidth =1
                    TextAlign =0
                    TextFontFamily =2
                    Left =1817
                    Top =435
                    Width =1230
                    Height =345
                    FontSize =10
                    FontWeight =400
                    BorderColor =4210752
                    ForeColor =0
                    Name ="Label180"
                    Caption ="تخفيفات:"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    BorderWidth =1
                    TextAlign =0
                    TextFontFamily =2
                    Left =1817
                    Top =1500
                    Width =1710
                    Height =375
                    FontSize =10
                    FontWeight =400
                    BorderColor =4210752
                    ForeColor =0
                    Name ="Label181"
                    Caption ="مبلغ قابل پرداخت:"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    BorderWidth =1
                    TextAlign =0
                    TextFontFamily =2
                    Left =1817
                    Top =2325
                    Width =1230
                    Height =375
                    FontSize =10
                    FontWeight =400
                    BorderColor =4210752
                    ForeColor =0
                    Name ="Label182"
                    Caption ="مانده فاكتور:"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    BorderWidth =1
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    NumeralShapes =2
                    Top =1920
                    Width =1756
                    Height =345
                    FontSize =10
                    FontWeight =400
                    BorderColor =4210752
                    ForeColor =0
                    Name ="JPAY"
                    Caption ="جمع مبالغ پرداختي:"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    BorderWidth =1
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    NumeralShapes =2
                    Left =6
                    Top =56
                    Width =1756
                    Height =345
                    FontSize =10
                    FontWeight =400
                    BorderColor =4210752
                    ForeColor =0
                    Name ="JF"
                    Caption ="جمع كل فاكتور:"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    BorderWidth =1
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    NumeralShapes =2
                    Top =795
                    Width =1756
                    Height =345
                    FontSize =10
                    FontWeight =400
                    BorderColor =4210752
                    ForeColor =0
                    Name ="HKH"
                    Caption ="هزينه خدمات:"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    BorderWidth =1
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    NumeralShapes =2
                    Top =435
                    Width =1756
                    Height =345
                    FontSize =10
                    FontWeight =400
                    BorderColor =4210752
                    ForeColor =0
                    Name ="TF"
                    Caption ="تخفيفات:"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    BorderWidth =1
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    NumeralShapes =2
                    Top =1500
                    Width =1756
                    Height =375
                    FontSize =10
                    FontWeight =400
                    BorderColor =4210752
                    ForeColor =0
                    Name ="GABEL"
                    Caption ="مبلغ قابل پرداخت:"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    BorderWidth =1
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    NumeralShapes =2
                    Top =2325
                    Width =1756
                    Height =375
                    FontSize =10
                    FontWeight =400
                    BorderColor =4210752
                    ForeColor =0
                    Name ="MAN"
                    Caption ="مانده فاكتور:"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =2
                    Left =9075
                    Top =450
                    Width =1230
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
                    ReadingOrder =2
                    NumeralShapes =2
                    Left =3495
                    Top =450
                    Width =5520
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
                    TextAlign =3
                    TextFontFamily =2
                    Left =6881
                    Top =2325
                    Width =3405
                    Height =435
                    FontSize =10
                    FontWeight =400
                    BackColor =128
                    ForeColor =0
                    Name ="MANDAH"
                    Caption ="مانده حساب:"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =2
                    Left =3405
                    Top =2325
                    Width =3405
                    Height =435
                    FontSize =10
                    FontWeight =400
                    BackColor =128
                    ForeColor =0
                    Name ="MANDG"
                    Caption ="مانده حساب:"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =1
                    BorderWidth =1
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3465
                    Top =15
                    Width =894
                    Height =421
                    FontSize =10
                    TabIndex =1
                    BorderColor =4210752
                    Name ="Text200"
                    ControlSource ="=Sum([MEGHk])"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    Left =15
                    Top =1530
                    Width =3435
                    BorderColor =0
                    Name ="Line201"
                End
                Begin Label
                    BorderWidth =1
                    TextAlign =0
                    TextFontFamily =2
                    Left =1815
                    Top =1140
                    Width =1920
                    Height =420
                    FontSize =10
                    FontWeight =400
                    BorderColor =4210752
                    ForeColor =0
                    Name ="Label224"
                    Caption ="3% ماليات و عوارض:"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    BorderWidth =1
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    NumeralShapes =2
                    Top =1170
                    Width =1756
                    Height =345
                    FontSize =10
                    FontWeight =400
                    BorderColor =4210752
                    ForeColor =0
                    Name ="MBAA"
                    Caption ="ماليات"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    NumeralShapes =2
                    Left =6406
                    Top =2852
                    Width =3870
                    Height =315
                    FontSize =8
                    FontWeight =400
                    BackColor =128
                    ForeColor =0
                    Name ="Label197"
                    Caption ="نرم افزار اتوماتيك Mr.CORRECT تلفن  03538421104"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    DecimalPlaces =2
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5730
                    Top =3150
                    Width =4554
                    Height =375
                    FontSize =10
                    TabIndex =2
                    BorderColor =32768
                    Name ="FACTORS"
                    ControlSource ="=\"=\""
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =2
                    Left =3435
                    Top =2835
                    Width =2970
                    Height =345
                    FontSize =10
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="vazn"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =1772
                    Top =2815
                    Width =1086
                    Height =345
                    FontSize =10
                    TabIndex =3
                    Name ="USERNAME"
                    ControlSource ="=ucurrentuser()"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextAlign =0
                            TextFontFamily =2
                            Left =2899
                            Top =2807
                            Width =480
                            Height =375
                            FontSize =10
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
                    Top =2815
                    Width =1206
                    Height =345
                    FontSize =10
                    TabIndex =4
                    Name ="Text206"
                    ControlSource ="=Now()"
                    Format ="Short Time"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextAlign =0
                            TextFontFamily =2
                            Left =1242
                            Top =2810
                            Width =510
                            Height =375
                            FontSize =10
                            FontWeight =400
                            Name ="Label207"
                            Caption ="ساعت:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =1
                    BorderWidth =1
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5175
                    Width =849
                    Height =421
                    FontSize =10
                    TabIndex =5
                    BorderColor =4210752
                    Name ="Text212"
                    ControlSource ="=Sum([MEGH])"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =3671
                    Top =1301
                    Width =5616
                    Height =930
                    FontSize =10
                    TabIndex =6
                    Name ="SHARAYET"
                    ControlSource ="SHARAYET"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TextAlign =1
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Top =3150
                    Width =5709
                    Height =375
                    FontSize =10
                    TabIndex =7
                    BorderColor =32768
                    Name ="PAYMENTS"
                    ControlSource ="=\"=\""
                    FontName ="WeblogmaYekan"
                End
                Begin BoundObjectFrame
                    SizeMode =3
                    OldBorderStyle =0
                    Left =6840
                    Top =4005
                    Width =1701
                    Height =1701
                    TabIndex =8
                    Name ="EMZA1"
                    ControlSource ="EMZA1"
                End
                Begin BoundObjectFrame
                    SizeMode =3
                    OldBorderStyle =0
                    Left =3435
                    Top =3990
                    Width =1701
                    Height =1701
                    TabIndex =9
                    Name ="EMZA2"
                    ControlSource ="EMZA2"
                End
                Begin BoundObjectFrame
                    SizeMode =3
                    OldBorderStyle =0
                    Top =3986
                    Width =1701
                    Height =1701
                    TabIndex =10
                    Name ="EMZA3"
                    ControlSource ="EMZA3"
                End
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Left =8220
                    Top =3303
                    TabIndex =11
                    Name ="SGN1usid"
                    ControlSource ="SGN1usid"
                End
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Left =8220
                    Top =3622
                    TabIndex =12
                    Name ="sgn2usid"
                    ControlSource ="sgn2usid"
                End
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Left =7361
                    Top =3658
                    TabIndex =13
                    Name ="sgn3usid"
                    ControlSource ="sgn3usid"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =8550
                    Top =4425
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
                    Top =4425
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
                    Top =4425
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
                    Top =4092
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
                    Top =4092
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
                    Top =4092
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
                    Left =4756
                    Top =3543
                    TabIndex =14
                    Name ="SGN1"
                    ControlSource ="SGN1"
                    Begin
                        Begin Label
                            TextAlign =0
                            Left =3055
                            Top =3543
                            Width =630
                            Height =255
                            Name ="Label225"
                            Caption ="SGN1:"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Left =4756
                    Top =3862
                    TabIndex =15
                    Name ="SGN2"
                    ControlSource ="SGN2"
                    Begin
                        Begin Label
                            TextAlign =0
                            Left =3055
                            Top =3862
                            Width =630
                            Height =255
                            Name ="Label226"
                            Caption ="SGN2:"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Left =4756
                    Top =4181
                    TabIndex =16
                    Name ="SGN3"
                    ControlSource ="SGN3"
                    Begin
                        Begin Label
                            TextAlign =0
                            Left =3055
                            Top =4181
                            Width =630
                            Height =255
                            Name ="Label227"
                            Caption ="SGN3:"
                        End
                    End
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
' See "INVOICE_FROOSH_22.cls"
