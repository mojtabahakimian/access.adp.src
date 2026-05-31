Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    FilterOn = NotDefault
    DefaultView =0
    AllowUpdating =4
    TabularCharSet =178
    TabularFamily =42
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =60
    GridY =60
    Width =6396
    DatasheetFontHeight =10
    ItemSuffix =226
    Left =1260
    Top =1590
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x251cdc2233ede540
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
            ControlSource ="[NUMBER]"
        End
        Begin BreakLevel
            ControlSource ="RADIF"
        End
        Begin PageHeader
            Height =2415
            OnFormat ="[Event Procedure]"
            Name ="PageHeader"
            Begin
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =2460
                    Width =1125
                    Height =330
                    FontSize =8
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
                    Left =1359
                    Top =300
                    Width =3333
                    Height =330
                    FontSize =8
                    TabIndex =1
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
                    Left =28
                    Top =302
                    Width =846
                    Height =300
                    FontSize =8
                    TabIndex =2
                    BorderColor =32768
                    Name ="DATE_N"
                    ControlSource ="DATE_N"
                    FontName ="WeblogmaYekan"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =911
                            Top =302
                            Width =465
                            Height =330
                            FontSize =8
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
                    Left =4700
                    Top =300
                    Width =681
                    Height =300
                    FontSize =8
                    TabIndex =3
                    BorderColor =32768
                    Name ="NUMBER1"
                    ControlSource ="NUMBER1"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =5417
                            Top =300
                            Width =930
                            Height =330
                            FontSize =7
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label93"
                            Caption ="شماره فاكتور:"
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
                    Left =2396
                    Top =1365
                    Width =3366
                    Height =300
                    FontSize =7
                    TabIndex =4
                    BorderColor =32768
                    Name ="HESAB"
                    ControlSource ="HESAB"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =5813
                            Top =1365
                            Width =540
                            Height =330
                            FontSize =7
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
                    Left =30
                    Top =1360
                    Width =966
                    Height =300
                    FontSize =7
                    TabIndex =5
                    BorderColor =32768
                    Name ="CUST_NO"
                    ControlSource ="CUST_NO"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =987
                            Top =1360
                            Width =300
                            Height =330
                            FontSize =7
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label96"
                            Caption ="كد:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Top =2010
                    Width =6366
                    BorderColor =0
                    Name ="Line99"
                End
                Begin Line
                    BorderWidth =1
                    Left =6341
                    Top =2025
                    Width =0
                    Height =390
                    BorderColor =0
                    Name ="Line130"
                End
                Begin Line
                    Left =5891
                    Top =2025
                    Width =0
                    Height =390
                    BorderColor =0
                    Name ="Line131"
                End
                Begin Line
                    Left =1500
                    Top =2025
                    Width =0
                    Height =390
                    BorderColor =0
                    Name ="Line132"
                End
                Begin Line
                    Left =3080
                    Top =2025
                    Width =0
                    Height =390
                    BorderColor =0
                    Name ="Line133"
                End
                Begin Line
                    Left =2465
                    Top =2025
                    Width =0
                    Height =390
                    BorderColor =0
                    Name ="Line136"
                End
                Begin Line
                    Left =1140
                    Top =2025
                    Width =0
                    Height =390
                    BorderColor =0
                    Name ="Line137"
                End
                Begin Line
                    BorderWidth =1
                    Top =2025
                    Width =0
                    Height =390
                    BorderColor =0
                    Name ="Line138"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =300
                    Top =2055
                    Width =525
                    Height =330
                    FontSize =7
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
                    Left =1678
                    Top =2070
                    Width =555
                    Height =330
                    FontSize =7
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
                    Left =2465
                    Top =2070
                    Width =600
                    Height =330
                    FontSize =7
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
                    Left =3605
                    Top =2055
                    Width =1680
                    Height =330
                    FontSize =7
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
                    Left =5520
                    Top =2059
                    Width =315
                    Height =330
                    FontSize =7
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
                    Left =5925
                    Top =2070
                    Width =390
                    Height =330
                    FontSize =7
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label146"
                    Caption ="رديف"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Top =2400
                    Width =6366
                    BorderColor =0
                    Name ="Line147"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2366
                    Top =1680
                    Width =3246
                    Height =300
                    FontSize =7
                    BorderColor =32768
                    Name ="MOLAH"
                    ControlSource ="MOLAH"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =5636
                            Top =1680
                            Width =720
                            Height =330
                            FontSize =7
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label155"
                            Caption ="ملاحظات:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =2
                    Top =675
                    Width =6381
                    BorderColor =0
                    Name ="Line164"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =35
                    Top =1660
                    Width =906
                    Height =300
                    FontSize =7
                    TabIndex =6
                    BorderColor =32768
                    Name ="TEL"
                    ControlSource ="TEL"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =935
                            Top =1660
                            Width =450
                            Height =330
                            FontSize =7
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label194"
                            Caption ="تلفن:"
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
                    Left =2810
                    Top =710
                    Width =2826
                    Height =300
                    FontSize =7
                    TabIndex =7
                    BorderColor =32768
                    Name ="Text195"
                    ControlSource ="=[Forms]![Baseknow]![NAME]"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =5643
                            Top =715
                            Width =720
                            Height =330
                            FontSize =7
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label196"
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
                    Left =1001
                    Top =705
                    Width =1356
                    Height =300
                    FontSize =7
                    TabIndex =8
                    BorderColor =32768
                    Name ="Text197"
                    ControlSource ="=[Forms]![Baseknow]![TFTEL]"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =2351
                            Top =705
                            Width =480
                            Height =330
                            FontSize =7
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label198"
                            Caption ="تلفن:"
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
                    Left =2655
                    Top =1020
                    Width =3186
                    Height =300
                    FontSize =7
                    TabIndex =9
                    BorderColor =32768
                    Name ="Text199"
                    ControlSource ="=[Forms]![Baseknow]![TFADDRESS]"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =5835
                            Top =1020
                            Width =540
                            Height =330
                            FontSize =7
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label200"
                            Caption ="آدرس:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Left =35
                    Top =1345
                    Width =6333
                    BorderColor =0
                    Name ="Line201"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1200
                    Top =990
                    Width =846
                    Height =345
                    FontSize =7
                    TabIndex =10
                    BorderColor =32768
                    Name ="DEPNAME"
                    ControlSource ="DEPNAME"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =2055
                            Top =990
                            Width =375
                            Height =330
                            FontSize =7
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
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Top =724
                    Width =531
                    Height =300
                    FontSize =7
                    TabIndex =11
                    BorderColor =32768
                    Name ="SHNAME"
                    ControlSource ="SHNAME"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =540
                            Top =725
                            Width =495
                            Height =300
                            FontSize =7
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label208"
                            Caption ="شيفت:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =15
                    Top =1024
                    Width =786
                    Height =300
                    FontSize =7
                    TabIndex =12
                    BorderColor =32768
                    Name ="USER"
                    ControlSource ="=UCURRENTUSER()"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =825
                            Top =1024
                            Width =360
                            Height =300
                            FontSize =7
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label210"
                            Caption ="كاربر:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin Line
                    Left =5465
                    Top =2025
                    Width =0
                    Height =390
                    BorderColor =0
                    Name ="Line212"
                End
                Begin Label
                    Visible = NotDefault
                    TextAlign =2
                    TextFontFamily =2
                    Left =5924
                    Top =2055
                    Width =240
                    Height =330
                    FontSize =7
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label213"
                    Caption ="كد"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =4233
                    Top =28
                    Height =225
                    FontSize =8
                    TabIndex =13
                    Name ="HTAG"
                    ControlSource ="HTAG"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =1140
                    Top =2055
                    Width =345
                    Height =330
                    FontSize =7
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label195"
                    Caption ="%ت"
                    FontName ="WeblogmaYekan"
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =375
            OnFormat ="[Event Procedure]"
            Name ="Detail"
            Begin
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5473
                    Top =16
                    Width =399
                    Height =330
                    FontSize =8
                    BorderColor =32768
                    Name ="ANBAR"
                    ControlSource ="ANBAR"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =0
                    RunningSum =1
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5925
                    Top =15
                    Width =414
                    Height =330
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
                    Left =3113
                    Top =16
                    Width =2319
                    Height =330
                    FontSize =8
                    TabIndex =2
                    BorderColor =32768
                    Name ="CODE1"
                    ControlSource ="=IIf([Forms]![baseknow]![codeview]=1,[kala] & \" \" & [code],[kala])"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2490
                    Top =15
                    Width =564
                    Height =330
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
                    Left =1515
                    Top =15
                    Width =924
                    Height =330
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
                    Top =15
                    Width =1044
                    Height =330
                    FontSize =8
                    TabIndex =5
                    BorderColor =32768
                    Name ="MABL_K"
                    ControlSource ="MABL_K"
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    BorderWidth =1
                    Top =375
                    Width =6366
                    BorderColor =0
                    Name ="Line100"
                End
                Begin Line
                    BorderWidth =1
                    Left =6341
                    Width =0
                    Height =375
                    BorderColor =0
                    Name ="Line101"
                End
                Begin Line
                    BorderWidth =1
                    Left =5891
                    Width =0
                    Height =375
                    BorderColor =0
                    Name ="Line102"
                End
                Begin Line
                    BorderWidth =1
                    Left =1500
                    Width =0
                    Height =375
                    BorderColor =0
                    Name ="Line103"
                End
                Begin Line
                    BorderWidth =1
                    Left =3080
                    Width =0
                    Height =375
                    BorderColor =0
                    Name ="Line104"
                End
                Begin Line
                    BorderWidth =1
                    Left =2465
                    Width =0
                    Height =375
                    BorderColor =0
                    Name ="Line107"
                End
                Begin Line
                    BorderWidth =1
                    Left =1140
                    Width =0
                    Height =375
                    BorderColor =0
                    Name ="Line108"
                End
                Begin Line
                    BorderWidth =1
                    Width =0
                    Height =375
                    BorderColor =0
                    Name ="Line109"
                End
                Begin TextBox
                    Visible = NotDefault
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5896
                    Top =16
                    Width =309
                    Height =330
                    FontSize =8
                    TabIndex =6
                    BorderColor =32768
                    Name ="Text214"
                    ControlSource ="CODE"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    BorderWidth =1
                    Left =5465
                    Width =0
                    Height =375
                    BorderColor =0
                    Name ="Line215"
                End
                Begin TextBox
                    Visible = NotDefault
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =2830
                    TabIndex =7
                    Name ="CODE"
                    ControlSource ="CODE"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1170
                    Top =16
                    Width =339
                    Height =330
                    FontSize =8
                    TabIndex =8
                    BorderColor =32768
                    Name ="N_KOL"
                    ControlSource ="N_KOL"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =4146
            OnFormat ="[Event Procedure]"
            Name ="GroupFooter3"
            Begin
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Width =6325
                    BorderColor =0
                    Name ="Line110"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Left =2250
                    Width =0
                    Height =2296
                    BorderColor =0
                    Name ="Line118"
                End
                Begin Line
                    BorderWidth =1
                    Width =0
                    Height =2281
                    BorderColor =0
                    Name ="Line119"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Top =2280
                    Width =6355
                    BorderColor =0
                    Name ="Line120"
                End
                Begin Line
                    Top =330
                    Width =2265
                    BorderColor =0
                    Name ="Line121"
                End
                Begin Line
                    Top =630
                    Width =2265
                    BorderColor =0
                    Name ="Line122"
                End
                Begin Line
                    Top =975
                    Width =2265
                    BorderColor =0
                    Name ="Line123"
                End
                Begin Line
                    Top =1620
                    Width =2265
                    BorderColor =0
                    Name ="Line124"
                End
                Begin Line
                    Top =1950
                    Width =2265
                    BorderColor =0
                    Name ="Line125"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =1125
                    Width =0
                    Height =2266
                    BorderColor =0
                    Name ="Line126"
                End
                Begin Line
                    BorderWidth =1
                    Left =6345
                    Width =0
                    Height =2296
                    BorderColor =0
                    Name ="Line127"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =2245
                    Top =720
                    Width =4098
                    BorderColor =0
                    Name ="Line128"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =2
                    Left =2300
                    Top =750
                    Width =4035
                    Height =1140
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
                    Left =15
                    Top =2640
                    Width =6378
                    Height =330
                    FontSize =8
                    BorderColor =32768
                    Name ="Text153"
                    ControlSource ="=[Forms]![Baseknow]![HIGH_D]"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =0
                    TextFontFamily =2
                    Left =1169
                    Top =1635
                    Width =1335
                    Height =345
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="Label177"
                    Caption ="جمع  پرداختي:"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =0
                    TextFontFamily =2
                    Left =1170
                    Top =15
                    Width =1080
                    Height =345
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="Label178"
                    Caption ="جمع كل فاكتور:"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =0
                    TextFontFamily =2
                    Left =1170
                    Top =345
                    Width =555
                    Height =345
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="Label179"
                    Caption ="خدمات:"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =0
                    TextFontFamily =2
                    Left =1166
                    Top =660
                    Width =645
                    Height =345
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="Label180"
                    Caption ="تخفيفات:"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =0
                    TextFontFamily =2
                    Left =1155
                    Top =1305
                    Width =1290
                    Height =345
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="Label181"
                    Caption ="قابل پرداخت:"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =0
                    TextFontFamily =2
                    Left =1160
                    Top =1965
                    Width =900
                    Height =345
                    FontSize =10
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
                    Left =45
                    Top =1635
                    Width =1051
                    Height =300
                    FontSize =8
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
                    Left =45
                    Width =1051
                    Height =300
                    FontSize =8
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
                    Left =45
                    Top =349
                    Width =1051
                    Height =300
                    FontSize =8
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
                    Left =45
                    Top =660
                    Width =1051
                    Height =300
                    FontSize =8
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
                    Left =45
                    Top =1305
                    Width =1051
                    Height =300
                    FontSize =8
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
                    Left =30
                    Top =1965
                    Width =1066
                    Height =300
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="MAN"
                    Caption ="مانده فاكتور:"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =2
                    Left =5462
                    Top =56
                    Width =840
                    Height =330
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
                    Left =2844
                    Top =56
                    Width =2625
                    Height =675
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="HR"
                    Caption ="مبلغ به حروف:"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    NumeralShapes =2
                    Left =3689
                    Top =2314
                    Width =2670
                    Height =315
                    FontSize =8
                    FontWeight =400
                    BackColor =128
                    ForeColor =0
                    Name ="Label197"
                    Caption ="نرم افزار اتوماتيك نگين تلفن  8259800-0351"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =15
                    Top =2295
                    Width =2529
                    Height =330
                    FontSize =8
                    TabIndex =1
                    BorderColor =32768
                    Name ="Text211"
                    ControlSource ="=Format(Now(),\"hh:nn:ss ampm\")"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =1
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2265
                    Width =564
                    Height =330
                    FontSize =8
                    TabIndex =2
                    BorderColor =4210752
                    Name ="Text222"
                    ControlSource ="=Sum([MEGHk])"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    Top =1320
                    Width =2265
                    BorderColor =0
                    Name ="Line223"
                End
                Begin Label
                    TextAlign =0
                    TextFontFamily =2
                    Left =1166
                    Top =1005
                    Width =1425
                    Height =345
                    FontSize =7
                    FontWeight =400
                    ForeColor =0
                    Name ="Label224"
                    Caption ="3%ماليات و عوارض:"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    NumeralShapes =2
                    Left =45
                    Top =1005
                    Width =1051
                    Height =300
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="MBAA"
                    Caption ="ماليات:"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =2
                    Left =2312
                    Top =1908
                    Width =4020
                    Height =360
                    FontWeight =400
                    BackColor =128
                    ForeColor =0
                    Name ="MANDAH"
                    Caption ="مانده حساب:"
                    FontName ="WeblogmaYekan"
                End
                Begin BoundObjectFrame
                    SizeMode =3
                    OldBorderStyle =0
                    Left =4260
                    Top =2970
                    Width =1236
                    Height =1176
                    TabIndex =3
                    Name ="EMZA1"
                    ControlSource ="EMZA1"
                End
                Begin BoundObjectFrame
                    SizeMode =3
                    OldBorderStyle =0
                    Left =2115
                    Top =2970
                    Width =1236
                    Height =1176
                    TabIndex =4
                    Name ="EMZA2"
                    ControlSource ="EMZA2"
                End
                Begin BoundObjectFrame
                    SizeMode =3
                    OldBorderStyle =0
                    Top =2966
                    Width =1236
                    Height =1176
                    TabIndex =5
                    Name ="EMZA3"
                    ControlSource ="EMZA3"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =5417
                    Top =3318
                    Width =960
                    Height =240
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
                    Left =3288
                    Top =3347
                    Width =960
                    Height =240
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
                    Left =1124
                    Top =3334
                    Width =960
                    Height =240
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
                    Left =5414
                    Top =2985
                    Width =960
                    Height =240
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
                    Left =3290
                    Top =3014
                    Width =960
                    Height =240
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
                    Left =1124
                    Top =3004
                    Width =960
                    Height =240
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
                    Left =3618
                    Top =1332
                    TabIndex =6
                    Name ="SGN1usid"
                    ControlSource ="SGN1usid"
                End
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Left =3618
                    Top =1651
                    TabIndex =7
                    Name ="sgn2usid"
                    ControlSource ="sgn2usid"
                End
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Left =2759
                    Top =1687
                    TabIndex =8
                    Name ="sgn3usid"
                    ControlSource ="sgn3usid"
                End
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Left =4337
                    Top =784
                    TabIndex =9
                    Name ="SGN1"
                    ControlSource ="SGN1"
                End
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Left =4337
                    Top =1103
                    TabIndex =10
                    Name ="SGN2"
                    ControlSource ="SGN2"
                End
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Left =4337
                    Top =1422
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
' See "INVOICE_FROOSH2_1.cls"
