Version =20
VersionRequired =20
Begin Form
    AutoCenter = NotDefault
    FilterOn = NotDefault
    DefaultView =2
    AllowUpdating =4
    ViewsAllowed =2
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =5100
    RowHeight =300
    ItemSuffix =92
    Left =6915
    Top =2550
    Right =21405
    Bottom =6810
    DatasheetGridlinesColor =12632256
    MaxRecords =1000000
    RecSrcDt = Begin
        0xae65816094c9e540
    End
    RecordSource ="SELECT HEAD_LST.NUMBER, HEAD_LST.TAG, HEAD_LST.DATE_N, HEAD_LST.CUST_NO, HEAD_LS"
        "T.MOLAH, HEAD_LST.M_NAGHD, HEAD_LST.MABL_VAR, HEAD_LST.MABL_HAV, HEAD_LST.MABL_H"
        "AZ, HEAD_LST.TAKHFIF, HEAD_LST.FNUMCO, HEAD_LST.USER_NAME, HEAD_LST.MBAA, HEAD_L"
        "ST.OKF, HEAD_LST.CDDATE, HEAD_LST.CDTIME, HEAD_LST.OKDATE, HEAD_LST.OKTIME, SUM("
        "INVO_LST.MEGHk) AS MEGHkS, SUM(INVO_LST.MABL_K) AS MABL_KS, CUST_HESAB.NAME, CUS"
        "T_HESAB.ADDRESS, CUST_HESAB.TEL, CUST_HESAB.CODE_E, CUST_HESAB.ECODE, CUST_HESAB"
        ".PCODE, CUST_HESAB.IYALAT, CUST_HESAB.CITY, CUST_HESAB.MCODEM, CUST_HESAB.TOZIH,"
        " CUST_HESAB.MOBILE, HEAD_LST.N_S, HEAD_LST.VAS, HEAD_LST.MAS, HEAD_LST.TAH, HEAD"
        "_LST.NUMBER1, HEAD_LST.ANBAR, HEAD_LST.MOIN_VAR, HEAD_LST.MOIN_HAV, HEAD_LST.MOI"
        "N_HAZ, HEAD_LST.MOIN_KHF, HEAD_LST.ANBARF, HEAD_LST.DEPATMAN, HEAD_LST.SHIFT, HE"
        "AD_LST.CUST_KIND, HEAD_LST.SHARAYET, HEAD_LST.SGN1, HEAD_LST.SGN2, HEAD_LST.SGN3"
        ", HEAD_LST.SGN4, HEAD_LST.HMBAA, HEAD_LST.TAMIR, HEAD_LST.TICMBAA, HEAD_LST.TKHF"
        ", HEAD_LST.SADER, HEAD_LST.ARZD, HEAD_LST.ARZKIND, HEAD_LST.JAY, HEAD_LST.MODAT_"
        "PPID, HEAD_LST.PEPID, HEAD_LST.PEID FROM HEAD_LST INNER JOIN INVO_LST ON HEAD_LS"
        "T.NUMBER = INVO_LST.NUMBER AND HEAD_LST.TAG - 11 = INVO_LST.TAG INNER JOIN CUST_"
        "HESAB ON HEAD_LST.CUST_NO = CUST_HESAB.hes GROUP BY HEAD_LST.NUMBER, HEAD_LST.TA"
        "G, HEAD_LST.DATE_N, HEAD_LST.CUST_NO, HEAD_LST.MOLAH, HEAD_LST.M_NAGHD, HEAD_LST"
        ".MABL_VAR, HEAD_LST.MABL_HAV, HEAD_LST.MABL_HAZ, HEAD_LST.TAKHFIF, HEAD_LST.FNUM"
        "CO, HEAD_LST.USER_NAME, HEAD_LST.MBAA, HEAD_LST.CDDATE, HEAD_LST.CDTIME, HEAD_LS"
        "T.OKDATE, HEAD_LST.OKTIME, HEAD_LST.OKF, CUST_HESAB.NAME, CUST_HESAB.ADDRESS, CU"
        "ST_HESAB.TEL, CUST_HESAB.CODE_E, CUST_HESAB.ECODE, CUST_HESAB.PCODE, CUST_HESAB."
        "IYALAT, CUST_HESAB.CITY, CUST_HESAB.MCODEM, CUST_HESAB.TOZIH, CUST_HESAB.MOBILE,"
        " HEAD_LST.N_S, HEAD_LST.VAS, HEAD_LST.MAS, HEAD_LST.TAH, HEAD_LST.NUMBER1, HEAD_"
        "LST.ANBAR, HEAD_LST.MOIN_VAR, HEAD_LST.MOIN_HAV, HEAD_LST.MOIN_HAZ, HEAD_LST.MOI"
        "N_KHF, HEAD_LST.ANBARF, HEAD_LST.DEPATMAN, HEAD_LST.SHIFT, HEAD_LST.CUST_KIND, H"
        "EAD_LST.SHARAYET, HEAD_LST.HMBAA, HEAD_LST.TAMIR, HEAD_LST.SADER, HEAD_LST.ARZD,"
        " HEAD_LST.ARZKIND, HEAD_LST.MODAT_PPID, HEAD_LST.PEPID, HEAD_LST.PEID, HEAD_LST."
        "SGN1, HEAD_LST.SGN2, HEAD_LST.SGN3, HEAD_LST.SGN4, HEAD_LST.TICMBAA, HEAD_LST.TK"
        "HF, HEAD_LST.JAY HAVING (HEAD_LST.TAG = 12) AND (HEAD_LST.DATE_N BETWEEN 940101 "
        "AND 970101)"
    Caption ="گزارش خريد به تفكيك فاكتور"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Tahoma"
    OnError ="[Event Procedure]"
    AllowFormView =0
    RecordSource ="SELECT HEAD_LST.NUMBER, HEAD_LST.TAG, HEAD_LST.DATE_N, HEAD_LST.CUST_NO, HEAD_LS"
        "T.MOLAH, HEAD_LST.M_NAGHD, HEAD_LST.MABL_VAR, HEAD_LST.MABL_HAV, HEAD_LST.MABL_H"
        "AZ, HEAD_LST.TAKHFIF, HEAD_LST.FNUMCO, HEAD_LST.USER_NAME, HEAD_LST.MBAA, HEAD_L"
        "ST.OKF, HEAD_LST.CDDATE, HEAD_LST.CDTIME, HEAD_LST.OKDATE, HEAD_LST.OKTIME, SUM("
        "INVO_LST.MEGHk) AS MEGHkS, SUM(INVO_LST.MABL_K) AS MABL_KS, CUST_HESAB.NAME, CUS"
        "T_HESAB.ADDRESS, CUST_HESAB.TEL, CUST_HESAB.CODE_E, CUST_HESAB.ECODE, CUST_HESAB"
        ".PCODE, CUST_HESAB.IYALAT, CUST_HESAB.CITY, CUST_HESAB.MCODEM, CUST_HESAB.TOZIH,"
        " CUST_HESAB.MOBILE, HEAD_LST.N_S, HEAD_LST.VAS, HEAD_LST.MAS, HEAD_LST.TAH, HEAD"
        "_LST.NUMBER1, HEAD_LST.ANBAR, HEAD_LST.MOIN_VAR, HEAD_LST.MOIN_HAV, HEAD_LST.MOI"
        "N_HAZ, HEAD_LST.MOIN_KHF, HEAD_LST.ANBARF, HEAD_LST.DEPATMAN, HEAD_LST.SHIFT, HE"
        "AD_LST.CUST_KIND, HEAD_LST.SHARAYET, HEAD_LST.SGN1, HEAD_LST.SGN2, HEAD_LST.SGN3"
        ", HEAD_LST.SGN4, HEAD_LST.HMBAA, HEAD_LST.TAMIR, HEAD_LST.TICMBAA, HEAD_LST.TKHF"
        ", HEAD_LST.SADER, HEAD_LST.ARZD, HEAD_LST.ARZKIND, HEAD_LST.JAY, HEAD_LST.MODAT_"
        "PPID, HEAD_LST.PEPID, HEAD_LST.PEID FROM HEAD_LST INNER JOIN INVO_LST ON HEAD_LS"
        "T.NUMBER = INVO_LST.NUMBER AND HEAD_LST.TAG - 11 = INVO_LST.TAG INNER JOIN CUST_"
        "HESAB ON HEAD_LST.CUST_NO = CUST_HESAB.hes GROUP BY HEAD_LST.NUMBER, HEAD_LST.TA"
        "G, HEAD_LST.DATE_N, HEAD_LST.CUST_NO, HEAD_LST.MOLAH, HEAD_LST.M_NAGHD, HEAD_LST"
        ".MABL_VAR, HEAD_LST.MABL_HAV, HEAD_LST.MABL_HAZ, HEAD_LST.TAKHFIF, HEAD_LST.FNUM"
        "CO, HEAD_LST.USER_NAME, HEAD_LST.MBAA, HEAD_LST.CDDATE, HEAD_LST.CDTIME, HEAD_LS"
        "T.OKDATE, HEAD_LST.OKTIME, HEAD_LST.OKF, CUST_HESAB.NAME, CUST_HESAB.ADDRESS, CU"
        "ST_HESAB.TEL, CUST_HESAB.CODE_E, CUST_HESAB.ECODE, CUST_HESAB.PCODE, CUST_HESAB."
        "IYALAT, CUST_HESAB.CITY, CUST_HESAB.MCODEM, CUST_HESAB.TOZIH, CUST_HESAB.MOBILE,"
        " HEAD_LST.N_S, HEAD_LST.VAS, HEAD_LST.MAS, HEAD_LST.TAH, HEAD_LST.NUMBER1, HEAD_"
        "LST.ANBAR, HEAD_LST.MOIN_VAR, HEAD_LST.MOIN_HAV, HEAD_LST.MOIN_HAZ, HEAD_LST.MOI"
        "N_KHF, HEAD_LST.ANBARF, HEAD_LST.DEPATMAN, HEAD_LST.SHIFT, HEAD_LST.CUST_KIND, H"
        "EAD_LST.SHARAYET, HEAD_LST.HMBAA, HEAD_LST.TAMIR, HEAD_LST.SADER, HEAD_LST.ARZD,"
        " HEAD_LST.ARZKIND, HEAD_LST.MODAT_PPID, HEAD_LST.PEPID, HEAD_LST.PEID, HEAD_LST."
        "SGN1, HEAD_LST.SGN2, HEAD_LST.SGN3, HEAD_LST.SGN4, HEAD_LST.TICMBAA, HEAD_LST.TK"
        "HF, HEAD_LST.JAY HAVING (HEAD_LST.TAG = 12) AND (HEAD_LST.DATE_N BETWEEN 940101 "
        "AND 970101)"
    RecordSourceQualifier ="dbo"
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
        End
        Begin CommandButton
            Width =1701
            Height =283
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Tahoma"
        End
        Begin OptionButton
            SpecialEffect =2
            LabelX =230
            LabelY =-30
        End
        Begin CheckBox
            SpecialEffect =2
            LabelX =230
            LabelY =-30
        End
        Begin OptionGroup
            SpecialEffect =3
            Width =1701
            Height =1701
        End
        Begin BoundObjectFrame
            SpecialEffect =2
            Width =4536
            Height =2835
            LabelX =-1701
        End
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            OldBorderStyle =0
            Width =1701
            LabelX =-1701
            FontName ="Tahoma"
            AsianLineBreak =255
        End
        Begin ListBox
            SpecialEffect =2
            Width =1701
            Height =1417
            LabelX =-1701
            FontName ="Tahoma"
        End
        Begin ComboBox
            SpecialEffect =2
            Width =1701
            LabelX =-1701
            FontName ="Tahoma"
        End
        Begin Section
            Height =22790
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
                    Left =1312
                    Top =113
                    ColumnWidth =1065
                    ColumnOrder =0
                    Name ="NUMBER"
                    ControlSource ="NUMBER"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3867
                            Top =113
                            Width =960
                            Height =240
                            Name ="Label1"
                            Caption ="شماره فاكتور"
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
                    Left =523
                    Top =793
                    Width =2490
                    ColumnWidth =930
                    ColumnOrder =2
                    TabIndex =1
                    Name ="DATE_N"
                    ControlSource ="DATE_N"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4002
                            Top =793
                            Width =825
                            Height =240
                            Name ="Label5"
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
                    Left =523
                    Top =1133
                    Width =2490
                    ColumnWidth =1260
                    ColumnOrder =5
                    TabIndex =2
                    Name ="CUST_NO"
                    ControlSource ="CUST_NO"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3702
                            Top =1133
                            Width =1125
                            Height =240
                            Name ="Label7"
                            Caption ="كد فروشنده"
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
                    Left =523
                    Top =1474
                    Width =2490
                    ColumnWidth =8715
                    ColumnOrder =10
                    TabIndex =3
                    Name ="MOLAH"
                    ControlSource ="MOLAH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4137
                            Top =1474
                            Width =690
                            Height =240
                            Name ="Label9"
                            Caption ="ملاحظات"
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
                    Left =1312
                    Top =1814
                    ColumnWidth =855
                    ColumnOrder =11
                    TabIndex =4
                    Name ="M_NAGHD"
                    ControlSource ="M_NAGHD"
                    Format ="#,###"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4182
                            Top =1814
                            Width =645
                            Height =240
                            Name ="Label11"
                            Caption ="مبلغ نقد"
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
                    Left =1312
                    Top =2154
                    ColumnWidth =975
                    ColumnOrder =12
                    TabIndex =5
                    Name ="MABL_VAR"
                    ControlSource ="MABL_VAR"
                    Format ="#,###"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3957
                            Top =2154
                            Width =870
                            Height =240
                            Name ="Label13"
                            Caption ="مبلغ واريزي"
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
                    Left =1312
                    Top =2494
                    ColumnWidth =885
                    ColumnOrder =13
                    TabIndex =6
                    Name ="MABL_HAV"
                    ControlSource ="MABL_HAV"
                    Format ="#,###"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4047
                            Top =2494
                            Width =780
                            Height =240
                            Name ="Label15"
                            Caption ="مبلغ حواله"
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
                    Left =1312
                    Top =2834
                    ColumnWidth =960
                    ColumnOrder =14
                    TabIndex =7
                    Name ="MABL_HAZ"
                    ControlSource ="MABL_HAZ"
                    Format ="#,###"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4017
                            Top =2834
                            Width =810
                            Height =240
                            Name ="Label17"
                            Caption ="مبلغ هزينه"
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
                    Left =1312
                    Top =3174
                    ColumnWidth =930
                    ColumnOrder =15
                    TabIndex =8
                    Name ="TAKHFIF"
                    ControlSource ="TAKHFIF"
                    Format ="#,###"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4302
                            Top =3174
                            Width =525
                            Height =240
                            Name ="Label19"
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
                    Left =1312
                    Top =3514
                    ColumnWidth =1620
                    ColumnOrder =1
                    TabIndex =9
                    Name ="FNUMCO"
                    ControlSource ="FNUMCO"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3207
                            Top =3514
                            Width =1620
                            Height =240
                            Name ="Label21"
                            Caption ="شماره فاكتور فروشنده"
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
                    Left =523
                    Top =3855
                    Width =2490
                    ColumnOrder =16
                    TabIndex =10
                    Name ="USER_NAME"
                    ControlSource ="USER_NAME"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3852
                            Top =3855
                            Width =975
                            Height =240
                            Name ="Label23"
                            Caption ="كاربر"
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
                    Left =1312
                    Top =4195
                    ColumnWidth =1005
                    ColumnOrder =8
                    TabIndex =11
                    Name ="MBAA"
                    ControlSource ="MBAA"
                    Format ="#,###"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4332
                            Top =4195
                            Width =495
                            Height =240
                            Name ="Label25"
                            Caption ="ارزش افزوده"
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
                    Left =523
                    Top =4875
                    Width =2490
                    ColumnWidth =855
                    ColumnOrder =17
                    TabIndex =12
                    Name ="CDDATE"
                    ControlSource ="CDDATE"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4155
                            Top =4875
                            Width =750
                            Height =240
                            Name ="Label29"
                            Caption ="تاريخ ايجاد"
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
                    Left =1312
                    Top =5215
                    ColumnOrder =18
                    TabIndex =13
                    Name ="CDTIME"
                    ControlSource ="CDTIME"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4185
                            Top =5220
                            Width =900
                            Height =240
                            Name ="Label31"
                            Caption ="ساعت ايجاد"
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
                    Left =523
                    Top =5555
                    Width =2490
                    ColumnOrder =19
                    TabIndex =14
                    Name ="OKDATE"
                    ControlSource ="OKDATE"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4155
                            Top =5550
                            Width =720
                            Height =240
                            Name ="Label33"
                            Caption ="تاريخ چاپ"
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
                    Left =1312
                    Top =5896
                    ColumnOrder =20
                    TabIndex =15
                    Name ="OKTIME"
                    ControlSource ="OKTIME"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4185
                            Top =5895
                            Width =870
                            Height =240
                            Name ="Label35"
                            Caption ="ساعت چاپ"
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
                    Left =1312
                    Top =6236
                    ColumnWidth =1065
                    ColumnOrder =9
                    TabIndex =16
                    Name ="MEGHkS"
                    ControlSource ="MEGHkS"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3867
                            Top =6236
                            Width =960
                            Height =240
                            Name ="Label37"
                            Caption ="مقدار كل كالا"
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
                    Left =1312
                    Top =6576
                    ColumnOrder =7
                    TabIndex =17
                    Name ="MABL_KS"
                    ControlSource ="MABL_KS"
                    Format ="#,###"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4215
                            Top =6570
                            Width =810
                            Height =240
                            Name ="Label39"
                            Caption ="مبلغ فاكتور"
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
                    Left =523
                    Top =6916
                    Width =2490
                    ColumnWidth =4770
                    ColumnOrder =6
                    TabIndex =18
                    Name ="NAME"
                    ControlSource ="NAME"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3642
                            Top =6916
                            Width =1185
                            Height =240
                            Name ="Label41"
                            Caption ="نام  فروشنده"
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
                    Left =523
                    Top =7256
                    Width =2490
                    ColumnWidth =4335
                    ColumnOrder =21
                    TabIndex =19
                    Name ="ADDRESS"
                    ControlSource ="ADDRESS"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4377
                            Top =7256
                            Width =450
                            Height =240
                            Name ="Label43"
                            Caption ="آدرس"
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
                    Left =523
                    Top =7596
                    Width =2490
                    ColumnWidth =1680
                    ColumnOrder =22
                    TabIndex =20
                    Name ="TEL"
                    ControlSource ="TEL"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4437
                            Top =7596
                            Width =390
                            Height =240
                            Name ="Label45"
                            Caption ="تلفن"
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
                    Left =523
                    Top =8277
                    Width =2490
                    ColumnOrder =23
                    TabIndex =21
                    Name ="ECODE"
                    ControlSource ="ECODE"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4245
                            Top =8280
                            Width =855
                            Height =240
                            Name ="Label49"
                            Caption ="كد اقتصادي"
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
                    Left =523
                    Top =8617
                    Width =2490
                    ColumnOrder =24
                    TabIndex =22
                    Name ="PCODE"
                    ControlSource ="PCODE"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4245
                            Top =8610
                            Width =735
                            Height =240
                            Name ="Label51"
                            Caption ="كد پستي"
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
                    Left =523
                    Top =8957
                    Width =2490
                    ColumnOrder =25
                    TabIndex =23
                    Name ="IYALAT"
                    ControlSource ="IYALAT"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4227
                            Top =8957
                            Width =600
                            Height =240
                            Name ="Label53"
                            Caption ="استان"
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
                    Left =523
                    Top =9297
                    Width =2490
                    ColumnOrder =26
                    TabIndex =24
                    Name ="CITY"
                    ControlSource ="CITY"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4407
                            Top =9297
                            Width =420
                            Height =240
                            Name ="Label55"
                            Caption ="شهر"
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
                    Left =523
                    Top =9637
                    Width =2490
                    ColumnOrder =27
                    TabIndex =25
                    Name ="MCODEM"
                    ControlSource ="MCODEM"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4092
                            Top =9637
                            Width =735
                            Height =240
                            Name ="Label57"
                            Caption ="كد ملي"
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
                    Left =523
                    Top =9977
                    Width =2490
                    ColumnOrder =28
                    TabIndex =26
                    Name ="TOZIH"
                    ControlSource ="TOZIH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4287
                            Top =9977
                            Width =540
                            Height =240
                            Name ="Label59"
                            Caption ="توضيح"
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
                    Left =523
                    Top =10318
                    Width =2490
                    ColumnOrder =29
                    TabIndex =27
                    Name ="MOBILE"
                    ControlSource ="MOBILE"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4197
                            Top =10318
                            Width =630
                            Height =240
                            Name ="Label61"
                            Caption ="موبايل"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =736
                    Top =11168
                    ColumnOrder =30
                    TabIndex =28
                    Name ="N_S"
                    ControlSource ="N_S"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3193
                            Top =11168
                            Width =945
                            Height =240
                            Name ="Label62"
                            Caption ="شماره سند"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =736
                    Top =11508
                    ColumnOrder =31
                    TabIndex =29
                    Name ="VAS"
                    ControlSource ="VAS"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3718
                            Top =11508
                            Width =420
                            Height =240
                            Name ="Label63"
                            Caption ="VAS:"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =736
                    Top =11848
                    ColumnOrder =32
                    TabIndex =30
                    Name ="MAS"
                    ControlSource ="MAS"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3690
                            Top =11850
                            Width =1185
                            Height =240
                            Name ="Label64"
                            Caption ="کرايه به عهده"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =736
                    Top =12188
                    ColumnOrder =33
                    TabIndex =31
                    Name ="TAH"
                    ControlSource ="TAH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3073
                            Top =12188
                            Width =1065
                            Height =240
                            Name ="Label65"
                            Caption ="تحويل گيردنده"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =736
                    Top =12529
                    ColumnWidth =1545
                    ColumnOrder =34
                    TabIndex =32
                    Name ="NUMBER1"
                    ControlSource ="NUMBER1"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2548
                            Top =12529
                            Width =1590
                            Height =240
                            Name ="Label66"
                            Caption ="شماره فاكتور برگشت"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =736
                    Top =12869
                    ColumnOrder =35
                    TabIndex =33
                    Name ="ANBAR"
                    ControlSource ="ANBAR"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3748
                            Top =12869
                            Width =390
                            Height =240
                            Name ="Label67"
                            Caption ="انبار"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =736
                    Top =13209
                    ColumnOrder =36
                    TabIndex =34
                    Name ="MOIN_VAR"
                    ControlSource ="MOIN_VAR"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3148
                            Top =13209
                            Width =990
                            Height =240
                            Name ="Label68"
                            Caption ="معين واريزي"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =736
                    Top =13549
                    ColumnOrder =37
                    TabIndex =35
                    Name ="MOIN_HAV"
                    ControlSource ="MOIN_HAV"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3238
                            Top =13549
                            Width =900
                            Height =240
                            Name ="Label69"
                            Caption ="معين حواله"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =736
                    Top =13889
                    ColumnOrder =38
                    TabIndex =36
                    Name ="MOIN_HAZ"
                    ControlSource ="MOIN_HAZ"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3208
                            Top =13889
                            Width =930
                            Height =240
                            Name ="Label70"
                            Caption ="معين هزينه"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =736
                    Top =14229
                    ColumnOrder =39
                    TabIndex =37
                    Name ="MOIN_KHF"
                    ControlSource ="MOIN_KHF"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3148
                            Top =14229
                            Width =990
                            Height =240
                            Name ="Label71"
                            Caption ="معين تخفيف"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =736
                    Top =14570
                    ColumnOrder =40
                    TabIndex =38
                    Name ="ANBARF"
                    ControlSource ="ANBARF"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3283
                            Top =14570
                            Width =855
                            Height =240
                            Name ="Label72"
                            Caption ="انبار فرعي"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =736
                    Top =14910
                    ColumnOrder =41
                    TabIndex =39
                    Name ="DEPATMAN"
                    ControlSource ="DEPATMAN"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3538
                            Top =14910
                            Width =600
                            Height =240
                            Name ="Label73"
                            Caption ="دپاتمان"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =736
                    Top =15250
                    ColumnOrder =42
                    TabIndex =40
                    Name ="SHIFT"
                    ControlSource ="SHIFT"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3538
                            Top =15250
                            Width =600
                            Height =240
                            Name ="Label74"
                            Caption ="شيفت"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =737
                    Top =15647
                    ColumnOrder =43
                    TabIndex =41
                    Name ="CUST_KIND"
                    ControlSource ="CUST_KIND"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3179
                            Top =15647
                            Width =960
                            Height =240
                            Name ="Label75"
                            Caption ="نوع مشتري"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =736
                    Top =15930
                    ColumnOrder =44
                    TabIndex =42
                    Name ="SHARAYET"
                    ControlSource ="SHARAYET"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3223
                            Top =15930
                            Width =915
                            Height =240
                            Name ="Label76"
                            Caption ="شرايط"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =736
                    Top =16270
                    ColumnOrder =45
                    TabIndex =43
                    Name ="SGN1"
                    ControlSource ="SGN1"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3613
                            Top =16270
                            Width =525
                            Height =240
                            Name ="Label77"
                            Caption ="امضا1"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =736
                    Top =16611
                    ColumnOrder =46
                    TabIndex =44
                    Name ="SGN2"
                    ControlSource ="SGN2"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3613
                            Top =16611
                            Width =525
                            Height =240
                            Name ="Label78"
                            Caption ="امضا2"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =736
                    Top =16951
                    ColumnOrder =47
                    TabIndex =45
                    Name ="SGN3"
                    ControlSource ="SGN3"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3613
                            Top =16951
                            Width =525
                            Height =240
                            Name ="Label79"
                            Caption ="امضا3"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =736
                    Top =17291
                    ColumnOrder =48
                    TabIndex =46
                    Name ="SGN4"
                    ControlSource ="SGN4"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3613
                            Top =17291
                            Width =525
                            Height =240
                            Name ="Label80"
                            Caption ="امضا4"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =736
                    Top =17631
                    ColumnOrder =49
                    TabIndex =47
                    Name ="HMBAA"
                    ControlSource ="HMBAA"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3480
                            Top =17625
                            Width =1410
                            Height =240
                            Name ="Label81"
                            Caption ="حساب ارزش افزوده"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =736
                    Top =17971
                    ColumnOrder =50
                    TabIndex =48
                    Name ="TAMIR"
                    ControlSource ="TAMIR"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3525
                            Top =17970
                            Width =675
                            Height =240
                            Name ="Label82"
                            Caption ="تعمير کار"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =736
                    Top =18311
                    ColumnOrder =51
                    TabIndex =49
                    Name ="TICMBAA"
                    ControlSource ="TICMBAA"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3330
                            Top =18315
                            Width =1215
                            Height =240
                            Name ="Label83"
                            Caption ="تيک ارزش افزوده"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =736
                    Top =18651
                    ColumnOrder =52
                    TabIndex =50
                    Name ="TKHF"
                    ControlSource ="TKHF"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3630
                            Top =18645
                            Width =945
                            Height =240
                            Name ="Label84"
                            Caption ="%ت.ن"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =736
                    Top =18992
                    ColumnOrder =53
                    TabIndex =51
                    Name ="SADER"
                    ControlSource ="SADER"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3508
                            Top =18992
                            Width =630
                            Height =240
                            Name ="Label85"
                            Caption ="صادرات"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =736
                    Top =19332
                    ColumnOrder =54
                    TabIndex =52
                    Name ="ARZD"
                    ControlSource ="ARZD"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3600
                            Top =19335
                            Width =675
                            Height =240
                            Name ="Label86"
                            Caption ="قيمت ارز"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =736
                    Top =19672
                    ColumnOrder =55
                    TabIndex =53
                    Name ="ARZKIND"
                    ControlSource ="ARZKIND"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3343
                            Top =19672
                            Width =795
                            Height =240
                            Name ="Label87"
                            Caption ="نوع ارز"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =736
                    Top =20012
                    ColumnOrder =56
                    TabIndex =54
                    Name ="JAY"
                    ControlSource ="JAY"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3735
                            Top =20010
                            Width =735
                            Height =240
                            Name ="Label88"
                            Caption ="تيک جايزه"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =736
                    Top =20352
                    ColumnOrder =57
                    TabIndex =55
                    Name ="MODAT_PPID"
                    ControlSource ="MODAT_PPID"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3028
                            Top =20352
                            Width =1110
                            Height =240
                            Name ="Label89"
                            Caption ="نحوه پرداخت"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =736
                    Top =20692
                    ColumnOrder =58
                    TabIndex =56
                    Name ="PEPID"
                    ControlSource ="PEPID"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3570
                            Top =20685
                            Width =915
                            Height =240
                            Name ="Label90"
                            Caption ="الگوي قيمت"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =736
                    Top =21033
                    ColumnOrder =59
                    TabIndex =57
                    Name ="PEID"
                    ControlSource ="PEID"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3660
                            Top =21030
                            Width =975
                            Height =240
                            Name ="Label91"
                            Caption ="الگوي تخفيف"
                        End
                    End
                End
                Begin TextBox
                    FELineBreak = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =178
                    TextAlign =3
                    IMESentenceMode =3
                    Left =850
                    Top =21543
                    ColumnWidth =1020
                    ColumnOrder =3
                    TabIndex =58
                    Name ="date_h"
                    ControlSource ="date_h"
                    InputMask ="####/##/##"
                    AsianLineBreak =0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =178
                            Left =3600
                            Top =21540
                            Width =810
                            Height =240
                            Name ="Label46"
                            Caption ="تاريخ رسيد"
                        End
                    End
                End
                Begin TextBox
                    FELineBreak = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =178
                    TextAlign =3
                    IMESentenceMode =3
                    Left =850
                    Top =21903
                    ColumnWidth =930
                    ColumnOrder =4
                    TabIndex =59
                    Name ="Text47"
                    ControlSource ="ddf"
                    AsianLineBreak =0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =178
                            Left =3601
                            Top =21908
                            Width =825
                            Height =240
                            Name ="Label48"
                            Caption ="تفاوت تاريخ"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "LIST_KHRID_KOL_SUB.cls"
