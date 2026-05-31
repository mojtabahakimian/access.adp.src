Version =20
VersionRequired =20
Begin Form
    AllowDeletions = NotDefault
    AllowAdditions = NotDefault
    KeyPreview = NotDefault
    OrderByOn = NotDefault
    AllowEdits = NotDefault
    DefaultView =2
    AllowUpdating =4
    ViewsAllowed =2
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =5100
    RowHeight =300
    ItemSuffix =48
    Left =1695
    Top =2550
    Right =14160
    Bottom =7770
    DatasheetGridlinesColor =12632256
    MaxRecords =100000
    RecSrcDt = Begin
        0xd71a5e085108e640
    End
    RecordSource ="SELECT PAY_GETD.N_SERI, PAY_GETD.BANK, PAY_GETD.DATE_S, PAY_GETD.DATE, PAY_GETD."
        "SHOBEH, PAY_GETD.MABL, PAY_GETD.NAME_TAH, PAY_GETD.N_HESAB, PAY_GETD.N_S, TCOD_B"
        "ANKS.NAMES, PAY_GETD.RADIF, PAY_GETD.N_KOL, PAY_GETD.N_MOIN, PAY_GETD.N_KOL2, PA"
        "Y_GETD.N_MOIN2, PAY_GETD.N_KOL3, PAY_GETD.N_MOIN3, COD_HESAB.BANK AS BKK, PAY_GE"
        "TD.N_TAF, PAY_GETD.N_TAF2, PAY_GETD.N_TAF3, PAY_GETD.VAZ, CUST_HESAB.NAME AS Exp"
        "r1, PAY_GETD.SANDUGH, ISNULL(HEAD_LST.CUST_KIND, ISNULL(PGET_HED.CUST_KIND, NULL"
        ")) AS CUST_KIND, ISNULL(HEAD_LST.USER_NAME, ISNULL(PGET_HED.USER_NAME, NULL)) AS"
        " USER_NAME, ISNULL(HEAD_LST.DEPATMAN, ISNULL(PGET_HED.DEPATMAN, NULL)) AS DEPATM"
        "AN, ISNULL(HEAD_LST.SHIFT, ISNULL(PGET_HED.SHIFT, NULL)) AS SHIFT, PAY_GETD.KIND"
        ", PAY_GETD.HES1, PAY_GETD.HES2, PAY_GETD.HES3, PAY_GETD.ESTELAM, PAY_GETD.CRT, P"
        "AY_GETD.UID, PAY_GETD.SAYADI FROM TCOD_BANKS INNER JOIN PAY_GETD LEFT OUTER JOIN"
        " COD_HESAB ON PAY_GETD.N_TAF = COD_HESAB.MOIN ON TCOD_BANKS.CODE = PAY_GETD.BANK"
        " LEFT OUTER JOIN PGET_HED INNER JOIN PGET_LST ON PGET_HED.ID = PGET_LST.ID AND P"
        "GET_HED.DATE = PGET_LST.DATE ON PAY_GETD.N_SERI = PGET_LST.N_SERI AND PAY_GETD.B"
        "ANK = PGET_LST.BANK AND PAY_GETD.MABL = PGET_LST.MABL LEFT OUTER JOIN HEAD_LST O"
        "N PAY_GETD.NUMBER = HEAD_LST.NUMBER AND PAY_GETD.TAG = HEAD_LST.TAG LEFT OUTER J"
        "OIN CUST_HESAB ON RTRIM(CAST(PAY_GETD.N_KOL AS nvarchar)) + '-' + RTRIM(CAST(PAY"
        "_GETD.N_MOIN AS nvarchar)) + '-' + RTRIM(CAST(PAY_GETD.N_TAF AS nvarchar)) = CUS"
        "T_HESAB.hes WHERE (PAY_GETD.N_KOL2 IS NULL) AND (PAY_GETD.N_KOL3 IS NULL) AND (P"
        "AY_GETD.N_KOL = 112)"
    Caption ="ليست چكهاي دريافتي موجود در بانك و صندوق وصول نشده"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    HelpFile ="dena.hlp"
    DatasheetFontName ="Tahoma"
    OnKeyPress ="[Event Procedure]"
    OnLoad ="[Event Procedure]"
    OnError ="[Event Procedure]"
    AllowFormView =0
    RecordSourceQualifier ="dbo"
    Begin
        Begin Label
            BackStyle =0
            TextFontCharSet =178
            FontName ="Tahoma"
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
            SpecialEffect =2
            OldBorderStyle =0
            TextFontCharSet =178
            Width =1701
            LabelX =-1701
            FontName ="Tahoma"
        End
        Begin ListBox
            SpecialEffect =2
            TextFontCharSet =178
            Width =1701
            Height =1417
            LabelX =-1701
            FontName ="Tahoma"
        End
        Begin ComboBox
            SpecialEffect =2
            TextFontCharSet =178
            Width =1701
            LabelX =-1701
            FontName ="Tahoma"
        End
        Begin PageHeader
            DisplayWhen =1
            Height =0
            BackColor =-2147483633
            Name ="PageHeader"
        End
        Begin Section
            Height =10680
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =93
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =1
                    NumeralShapes =1
                    Left =1233
                    Top =113
                    Height =255
                    ColumnWidth =1095
                    ColumnOrder =1
                    Name ="N_SERI"
                    ControlSource ="N_SERI"
                    Format ="Fixed"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =3800
                            Top =113
                            Width =1005
                            Height =240
                            Name ="Label1"
                            Caption ="شماره سريال"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =1
                    NumeralShapes =1
                    Left =1233
                    Top =793
                    Height =255
                    ColumnWidth =1125
                    ColumnOrder =2
                    TabIndex =1
                    Name ="DATE_S"
                    ControlSource ="DATE_S"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3755
                            Top =793
                            Width =1050
                            Height =240
                            Name ="Label5"
                            Caption ="تاريخ سررسيد"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =1
                    NumeralShapes =1
                    Left =1233
                    Top =1133
                    Height =255
                    ColumnWidth =1110
                    ColumnOrder =3
                    TabIndex =2
                    Name ="DATE"
                    ControlSource ="DATE"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3920
                            Top =1133
                            Width =885
                            Height =240
                            Name ="Label7"
                            Caption ="تاريخ دريافت"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =1
                    NumeralShapes =1
                    Left =1224
                    Top =1474
                    Width =1710
                    Height =255
                    ColumnWidth =2025
                    ColumnOrder =5
                    TabIndex =3
                    Name ="SHOBEH"
                    ControlSource ="SHOBEH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4085
                            Top =1474
                            Width =720
                            Height =240
                            Name ="Label9"
                            Caption ="نام شعبه"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =1
                    NumeralShapes =1
                    Left =1233
                    Top =1814
                    Height =255
                    ColumnWidth =1845
                    ColumnOrder =6
                    TabIndex =4
                    Name ="MABL"
                    ControlSource ="MABL"
                    Format ="Standard"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4430
                            Top =1814
                            Width =375
                            Height =240
                            Name ="Label11"
                            Caption ="مبلغ"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =1
                    NumeralShapes =1
                    Left =354
                    Top =2154
                    Width =2580
                    Height =255
                    ColumnWidth =2925
                    ColumnOrder =7
                    TabIndex =5
                    Name ="NAME_TAH"
                    ControlSource ="NAME_TAH"
                    OnDblClick ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3630
                            Top =2160
                            Width =1260
                            Height =240
                            Name ="Label13"
                            Caption ="*نام  پرداخت كننده"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =1
                    NumeralShapes =1
                    Left =1233
                    Top =2494
                    Height =255
                    ColumnOrder =10
                    TabIndex =6
                    Name ="N_HESAB"
                    ControlSource ="N_HESAB"
                    Format ="Fixed"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3770
                            Top =2494
                            Width =1035
                            Height =240
                            Name ="Label15"
                            Caption ="شماره حساب"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =1
                    NumeralShapes =1
                    Left =1233
                    Top =2834
                    Height =255
                    ColumnWidth =885
                    ColumnOrder =11
                    TabIndex =7
                    Name ="N_S"
                    ControlSource ="N_S"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3920
                            Top =2834
                            Width =885
                            Height =240
                            Name ="Label17"
                            Caption ="شماره سند"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =1
                    NumeralShapes =1
                    Left =144
                    Top =3174
                    Width =2790
                    Height =255
                    ColumnWidth =870
                    ColumnOrder =4
                    TabIndex =8
                    Name ="NAMES"
                    ControlSource ="NAMES"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4190
                            Top =3174
                            Width =615
                            Height =240
                            Name ="Label19"
                            Caption ="نام بانك"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =93
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =1
                    NumeralShapes =1
                    Left =1245
                    Top =453
                    Height =255
                    ColumnWidth =645
                    ColumnOrder =0
                    TabIndex =9
                    Name ="RADIF"
                    ControlSource ="RADIF"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4177
                            Top =450
                            Width =630
                            Height =240
                            Name ="Label20"
                            Caption ="شماره "
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =1
                    NumeralShapes =1
                    Left =400
                    Top =3549
                    ColumnOrder =12
                    TabIndex =10
                    Name ="N_KOL"
                    ControlSource ="N_KOL"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2965
                            Top =3549
                            Width =1005
                            Height =240
                            Name ="Label22"
                            Caption =" به حساب كل"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =1
                    NumeralShapes =1
                    Left =400
                    Top =3889
                    ColumnOrder =13
                    TabIndex =11
                    Name ="N_MOIN"
                    ControlSource ="N_MOIN"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2800
                            Top =3894
                            Width =1125
                            Height =240
                            Name ="Label23"
                            Caption ="به حساب معين"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =1
                    NumeralShapes =1
                    Left =400
                    Top =4229
                    ColumnWidth =1545
                    ColumnOrder =15
                    TabIndex =12
                    Name ="N_KOL2"
                    ControlSource ="N_KOL2"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =2530
                            Top =4224
                            Width =1665
                            Height =405
                            Name ="Label24"
                            Caption ="برگشت به حساب كل"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =1
                    NumeralShapes =1
                    Left =400
                    Top =4569
                    ColumnWidth =1710
                    ColumnOrder =16
                    TabIndex =13
                    Name ="N_MOIN2"
                    ControlSource ="N_MOIN2"
                    Begin
                        Begin Label
                            OverlapFlags =223
                            Left =2665
                            Top =4569
                            Width =1545
                            Height =405
                            Name ="Label25"
                            Caption ="برگشت به حساب معين"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =1
                    NumeralShapes =1
                    Left =400
                    Top =4909
                    ColumnWidth =1725
                    ColumnOrder =18
                    TabIndex =14
                    Name ="N_KOL3"
                    ControlSource ="N_KOL3"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =2830
                            Top =4914
                            Width =1395
                            Height =240
                            Name ="Label26"
                            Caption ="وصول به حساب كل"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =1
                    NumeralShapes =1
                    Left =400
                    Top =5249
                    ColumnWidth =1680
                    ColumnOrder =19
                    TabIndex =15
                    Name ="N_MOIN3"
                    ControlSource ="N_MOIN3"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2665
                            Top =5244
                            Width =1560
                            Height =240
                            Name ="Label27"
                            Caption ="وصول به حساب معين"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    ReadingOrder =2
                    ScrollBarAlign =1
                    NumeralShapes =1
                    Left =144
                    Top =5809
                    ColumnWidth =1860
                    ColumnOrder =17
                    TabIndex =16
                    Name ="N_TAF2"
                    ControlSource ="N_TAF2"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2353
                            Top =5814
                            Width =1905
                            Height =240
                            Name ="Label35"
                            Caption ="برگشت به حساب تفصيلي"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    ReadingOrder =2
                    ScrollBarAlign =1
                    NumeralShapes =1
                    Left =144
                    Top =6150
                    ColumnWidth =1830
                    ColumnOrder =20
                    TabIndex =17
                    Name ="N_TAF3"
                    ControlSource ="N_TAF3"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2353
                            Top =6144
                            Width =1755
                            Height =240
                            Name ="Label36"
                            Caption ="وصول به حساب تفصيلي"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    ReadingOrder =2
                    ScrollBarAlign =1
                    NumeralShapes =1
                    Left =144
                    Top =6490
                    ColumnWidth =1335
                    ColumnOrder =14
                    TabIndex =18
                    Name ="N_TAF"
                    ControlSource ="N_TAF"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2353
                            Top =6489
                            Width =1320
                            Height =240
                            Name ="Label37"
                            Caption ="به حساب تفصيلي"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =1
                    NumeralShapes =1
                    Left =1077
                    Top =5499
                    Height =255
                    ColumnWidth =690
                    ColumnOrder =21
                    TabIndex =19
                    Name ="VAZ"
                    ControlSource ="VAZ"
                    Format ="Fixed"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3644
                            Top =5499
                            Width =1005
                            Height =240
                            Name ="Label29"
                            Caption ="صندوق"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OldBorderStyle =0
                    OverlapFlags =255
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =1
                    NumeralShapes =1
                    IMESentenceMode =3
                    ColumnCount =2
                    Width =1581
                    Height =300
                    ColumnOrder =22
                    TabIndex =20
                    Name ="KIND"
                    ControlSource ="KIND"
                    RowSourceType ="Value List"
                    RowSource ="1;\"تجاري\";0;\"غير تجاري\""
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            ReadingOrder =2
                            NumeralShapes =1
                            Width =1005
                            Height =240
                            Name ="Label30"
                            Caption ="نوع"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =1
                    NumeralShapes =1
                    IMESentenceMode =3
                    Left =480
                    Top =7050
                    ColumnOrder =23
                    TabIndex =21
                    Name ="CUST_KIND"
                    ControlSource ="CUST_KIND"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2922
                            Top =7050
                            Width =960
                            Height =240
                            Name ="Label32"
                            Caption ="نوع مشتري"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =1
                    NumeralShapes =1
                    IMESentenceMode =3
                    Left =480
                    Top =7338
                    ColumnOrder =24
                    TabIndex =22
                    Name ="USER_NAME"
                    ControlSource ="USER_NAME"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2847
                            Top =7338
                            Width =1035
                            Height =240
                            Name ="Label33"
                            Caption ="نام كاربر"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =1
                    NumeralShapes =1
                    IMESentenceMode =3
                    Left =480
                    Top =7626
                    ColumnOrder =25
                    TabIndex =23
                    Name ="DEPATMAN"
                    ControlSource ="DEPATMAN"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3282
                            Top =7626
                            Width =600
                            Height =240
                            Name ="Label34"
                            Caption ="دپاتمان"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =1
                    NumeralShapes =1
                    IMESentenceMode =3
                    Left =480
                    Top =7914
                    ColumnOrder =26
                    TabIndex =24
                    Name ="SHIFT"
                    ControlSource ="SHIFT"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3282
                            Top =7914
                            Width =600
                            Height =240
                            Name ="Label38"
                            Caption ="شيفت"
                        End
                    End
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =1
                    NumeralShapes =1
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =30
                    Top =8790
                    ColumnWidth =3840
                    ColumnOrder =27
                    TabIndex =25
                    Name ="bhes"
                    ControlSource ="HES1"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT hes, hes + N' : ' + NAME AS Expr1 FROM CUST_HESAB"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2595
                            Top =8790
                            Width =1005
                            Height =240
                            Name ="Label39"
                            Caption =" به حساب "
                        End
                    End
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =1
                    NumeralShapes =1
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =30
                    Top =9470
                    ColumnWidth =3855
                    ColumnOrder =28
                    TabIndex =26
                    Name ="brhes"
                    ControlSource ="HES2"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT hes, hes + N' : ' + NAME AS Expr1 FROM CUST_HESAB"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2160
                            Top =9465
                            Width =1665
                            Height =405
                            Name ="Label41"
                            Caption ="برگشت به حساب "
                        End
                    End
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =1
                    NumeralShapes =1
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =30
                    Top =10150
                    ColumnWidth =3525
                    ColumnOrder =29
                    TabIndex =27
                    Name ="vshes"
                    ControlSource ="HES3"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT hes, hes + N' : ' + NAME AS Expr1 FROM CUST_HESAB"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2460
                            Top =10155
                            Width =1395
                            Height =240
                            Name ="Label43"
                            Caption ="وصول به حساب "
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =247
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =1
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1133
                    Top =3571
                    ColumnOrder =30
                    TabIndex =28
                    Name ="BANK"
                    ControlSource ="BANK"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4130
                            Top =3571
                            Width =405
                            Height =240
                            Name ="Label44"
                            Caption ="بانك"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    ReadingOrder =2
                    ScrollBarAlign =1
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =566
                    Top =8277
                    ColumnOrder =31
                    TabIndex =29
                    Name ="SANDUGH"
                    ControlSource ="SANDUGH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3113
                            Top =8277
                            Width =855
                            Height =240
                            Name ="Label45"
                            Caption ="صندوق"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =255
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Width =3111
                    Height =315
                    ColumnWidth =2595
                    ColumnOrder =32
                    TabIndex =30
                    Name ="SAYADI"
                    ControlSource ="SAYADI"
                    Begin
                        Begin Label
                            OverlapFlags =223
                            TextFontCharSet =0
                            Left =3220
                            Top =2
                            Width =1095
                            Height =240
                            Name ="Label78"
                            Caption ="شماره صيادي"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OldBorderStyle =0
                    OverlapFlags =247
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Width =2211
                    Height =315
                    ColumnOrder =8
                    TabIndex =31
                    Name ="Combo46"
                    ControlSource ="VAZ"
                    RowSourceType ="Value List"
                    RowSource ="1;\"نزد صندوق\";2;\"نزد بانك\";3;\"وصول شده\";4;\"واگذار شده\";5;\"برگشت شده\";6"
                        ";\"مسترد شده\";7;\"حذف شده\""
                    ColumnWidths ="0"
                    DefaultValue ="1"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextFontCharSet =0
                            Left =2342
                            Width =1005
                            Height =315
                            Name ="Label50"
                            Caption ="وضعيت چك*"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =247
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Top =435
                    Width =2211
                    Height =315
                    ColumnOrder =9
                    TabIndex =32
                    Name ="Combo47"
                    ControlSource ="SANDUGH"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TNUMBER, NAME FROM TDETA_HES WHERE (N_KOL = 113) AND (NUMBER = 1)"
                    ColumnWidths ="0"
                    DefaultValue ="1"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextFontCharSet =0
                            Left =2342
                            Top =435
                            Width =1005
                            Height =315
                            Name ="Label40"
                            Caption ="موقعيت چك*"
                        End
                    End
                End
            End
        End
        Begin PageFooter
            DisplayWhen =1
            Height =1134
            BackColor =-2147483633
            Name ="PageFooter"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =1
                    NumeralShapes =1
                    Name ="smabl"
                    ControlSource ="=Sum([MABL])"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3197
                            Width =375
                            Height =240
                            Name ="Label21"
                            Caption ="مبلغ"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "CHEK_VLISTALL.cls"
