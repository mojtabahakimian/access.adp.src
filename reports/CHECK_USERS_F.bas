Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    AllowUpdating =4
    TabularCharSet =178
    TabularFamily =48
    DateGrouping =1
    Orientation =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =64
    GridY =64
    Width =10836
    DatasheetFontHeight =10
    ItemSuffix =41
    Left =330
    Top =7140
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xf7815b18fa66e440
    End
    RecordSource ="SELECT PAY_GETD.*, CUST_HESAB.NAME AS CUSTNAM, TCOD_BANKS.NAMES + N'-' + ISNULL("
        "PAY_GETD.SHOBEH, N' ') AS BANKNAME FROM HEAD_LST INNER JOIN PAY_GETD ON HEAD_LST"
        ".TAG = PAY_GETD.TAG AND HEAD_LST.NUMBER = PAY_GETD.NUMBER INNER JOIN CUST_HESAB "
        "ON HEAD_LST.CUST_NO = CUST_HESAB.hes INNER JOIN TCOD_BANKS ON PAY_GETD.BANK = TC"
        "OD_BANKS.CODE WHERE (HEAD_LST.TAG = 2 OR HEAD_LST.TAG = 14) AND (HEAD_LST.DATE_N"
        " BETWEEN 1 AND 910202) AND (HEAD_LST.DEPATMAN = 1) AND (CAST(HEAD_LST.SHIFT AS N"
        "VARCHAR) LIKE N'%') AND (HEAD_LST.USER_NAME LIKE N'%')"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial"
    Begin
        Begin Label
            BackStyle =0
            TextFontFamily =2
            FontName ="Arial"
        End
        Begin Line
            Width =1701
        End
        Begin TextBox
            FELineBreak = NotDefault
            OldBorderStyle =0
            TextFontFamily =2
            FontName ="Arial"
            AsianLineBreak =255
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =975
            Name ="ReportHeader"
            Begin
                Begin Label
                    TextFontCharSet =178
                    TextAlign =3
                    ReadingOrder =2
                    NumeralShapes =2
                    Left =10320
                    Top =495
                    Width =405
                    Height =330
                    FontWeight =700
                    Name ="Label245"
                    Caption ="رديف"
                    FontName ="Titr"
                End
                Begin Label
                    TextFontCharSet =178
                    TextAlign =2
                    ReadingOrder =2
                    NumeralShapes =2
                    Left =9435
                    Top =510
                    Width =870
                    Height =315
                    FontWeight =700
                    Name ="Label246"
                    Caption ="شماره فاكتور"
                    FontName ="Titr"
                End
                Begin Label
                    TextFontCharSet =178
                    TextAlign =2
                    ReadingOrder =2
                    NumeralShapes =2
                    Left =7440
                    Top =510
                    Width =1995
                    Height =315
                    FontWeight =700
                    Name ="Label247"
                    Caption ="نام مشتري"
                    FontName ="Titr"
                End
                Begin Label
                    TextFontCharSet =178
                    TextAlign =2
                    ReadingOrder =2
                    NumeralShapes =2
                    Left =6090
                    Top =510
                    Width =1335
                    Height =315
                    FontWeight =700
                    Name ="Label248"
                    Caption ="سريال"
                    FontName ="Titr"
                End
                Begin Label
                    TextFontCharSet =178
                    TextAlign =2
                    ReadingOrder =2
                    NumeralShapes =2
                    Left =4185
                    Top =513
                    Width =1875
                    Height =285
                    FontWeight =700
                    Name ="Label249"
                    Caption ="نام بانك"
                    FontName ="Titr"
                End
                Begin Label
                    TextFontCharSet =178
                    TextAlign =2
                    ReadingOrder =2
                    NumeralShapes =2
                    Left =3195
                    Top =510
                    Width =945
                    Height =315
                    FontWeight =700
                    Name ="Label252"
                    Caption ="تاريخ سررسيد"
                    FontName ="Titr"
                End
                Begin Label
                    TextFontCharSet =178
                    TextAlign =2
                    ReadingOrder =2
                    NumeralShapes =2
                    Left =1515
                    Top =510
                    Width =1695
                    Height =315
                    FontWeight =700
                    Name ="Label253"
                    Caption ="شماره حساب"
                    FontName ="Titr"
                End
                Begin Label
                    TextFontCharSet =178
                    TextAlign =2
                    ReadingOrder =2
                    NumeralShapes =2
                    Top =510
                    Width =1455
                    Height =315
                    FontWeight =700
                    Name ="Label254"
                    Caption ="مبلغ"
                    FontName ="Titr"
                End
                Begin Label
                    TextFontCharSet =178
                    TextAlign =2
                    Left =4275
                    Width =2355
                    Height =465
                    FontSize =12
                    FontWeight =700
                    ForeColor =255
                    Name ="LIST_PISH_ANBARS_DTL Label"
                    Caption ="چكهاي دريافت  شده"
                    FontName ="Titr"
                    EventProcPrefix ="LIST_PISH_ANBARS_DTL_Label"
                End
                Begin Line
                    BorderWidth =3
                    Top =480
                    Width =10836
                    Name ="Line99"
                End
                Begin Line
                    BorderWidth =3
                    Left =10815
                    Top =465
                    Width =0
                    Height =510
                    Name ="Line130"
                End
                Begin Line
                    BorderWidth =3
                    Top =975
                    Width =10836
                    Name ="Line147"
                End
                Begin Line
                    BorderWidth =3
                    Top =465
                    Width =0
                    Height =510
                    Name ="Line19"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =10328
                    Top =495
                    Width =0
                    Height =457
                    Name ="Line34"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =9480
                    Top =495
                    Width =0
                    Height =457
                    Name ="Line35"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =7410
                    Top =495
                    Width =0
                    Height =457
                    Name ="Line36"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =6078
                    Top =495
                    Width =0
                    Height =457
                    Name ="Line37"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =4125
                    Top =495
                    Width =0
                    Height =457
                    Name ="Line38"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =3180
                    Top =495
                    Width =0
                    Height =457
                    Name ="Line39"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =1440
                    Top =495
                    Width =0
                    Height =457
                    Name ="Line40"
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =397
            Name ="Detail"
            Begin
                Begin TextBox
                    TextFontCharSet =178
                    TextAlign =2
                    IMESentenceMode =3
                    Left =9495
                    Top =15
                    Width =825
                    Height =330
                    FontWeight =700
                    Name ="NUMBER"
                    ControlSource ="NUMBER"
                    FontName ="Titr"
                End
                Begin TextBox
                    TextFontCharSet =178
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6105
                    Top =15
                    Width =1305
                    Height =330
                    FontWeight =700
                    TabIndex =1
                    Name ="N_SERI"
                    ControlSource ="N_SERI"
                    FontName ="Titr"
                End
                Begin TextBox
                    TextFontCharSet =178
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4188
                    Top =15
                    Width =1875
                    Height =330
                    FontWeight =700
                    TabIndex =2
                    Name ="BANK"
                    ControlSource ="BANKNAME"
                    FontName ="Titr"
                End
                Begin TextBox
                    TextFontCharSet =178
                    TextAlign =3
                    IMESentenceMode =3
                    Left =3180
                    Top =15
                    Width =930
                    Height =330
                    FontWeight =700
                    TabIndex =3
                    Name ="DATE_S"
                    ControlSource ="DATE_S"
                    FontName ="Titr"
                    InputMask ="####/##/##"
                End
                Begin TextBox
                    TextFontCharSet =178
                    TextAlign =3
                    IMESentenceMode =3
                    Top =15
                    Height =330
                    FontWeight =700
                    TabIndex =4
                    Name ="MABL"
                    ControlSource ="MABL"
                    Format ="#,###"
                    FontName ="Titr"
                End
                Begin TextBox
                    TextFontCharSet =178
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1515
                    Top =15
                    Width =1650
                    Height =330
                    FontWeight =700
                    TabIndex =5
                    Name ="N_HESAB"
                    ControlSource ="N_HESAB"
                    FontName ="Titr"
                End
                Begin TextBox
                    TextFontCharSet =178
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7425
                    Top =15
                    Width =2025
                    Height =330
                    FontWeight =700
                    TabIndex =6
                    Name ="CUSTNAM"
                    ControlSource ="CUSTNAM"
                    FontName ="Titr"
                End
                Begin TextBox
                    RunningSum =2
                    TextFontCharSet =178
                    TextAlign =2
                    IMESentenceMode =3
                    Left =10335
                    Top =15
                    Width =465
                    Height =345
                    FontWeight =700
                    TabIndex =7
                    Name ="Text11"
                    ControlSource ="=1"
                    FontName ="Titr"
                End
                Begin Line
                    Left =10328
                    Width =0
                    Height =397
                    Name ="Line12"
                End
                Begin Line
                    Left =9480
                    Width =0
                    Height =397
                    Name ="Line13"
                End
                Begin Line
                    Left =7410
                    Width =0
                    Height =397
                    Name ="Line14"
                End
                Begin Line
                    Left =6078
                    Width =0
                    Height =397
                    Name ="Line15"
                End
                Begin Line
                    Left =4125
                    Width =0
                    Height =397
                    Name ="Line16"
                End
                Begin Line
                    Left =3180
                    Width =0
                    Height =397
                    Name ="Line17"
                End
                Begin Line
                    Left =1440
                    Width =0
                    Height =397
                    Name ="Line18"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Top =375
                    Width =10821
                    Name ="Line100"
                End
                Begin Line
                    BorderWidth =3
                    Left =10815
                    Width =0
                    Height =397
                    Name ="Line101"
                End
                Begin Line
                    BorderWidth =2
                    Width =0
                    Height =397
                    Name ="Line109"
                End
            End
        End
        Begin FormFooter
            KeepTogether = NotDefault
            Height =465
            Name ="ReportFooter"
            Begin
                Begin Line
                    BorderWidth =3
                    Top =30
                    Width =10780
                    Name ="Line110"
                End
                Begin TextBox
                    FELineBreak = NotDefault
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    BackStyle =0
                    NumeralShapes =2
                    Left =26
                    Width =1884
                    Height =405
                    FontSize =10
                    FontWeight =700
                    BorderColor =32768
                    Name ="GHABELG"
                    ControlSource ="=Sum([MABL])"
                    Format ="#,###"
                    FontName ="Traffic"
                    AsianLineBreak =0
                End
                Begin Line
                    BorderWidth =3
                    Top =465
                    Width =10795
                    Name ="Line176"
                End
                Begin Line
                    BorderWidth =3
                    Width =0
                    Height =450
                    Name ="Line177"
                End
                Begin Line
                    BorderWidth =3
                    Left =10815
                    Width =0
                    Height =450
                    Name ="Line178"
                End
                Begin Line
                    BorderWidth =3
                    Left =2025
                    Width =0
                    Height =450
                    Name ="Line180"
                End
                Begin Label
                    TextFontCharSet =178
                    TextAlign =2
                    TextFontFamily =10
                    Left =4785
                    Width =1575
                    Height =405
                    FontSize =10
                    FontWeight =700
                    BorderColor =32768
                    Name ="Label181"
                    Caption ="جمع مبالغ "
                    FontName ="Traffic"
                End
            End
        End
    End
End
CodeBehindForm
' See "CHECK_USERS_F.cls"
