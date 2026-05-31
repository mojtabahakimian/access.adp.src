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
    ItemSuffix =38
    Left =330
    Top =5940
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x9512e5147026e440
    End
    RecordSource ="SELECT TOP 100 PERCENT HEAD_LST.NUMBER, 0 AS SMAB, HEAD_LST.TAKHFIF, HEAD_LST.MB"
        "AA, HEAD_LST.MABL_VAR, HEAD_LST.MABL_HAV, HEAD_LST.M_NAGHD, 0 AS HAZ, 0 AS CHK, "
        "CUST_HESAB.NAME, 0 AS MAND FROM HEAD_LST INNER JOIN CUST_HESAB ON HEAD_LST.CUST_"
        "NO = CUST_HESAB.hes WHERE (HEAD_LST.TAG = 2) AND (0 > 0) ORDER BY HEAD_LST.NUMBE"
        "R"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial"
    RecordSourceQualifier ="dbo"
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
            Height =945
            Name ="ReportHeader"
            Begin
                Begin Label
                    TextFontCharSet =178
                    TextAlign =2
                    Left =4260
                    Width =2385
                    Height =465
                    FontSize =12
                    FontWeight =700
                    ForeColor =255
                    Name ="LIST_PISH_ANBARS_DTL Label"
                    Caption ="ليست بدهكاران"
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
                    Height =480
                    Name ="Line130"
                End
                Begin Line
                    BorderWidth =3
                    Top =915
                    Width =10836
                    Name ="Line147"
                End
                Begin Line
                    BorderWidth =3
                    Top =465
                    Width =0
                    Height =480
                    Name ="Line19"
                End
                Begin Line
                    Left =10403
                    Top =510
                    Width =0
                    Height =383
                    Name ="Line22"
                End
                Begin Line
                    Left =9540
                    Top =510
                    Width =0
                    Height =383
                    Name ="Line23"
                End
                Begin Line
                    Left =5926
                    Top =510
                    Width =0
                    Height =383
                    Name ="Line24"
                End
                Begin Line
                    Left =4740
                    Top =510
                    Width =0
                    Height =383
                    Name ="Line25"
                End
                Begin Line
                    Left =3555
                    Top =510
                    Width =0
                    Height =383
                    Name ="Line26"
                End
                Begin Line
                    Left =2385
                    Top =510
                    Width =0
                    Height =383
                    Name ="Line27"
                End
                Begin Line
                    Left =1201
                    Top =510
                    Width =0
                    Height =383
                    Name ="Line28"
                End
                Begin Label
                    TextFontCharSet =178
                    TextAlign =3
                    ReadingOrder =2
                    NumeralShapes =2
                    Left =10444
                    Top =513
                    Width =360
                    Height =330
                    FontWeight =700
                    Name ="Label257"
                    Caption ="رديف"
                    FontName ="Titr"
                End
                Begin Label
                    TextFontCharSet =178
                    TextAlign =2
                    ReadingOrder =2
                    NumeralShapes =2
                    Left =9585
                    Top =540
                    Width =810
                    Height =315
                    FontWeight =700
                    Name ="Label258"
                    Caption ="شماره فاكتور"
                    FontName ="Titr"
                End
                Begin Label
                    TextFontCharSet =178
                    TextAlign =2
                    ReadingOrder =2
                    NumeralShapes =2
                    Left =5940
                    Top =555
                    Width =3510
                    Height =315
                    FontWeight =700
                    Name ="Label259"
                    Caption ="نام مشتري"
                    FontName ="Titr"
                End
                Begin Label
                    TextFontCharSet =178
                    TextAlign =2
                    ReadingOrder =2
                    NumeralShapes =2
                    Left =4740
                    Top =555
                    Width =1137
                    Height =315
                    FontWeight =700
                    Name ="Label260"
                    Caption ="مبلغ فاكتور"
                    FontName ="Titr"
                End
                Begin Label
                    TextFontCharSet =178
                    TextAlign =2
                    ReadingOrder =2
                    NumeralShapes =2
                    Left =3600
                    Top =555
                    Width =1110
                    Height =315
                    FontWeight =700
                    Name ="Label261"
                    Caption ="نقد"
                    FontName ="Titr"
                End
                Begin Label
                    TextFontCharSet =178
                    TextAlign =2
                    ReadingOrder =2
                    NumeralShapes =2
                    Left =2430
                    Top =555
                    Width =1077
                    Height =315
                    FontWeight =700
                    Name ="Label262"
                    Caption ="مبلغ چكها"
                    FontName ="Titr"
                End
                Begin Label
                    TextFontCharSet =178
                    TextAlign =2
                    ReadingOrder =2
                    NumeralShapes =2
                    Left =1230
                    Top =555
                    Width =1092
                    Height =315
                    FontWeight =700
                    Name ="Label263"
                    Caption ="ساير"
                    FontName ="Titr"
                End
                Begin Label
                    TextFontCharSet =178
                    TextAlign =2
                    ReadingOrder =2
                    NumeralShapes =2
                    Top =555
                    Width =1152
                    Height =315
                    FontWeight =700
                    Name ="Label264"
                    Caption ="مانده"
                    FontName ="Titr"
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =390
            Name ="Detail"
            Begin
                Begin TextBox
                    TextFontCharSet =178
                    TextAlign =2
                    IMESentenceMode =3
                    Left =9570
                    Top =30
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
                    Left =4755
                    Top =30
                    Width =1152
                    Height =330
                    FontWeight =700
                    TabIndex =1
                    Name ="MABK"
                    ControlSource ="SMAB"
                    Format ="#,###"
                    FontName ="Titr"
                End
                Begin TextBox
                    TextFontCharSet =178
                    TextAlign =3
                    IMESentenceMode =3
                    Left =3570
                    Top =30
                    Width =1152
                    Height =330
                    FontWeight =700
                    TabIndex =2
                    Name ="NAGHD"
                    ControlSource ="M_NAGHD"
                    Format ="#,###"
                    FontName ="Titr"
                End
                Begin TextBox
                    TextFontCharSet =178
                    TextAlign =3
                    IMESentenceMode =3
                    Left =2400
                    Top =30
                    Width =1152
                    Height =330
                    FontWeight =700
                    TabIndex =3
                    Name ="CHKM"
                    ControlSource ="CHK"
                    Format ="#,###"
                    FontName ="Titr"
                    InputMask ="##/##/##"
                End
                Begin TextBox
                    TextFontCharSet =178
                    TextAlign =3
                    IMESentenceMode =3
                    Left =30
                    Top =30
                    Width =1152
                    Height =330
                    FontWeight =700
                    TabIndex =4
                    Name ="MAND"
                    ControlSource ="MAND"
                    Format ="#,###"
                    FontName ="Titr"
                End
                Begin TextBox
                    TextFontCharSet =178
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1215
                    Top =30
                    Width =1152
                    Height =330
                    FontWeight =700
                    TabIndex =5
                    Name ="SAYER"
                    ControlSource ="=[MABL_HAV]+[MABL_VAR]"
                    Format ="#,###"
                    FontName ="Titr"
                End
                Begin TextBox
                    TextFontCharSet =178
                    TextAlign =3
                    IMESentenceMode =3
                    Left =5970
                    Top =30
                    Width =3540
                    Height =330
                    FontWeight =700
                    TabIndex =6
                    Name ="PERSON1"
                    ControlSource ="NAME"
                    FontName ="Titr"
                End
                Begin TextBox
                    RunningSum =2
                    TextFontCharSet =178
                    TextAlign =2
                    IMESentenceMode =3
                    Left =10410
                    Width =345
                    Height =375
                    FontWeight =700
                    TabIndex =7
                    Name ="Text11"
                    ControlSource ="=1"
                    FontName ="Titr"
                End
                Begin Line
                    Left =10403
                    Width =0
                    Height =383
                    Name ="Line12"
                End
                Begin Line
                    Left =9540
                    Width =0
                    Height =383
                    Name ="Line13"
                End
                Begin Line
                    Left =5925
                    Width =0
                    Height =383
                    Name ="Line14"
                End
                Begin Line
                    Left =4740
                    Width =0
                    Height =383
                    Name ="Line15"
                End
                Begin Line
                    Left =3555
                    Width =0
                    Height =383
                    Name ="Line16"
                End
                Begin Line
                    Left =2385
                    Width =0
                    Height =383
                    Name ="Line17"
                End
                Begin Line
                    Left =1200
                    Width =0
                    Height =383
                    Name ="Line18"
                End
                Begin Line
                    BorderWidth =3
                    Left =10815
                    Width =0
                    Height =390
                    Name ="Line20"
                End
                Begin Line
                    BorderWidth =3
                    Width =0
                    Height =390
                    Name ="Line21"
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
                    ControlSource ="=Sum([MAND])"
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
' See "BEDEHKARAN_USERS_F.cls"
