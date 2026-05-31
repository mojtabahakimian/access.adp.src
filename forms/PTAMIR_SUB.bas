Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    AllowEdits = NotDefault
    DefaultView =2
    AllowUpdating =4
    ViewsAllowed =2
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =11326
    RowHeight =315
    Left =525
    Top =1095
    Right =15045
    Bottom =9465
    DatasheetGridlinesColor =12632256
    MaxRecords =100000
    RecSrcDt = Begin
        0x78b7c6a7ed2ce340
    End
    RecordSource ="PTAMIRAT"
    DatasheetFontName ="Tahoma"
    OnError ="[Event Procedure]"
    AllowFormView =0
    RecordSourceQualifier ="dbo"
    Begin
        Begin Label
            BackStyle =0
            TextFontCharSet =178
            FontName ="Tahoma"
        End
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            OldBorderStyle =0
            TextFontCharSet =178
            Width =1701
            LabelX =-1701
            FontName ="Tahoma"
            AsianLineBreak =255
        End
        Begin ComboBox
            SpecialEffect =2
            TextFontCharSet =178
            Width =1701
            LabelX =-1701
            FontName ="Tahoma"
        End
        Begin Subform
            SpecialEffect =2
            Width =1701
            Height =1701
        End
        Begin Section
            CanGrow = NotDefault
            Height =11506
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =8445
                    Top =7
                    Height =315
                    ColumnWidth =1125
                    ColumnOrder =0
                    Name ="PIDT"
                    ControlSource ="PIDT"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =10203
                            Top =7
                            Width =1020
                            Height =340
                            Name ="Label1"
                            Caption ="شماره پذيرش"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    NumeralShapes =3
                    IMESentenceMode =3
                    Left =4522
                    Width =3120
                    Height =340
                    ColumnWidth =1980
                    ColumnOrder =1
                    TabIndex =1
                    BackColor =13434828
                    Name ="PSERIAL"
                    ControlSource ="PSERIAL"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =7760
                            Width =510
                            Height =340
                            Name ="Label3"
                            Caption ="سريال"
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
                    Left =2122
                    Top =7
                    Width =1350
                    Height =340
                    ColumnWidth =990
                    ColumnOrder =2
                    TabIndex =2
                    Name ="PINDATE"
                    ControlSource ="PINDATE"
                    DefaultValue ="=FARSIDATE(Now())"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3533
                            Top =7
                            Width =885
                            Height =340
                            Name ="Label5"
                            Caption ="تاريخ پذيرش"
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
                    Top =7
                    Width =801
                    Height =340
                    ColumnWidth =1140
                    ColumnOrder =3
                    TabIndex =3
                    Name ="PINTIME"
                    ControlSource ="PINTIME"
                    DefaultValue ="Hour(Now())*100+Minute(Now())"
                    InputMask ="##:##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1039
                            Top =7
                            Width =1035
                            Height =340
                            Name ="Label7"
                            Caption ="ساعت پذيرش"
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
                    Left =3179
                    Top =442
                    Width =3960
                    Height =340
                    ColumnWidth =2355
                    ColumnOrder =4
                    TabIndex =4
                    Name ="MODEL"
                    ControlSource ="MODEL"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =7200
                            Top =442
                            Width =360
                            Height =340
                            Name ="Label11"
                            Caption ="مدل"
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
                    Top =442
                    Width =1575
                    Height =340
                    ColumnOrder =5
                    TabIndex =5
                    Name ="GRENDDATE"
                    ControlSource ="GRENDDATE"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1738
                            Top =442
                            Width =1320
                            Height =340
                            Name ="Label13"
                            Caption ="تاريخ اتمام گارانتي"
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
                    Top =2166
                    Width =5220
                    Height =330
                    ColumnWidth =1665
                    ColumnOrder =9
                    TabIndex =6
                    Name ="EYBEZHAR"
                    ControlSource ="EYBEZHAR"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1700
                            Top =1882
                            Width =1560
                            Height =240
                            Name ="Label29"
                            Caption ="عيب به اظهار مشتري"
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
                    Left =5385
                    Top =2167
                    Width =4920
                    Height =345
                    ColumnOrder =10
                    TabIndex =7
                    Name ="LAVAZEM"
                    ControlSource ="LAVAZEM"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =10374
                            Top =2167
                            Width =870
                            Height =240
                            Name ="Label31"
                            Caption ="لوازم همراه"
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
                    Top =2551
                    Width =10200
                    Height =340
                    ColumnOrder =11
                    TabIndex =8
                    Name ="IDIATEK"
                    ControlSource ="IDIATEK"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =10311
                            Top =2551
                            Width =915
                            Height =340
                            Name ="Label33"
                            Caption ="نظر تكنسين"
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
                    Top =2968
                    Width =10200
                    Height =340
                    ColumnWidth =5505
                    ColumnOrder =12
                    TabIndex =9
                    Name ="ADRESS"
                    ControlSource ="ADRESS"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =10358
                            Top =2968
                            Width =930
                            Height =340
                            Name ="Label35"
                            Caption ="آدرس و تلفن"
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
                    Top =3376
                    Width =10170
                    Height =645
                    ColumnWidth =2385
                    ColumnOrder =13
                    TabIndex =10
                    Name ="PAYAM"
                    ControlSource ="PAYAM"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =10261
                            Top =3376
                            Width =1065
                            Height =340
                            Name ="Label37"
                            Caption ="متن زير پذيرش"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =7702
                    Top =442
                    Width =2616
                    Height =340
                    ColumnOrder =14
                    TabIndex =11
                    Name ="PKIND"
                    ControlSource ="PKIND"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT PKINDCO, PKIND FROM PKINDCODING"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =10424
                            Top =442
                            Width =795
                            Height =340
                            Name ="Label9"
                            Caption ="نوع پذيرش"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =3
                    Left =5377
                    Top =880
                    Width =4830
                    Height =340
                    ColumnOrder =15
                    TabIndex =12
                    Name ="POWNER"
                    ControlSource ="POWNER"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT RTRIM(CAST(N_KOL AS nvarchar)) + '-' + RTRIM(CAST(NUMBER AS nvarchar)) + "
                        "'-' + RTRIM(CAST(TNUMBER AS nvarchar)) AS hes, NAME AS nam, RTRIM(CAST(N_KOL AS "
                        "nvarchar)) + '-' + RTRIM(CAST(NUMBER AS nvarchar)) + '-' + RTRIM(CAST(TNUMBER AS"
                        " nvarchar)) AS Expr1 FROM TDETA_HES"
                    ColumnWidths ="0;3689;1134"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =10313
                            Top =880
                            Width =945
                            Height =340
                            Name ="Label15"
                            Caption ="مالك دستگاه"
                        End
                    End
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Top =877
                    Width =4320
                    Height =340
                    ColumnOrder =16
                    TabIndex =13
                    Name ="POWNERNAM"
                    ControlSource ="POWNERNAM"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT POWNERNAM FROM PTAMIRAT GROUP BY POWNERNAM ORDER BY POWNERNAM"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4414
                            Top =877
                            Width =900
                            Height =340
                            Name ="Label17"
                            Caption ="نماينده مالك"
                        End
                    End
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =5377
                    Top =1308
                    Width =4380
                    Height =340
                    ColumnOrder =17
                    TabIndex =14
                    Name ="GRCOMPA"
                    ControlSource ="GRCOMPA"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT GRCOMPA FROM PTAMIRAT GROUP BY GRCOMPA ORDER BY GRCOMPA"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =9803
                            Top =1308
                            Width =1485
                            Height =340
                            Name ="Label19"
                            Caption ="شركت گارانتي كننده"
                        End
                    End
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2726
                    Top =1297
                    Width =2040
                    Height =340
                    ColumnOrder =18
                    TabIndex =15
                    Name ="IYALAT"
                    ControlSource ="IYALAT"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT IYALAT FROM PTAMIRAT GROUP BY IYALAT ORDER BY IYALAT"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4812
                            Top =1297
                            Width =495
                            Height =340
                            Name ="Label21"
                            Caption ="استان"
                        End
                    End
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Top =1308
                    Width =1695
                    Height =340
                    ColumnOrder =19
                    TabIndex =16
                    Name ="CITY"
                    ControlSource ="CITY"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT CITY FROM PTAMIRAT GROUP BY CITY ORDER BY CITY"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1868
                            Top =1308
                            Width =780
                            Height =340
                            Name ="Label23"
                            Caption ="شهرستان"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =8612
                    Top =1717
                    Height =340
                    ColumnWidth =1545
                    ColumnOrder =20
                    TabIndex =17
                    Name ="PINPKIND"
                    ControlSource ="PINPKIND"
                    RowSourceType ="Value List"
                    RowSource ="1;\"در محل تعميرگاه\";2;\"در محل مشتري\""
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =10368
                            Top =1717
                            Width =885
                            Height =340
                            Name ="Label25"
                            Caption ="محل پذيرش"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =3
                    Left =5908
                    Top =4126
                    Width =4515
                    Height =340
                    ColumnWidth =1575
                    ColumnOrder =6
                    TabIndex =18
                    Name ="TAMIRKAR1"
                    ControlSource ="TAMIRKAR1"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT RTRIM(CAST(N_KOL AS nvarchar)) + '-' + RTRIM(CAST(NUMBER AS nvarchar)) + "
                        "'-' + RTRIM(CAST(TNUMBER AS nvarchar)) AS hes, NAME AS nam, RTRIM(CAST(N_KOL AS "
                        "nvarchar)) + '-' + RTRIM(CAST(NUMBER AS nvarchar)) + '-' + RTRIM(CAST(TNUMBER AS"
                        " nvarchar)) AS Expr1 FROM TDETA_HES"
                    ColumnWidths ="0;3690;1134"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =10484
                            Top =4126
                            Width =765
                            Height =340
                            Name ="Label39"
                            Caption ="تعمير كار1"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =3
                    Top =4126
                    Width =4950
                    Height =340
                    ColumnOrder =7
                    TabIndex =19
                    Name ="TAMIRKAR2"
                    ControlSource ="TAMIRKAR2"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT RTRIM(CAST(N_KOL AS nvarchar)) + '-' + RTRIM(CAST(NUMBER AS nvarchar)) + "
                        "'-' + RTRIM(CAST(TNUMBER AS nvarchar)) AS hes, NAME AS nam, RTRIM(CAST(N_KOL AS "
                        "nvarchar)) + '-' + RTRIM(CAST(NUMBER AS nvarchar)) + '-' + RTRIM(CAST(TNUMBER AS"
                        " nvarchar)) AS Expr1 FROM TDETA_HES"
                    ColumnWidths ="0;3690;1134"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5041
                            Top =4126
                            Width =810
                            Height =340
                            Name ="Label41"
                            Caption ="تعمير كار 2"
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
                    Left =5370
                    Top =1732
                    Width =2166
                    Height =340
                    ColumnWidth =1110
                    ColumnOrder =8
                    TabIndex =20
                    Name ="USER_NAME"
                    ControlSource ="USER_NAME"
                    DefaultValue ="=UCURRENTUSER()"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =7590
                            Top =1727
                            Width =960
                            Height =330
                            Name ="Label27"
                            Caption ="كاربر پذيرش"
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
                    Left =8977
                    Top =4539
                    Width =1350
                    Height =340
                    ColumnWidth =975
                    TabIndex =21
                    Name ="TAHV_DATE"
                    ControlSource ="TAHV_DATE"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =10388
                            Top =4539
                            Width =885
                            Height =340
                            Name ="Label43"
                            Caption ="تاريخ  تحويل"
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
                    Left =6930
                    Top =4539
                    Width =801
                    Height =340
                    ColumnWidth =1080
                    TabIndex =22
                    Name ="TAHV_TIME"
                    ControlSource ="TAHV_TIME"
                    InputMask ="##:##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =7894
                            Top =4539
                            Width =1035
                            Height =340
                            Name ="Label45"
                            Caption ="ساعت تحويل"
                        End
                    End
                End
                Begin Subform
                    OverlapFlags =85
                    Left =790
                    Top =5213
                    Width =7251
                    Height =1191
                    TabIndex =23
                    Name ="PTVAZ_SUB"
                    SourceObject ="Form.PTVAZ_SUB"
                    LinkChildFields ="PIDT"
                    LinkMasterFields ="PIDT"
                End
            End
        End
    End
End
CodeBehindForm
' See "PTAMIR_SUB.cls"
