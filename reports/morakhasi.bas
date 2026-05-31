Version =20
VersionRequired =20
Begin Report
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    DividingLines = NotDefault
    OrderByOn = NotDefault
    DefaultView =0
    AllowUpdating =4
    ScrollBars =0
    TabularCharSet =178
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =6747
    RowHeight =300
    ItemSuffix =78
    Left =315
    DatasheetForeColor =33554432
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x1b59f87d93ece540
    End
    RecordSource ="SELECT PERSONEL.BIMEH_NUM, PERSONEL.KHNOWNUM, PERSONEL.FATHER, PERSONEL.WSDATE, "
        "PERSONEL.MELLICOD, PMORAKH.MODATE, PMORAKH.MORAKHDAY, PMORAKH.MAND, PMORAKH.KIND"
        "M, PMORAKH.MOLAH, PMORAKH.CODE, PMORAKH.MOSTDATE, PMORAKH.MOENDATE, PMORAKH.IDNU"
        "M, PMORAKH.sgn1usid, PMORAKH.sgn2usid, PMORAKH.sgn3usid, PMORAKH.SGN1, PMORAKH.S"
        "GN2, PMORAKH.SGN3, SALA_DTL.EMZA AS EMZA1, SALA_DTL_1.EMZA AS EMZA2, SALA_DTL_2."
        "EMZA AS EMZA3 FROM SALA_DTL RIGHT OUTER JOIN SALA_DTL SALA_DTL_1 RIGHT OUTER JOI"
        "N PERSONEL INNER JOIN PMORAKH ON PERSONEL.CODE = PMORAKH.CODE LEFT OUTER JOIN SA"
        "LA_DTL SALA_DTL_2 ON PMORAKH.sgn3usid = SALA_DTL_2.IDD ON SALA_DTL_1.IDD = PMORA"
        "KH.sgn2usid ON SALA_DTL.IDD = PMORAKH.sgn1usid"
    Caption ="حكم كارگزيني پرسنل"
    DatasheetFontName ="Tahoma"
    OnLoad ="[Event Procedure]"
    OnError ="[Event Procedure]"
    RecordSourceQualifier ="dbo"
    Begin
        Begin Label
            BackStyle =0
            TextFontCharSet =178
            FontName ="Tahoma"
        End
        Begin Rectangle
            Width =850
            Height =850
        End
        Begin Line
            Width =1701
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
        Begin FormHeader
            Height =885
            Name ="FormHeader"
            Begin
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =2625
                    Top =450
                    Width =1440
                    Height =435
                    FontSize =12
                    Name ="Label30"
                    Caption ="برگ مرخصي"
                    FontName ="Titr"
                End
                Begin TextBox
                    SpecialEffect =0
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    BackStyle =0
                    NumeralShapes =2
                    Width =6693
                    Height =435
                    FontSize =12
                    BorderColor =32768
                    Name ="Text90"
                    ControlSource ="=[Forms]![Baseknow]![WIDTH_D]"
                    FontName ="Titr"
                End
            End
        End
        Begin Section
            SpecialEffect =1
            Height =6435
            OnFormat ="[Event Procedure]"
            Name ="Detail"
            Begin
                Begin ComboBox
                    LimitToList = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =1984
                    Top =113
                    Width =3276
                    Height =315
                    ColumnWidth =3060
                    ColumnOrder =1
                    Name ="CODE"
                    ControlSource ="CODE"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT CODE, PNAME + N' ' + PFAMILY + N' ' + RTRIM(CAST(CODE AS NVARCHAR)) AS PE"
                        "R FROM PERSONEL ORDER BY PNAME + N' ' + PFAMILY + N' ' + RTRIM(CAST(CODE AS NVAR"
                        "CHAR))"
                    ColumnWidths ="0"
                    OnNotInList ="[Event Procedure]"
                    Begin
                        Begin Label
                            TextFontCharSet =0
                            TextAlign =3
                            Left =5299
                            Top =113
                            Width =1305
                            Height =330
                            Name ="Label22"
                            Caption ="نام ونام خانوادگي:"
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =2954
                    Top =570
                    Width =2256
                    Height =285
                    TabIndex =1
                    Name ="KHNOWNUM"
                    ControlSource ="KHNOWNUM"
                    Begin
                        Begin Label
                            TextAlign =3
                            Left =5276
                            Top =570
                            Width =1350
                            Height =285
                            Name ="Label54"
                            Caption ="شماره شناسنامه:"
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =3230
                    Top =907
                    Width =2781
                    Height =285
                    TabIndex =2
                    Name ="FATHER"
                    ControlSource ="FATHER"
                    Begin
                        Begin Label
                            TextAlign =3
                            Left =6062
                            Top =907
                            Width =570
                            Height =285
                            Name ="Label55"
                            Caption ="نام پدر:"
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =117
                    Top =907
                    Height =285
                    TabIndex =3
                    Name ="WSDATE"
                    ControlSource ="WSDATE"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            TextAlign =3
                            Left =1854
                            Top =907
                            Width =1320
                            Height =285
                            Name ="Label56"
                            Caption ="تاريخ شروع به كار:"
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =59
                    Top =570
                    Width =1836
                    Height =285
                    TabIndex =4
                    Name ="MELLICOD"
                    ControlSource ="MELLICOD"
                    Begin
                        Begin Label
                            TextAlign =3
                            Left =1961
                            Top =570
                            Width =705
                            Height =285
                            Name ="Label57"
                            Caption ="كد ملي :"
                        End
                    End
                End
                Begin Rectangle
                    BackStyle =0
                    Left =30
                    Top =60
                    Width =6663
                    Height =1304
                    Name ="Box58"
                End
                Begin TextBox
                    SpecialEffect =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4137
                    Top =1530
                    Width =1236
                    ColumnWidth =1485
                    TabIndex =5
                    Name ="MODATE"
                    ControlSource ="MODATE"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            Left =5454
                            Top =1530
                            Width =1170
                            Height =240
                            Name ="Label63"
                            Caption ="تاريخ درخواست:"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    SpecialEffect =0
                    IMESentenceMode =3
                    Left =45
                    Top =3870
                    Width =561
                    ColumnWidth =780
                    TabIndex =6
                    Name ="MORAKHDAY"
                    ControlSource ="MORAKHDAY"
                    Begin
                        Begin Label
                            Visible = NotDefault
                            Left =1383
                            Top =3814
                            Width =420
                            Height =240
                            Name ="Label64"
                            Caption ="تعداد روز:"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    SpecialEffect =0
                    IMESentenceMode =3
                    Left =45
                    Top =4210
                    Width =561
                    TabIndex =7
                    Name ="MAND"
                    ControlSource ="MAND"
                    Begin
                        Begin Label
                            Visible = NotDefault
                            Left =1256
                            Top =4154
                            Width =660
                            Height =240
                            Name ="Label65"
                            Caption ="مانده:"
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =45
                    Top =2085
                    Width =5796
                    TabIndex =9
                    Name ="MOLAH"
                    ControlSource ="MOLAH"
                    Begin
                        Begin Label
                            Left =5895
                            Top =2090
                            Width =735
                            Height =240
                            Name ="Label67"
                            Caption ="توضيحات:"
                        End
                    End
                End
                Begin ComboBox
                    RowSourceTypeInt =1
                    SpecialEffect =0
                    OldBorderStyle =0
                    TextAlign =3
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =291
                    Top =1542
                    Width =2016
                    TabIndex =8
                    Name ="KINDM"
                    ControlSource ="KINDM"
                    RowSourceType ="Value List"
                    RowSource ="0;\"استحقاقي\";1;\"استراحت پزشكي\";2;\"ساعتي\";3;\"بدون حقوق\""
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            Left =2361
                            Top =1538
                            Width =975
                            Height =240
                            Name ="Label66"
                            Caption ="نوع مرخصي:"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3702
                    Top =3888
                    Width =666
                    Height =315
                    TabIndex =10
                    Name ="minu"
                    Format ="Fixed"
                    ValidationRule ="Not Is Null And <60"
                    DefaultValue ="0"
                    ControlTipText ="دقيقه"
                    Begin
                        Begin Label
                            TextFontCharSet =0
                            Left =4468
                            Top =3885
                            Width =1590
                            Height =300
                            Name ="Label11"
                            Caption ="مرخصي درخواستي"
                        End
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    DecimalPlaces =0
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3707
                    Top =4548
                    Width =666
                    Height =315
                    TabIndex =11
                    Name ="MMINU"
                    Format ="Fixed"
                    DefaultValue ="0"
                    ControlTipText ="دقيقه"
                    Begin
                        Begin Label
                            TextFontCharSet =0
                            Left =4472
                            Top =4548
                            Width =1080
                            Height =300
                            Name ="Label13"
                            Caption ="مانده مرخصي"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2979
                    Top =3888
                    Width =666
                    Height =315
                    TabIndex =12
                    Name ="HOU"
                    Format ="Fixed"
                    ValidationRule ="Not Is Null And <8"
                    DefaultValue ="0"
                    ControlTipText ="ساعت"
                End
                Begin TextBox
                    TabStop = NotDefault
                    DecimalPlaces =0
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2984
                    Top =4548
                    Width =666
                    Height =315
                    TabIndex =13
                    Name ="MHOU"
                    Format ="Fixed"
                    DefaultValue ="0"
                    ControlTipText ="ساعت"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2263
                    Top =3885
                    Width =666
                    Height =315
                    TabIndex =14
                    Name ="DA"
                    Format ="Fixed"
                    ValidationRule ="Not Is Null"
                    DefaultValue ="0"
                    ControlTipText ="روز"
                End
                Begin TextBox
                    TabStop = NotDefault
                    DecimalPlaces =0
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2268
                    Top =4545
                    Width =666
                    Height =315
                    TabIndex =15
                    Name ="MDA"
                    Format ="Fixed"
                    DefaultValue ="0"
                    ControlTipText ="روز"
                End
                Begin Label
                    BackStyle =1
                    TextFontCharSet =0
                    TextAlign =2
                    Left =2398
                    Top =3526
                    Width =390
                    Height =300
                    Name ="Label20"
                    Caption ="روز"
                End
                Begin Label
                    BackStyle =1
                    TextFontCharSet =0
                    TextAlign =2
                    Left =3041
                    Top =3525
                    Width =510
                    Height =300
                    Name ="Label21"
                    Caption ="ساعت"
                End
                Begin Label
                    BackStyle =1
                    TextFontCharSet =0
                    TextAlign =2
                    Left =3761
                    Top =3525
                    Width =510
                    Height =300
                    Name ="Label70"
                    Caption ="دقيقه"
                End
                Begin TextBox
                    SpecialEffect =0
                    TextAlign =2
                    IMESentenceMode =3
                    Left =4138
                    Top =2745
                    Width =1236
                    TabIndex =16
                    BackColor =13434828
                    Name ="Text71"
                    ControlSource ="MOSTDATE"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            Left =5455
                            Top =2745
                            Width =1170
                            Height =240
                            Name ="Label72"
                            Caption ="تاريخ شروع:"
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    TextAlign =2
                    IMESentenceMode =3
                    Left =281
                    Top =2751
                    Width =1236
                    TabIndex =17
                    BackColor =13434828
                    Name ="Text73"
                    ControlSource ="MOENDATE"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            Left =1598
                            Top =2751
                            Width =1170
                            Height =240
                            Name ="Label74"
                            Caption ="تاريخ پايان:"
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =60
                    Top =120
                    Width =651
                    Height =285
                    TabIndex =18
                    Name ="IDNUM"
                    ControlSource ="IDNUM"
                    Begin
                        Begin Label
                            TextAlign =3
                            Left =777
                            Top =120
                            Width =705
                            Height =285
                            Name ="Label76"
                            Caption ="شماره :"
                        End
                    End
                End
                Begin BoundObjectFrame
                    SizeMode =3
                    SpecialEffect =0
                    OldBorderStyle =0
                    Left =4316
                    Top =5024
                    Width =1236
                    Height =1176
                    TabIndex =19
                    Name ="EMZA"
                    ControlSource ="EMZA1"
                End
                Begin BoundObjectFrame
                    SizeMode =3
                    SpecialEffect =0
                    OldBorderStyle =0
                    Left =2171
                    Top =5024
                    Width =1236
                    Height =1176
                    TabIndex =20
                    Name ="OLEBound78"
                    ControlSource ="EMZA2"
                End
                Begin BoundObjectFrame
                    SizeMode =3
                    SpecialEffect =0
                    OldBorderStyle =0
                    Left =56
                    Top =5020
                    Width =1236
                    Height =1176
                    TabIndex =21
                    Name ="OLEBound79"
                    ControlSource ="EMZA3"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =5473
                    Top =5372
                    Width =960
                    Height =240
                    BackColor =128
                    Name ="semat1"
                    Caption ="."
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =3344
                    Top =5401
                    Width =960
                    Height =240
                    BackColor =128
                    Name ="semat2"
                    Caption ="."
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =1180
                    Top =5388
                    Width =960
                    Height =240
                    BackColor =128
                    Name ="semat3"
                    Caption ="."
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =5470
                    Top =5039
                    Width =960
                    Height =240
                    BackColor =128
                    Name ="nemz1"
                    Caption ="."
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =3346
                    Top =5068
                    Width =960
                    Height =240
                    BackColor =128
                    Name ="nemz2"
                    Caption ="."
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =1180
                    Top =5058
                    Width =960
                    Height =240
                    BackColor =128
                    Name ="nemz3"
                    Caption ="."
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    SpecialEffect =0
                    TextFontFamily =42
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3674
                    Top =3344
                    FontSize =9
                    TabIndex =22
                    Name ="SGN1usid"
                    ControlSource ="SGN1usid"
                    FontName ="Arial (Arabic)"
                End
                Begin TextBox
                    Visible = NotDefault
                    SpecialEffect =0
                    TextFontFamily =42
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3674
                    Top =3663
                    FontSize =9
                    TabIndex =23
                    Name ="sgn2usid"
                    ControlSource ="sgn2usid"
                    FontName ="Arial (Arabic)"
                End
                Begin TextBox
                    Visible = NotDefault
                    SpecialEffect =0
                    TextFontFamily =42
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2815
                    Top =3741
                    FontSize =9
                    TabIndex =24
                    Name ="sgn3usid"
                    ControlSource ="sgn3usid"
                    FontName ="Arial (Arabic)"
                End
            End
        End
        Begin FormFooter
            ForceNewPage =2
            Height =0
            BackColor =-2147483633
            Name ="FormFooter"
        End
    End
End
CodeBehindForm
' See "morakhasi.cls"
