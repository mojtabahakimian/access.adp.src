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
    Width =6753
    RowHeight =300
    ItemSuffix =65
    Left =315
    DatasheetForeColor =33554432
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x8bac84d8f4ece440
    End
    RecordSource ="SELECT PERSONEL.BIMEH_NUM, PERSONEL.KHNOWNUM, PERSONEL.FATHER, PERSONEL.WSDATE, "
        "PERSONEL.MELLICOD, PHOKM.HONUM, PHOKM.CODE, PHOKM.HDATE, PHOKM.SALARY_DAYLY, PHO"
        "KM.SALARY_DAYLYB, PHOKM.HOME, PHOKM.CHILDREN, PHOKM.CONDITIONS, PHOKM.BON, PHOKM"
        ".JAZB, PHOKM.SAYER, PHOKM.MADRAK, PHOKM.TAHOL, PHOKM.MABEZT, PHOKM.MABEZ, PHOKM."
        "HSHIFT, PHOKM.bonj FROM PHOKM INNER JOIN PERSONEL ON PHOKM.CODE = PERSONEL.CODE"
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
                    Caption =" حكم كارگزيني"
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
            Height =8163
            Name ="Detail"
            Begin
                Begin TextBox
                    TabStop = NotDefault
                    SpecialEffect =0
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4733
                    Top =113
                    Width =906
                    Height =315
                    ColumnWidth =945
                    ColumnOrder =0
                    Name ="HONUM"
                    ControlSource ="HONUM"
                    Begin
                        Begin Label
                            TextFontCharSet =0
                            TextAlign =3
                            Left =5680
                            Top =113
                            Width =945
                            Height =300
                            Name ="Label1"
                            Caption ="شماره حكم:"
                        End
                    End
                End
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
                    Left =75
                    Top =120
                    Width =3276
                    Height =315
                    ColumnWidth =3060
                    ColumnOrder =1
                    TabIndex =1
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
                            Left =3390
                            Top =120
                            Width =1305
                            Height =330
                            Name ="Label22"
                            Caption ="نام ونام خانوادگي:"
                        End
                    End
                End
                Begin Label
                    TextFontCharSet =0
                    TextAlign =3
                    Top =4377
                    Width =2820
                    Height =735
                    Name ="Label23"
                    Caption ="توضيح : تمام مبالغ  بجز حق خواربار و مسكن و حق اولاد بر اساس روز محاسبه مي گردد"
                End
                Begin TextBox
                    DecimalPlaces =0
                    SpecialEffect =0
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3152
                    Top =2010
                    Width =1836
                    Height =315
                    ColumnWidth =1590
                    TabIndex =3
                    BackColor =16777164
                    ForeColor =10040115
                    Name ="SALARY_DAYLYB"
                    ControlSource ="SALARY_DAYLYB"
                    Format ="#,##0.00;-#,##0.00"
                    OnKeyPress ="[Event Procedure]"
                    Begin
                        Begin Label
                            TextFontCharSet =0
                            TextAlign =1
                            Left =5081
                            Top =2017
                            Width =1515
                            Height =300
                            Name ="Label15"
                            Caption ="حقوق روزانه رسمي:"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    SpecialEffect =0
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3152
                    Top =2820
                    Width =1836
                    Height =315
                    ColumnWidth =1620
                    TabIndex =5
                    Name ="HOME"
                    ControlSource ="HOME"
                    Format ="#,##0.00;-#,##0.00"
                    OnKeyPress ="[Event Procedure]"
                    Begin
                        Begin Label
                            TextFontCharSet =0
                            TextAlign =1
                            Left =5083
                            Top =2820
                            Width =1545
                            Height =300
                            Name ="Label17"
                            Caption ="حق خواربار و مسكن:"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    SpecialEffect =0
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3152
                    Top =3232
                    Width =1836
                    Height =315
                    TabIndex =6
                    Name ="CHILDREN"
                    ControlSource ="CHILDREN"
                    Format ="#,##0.00;-#,##0.00"
                    OnKeyPress ="[Event Procedure]"
                    Begin
                        Begin Label
                            TextFontCharSet =0
                            TextAlign =1
                            Left =5085
                            Top =3225
                            Width =990
                            Height =300
                            Name ="Label19"
                            Caption ="حق اولاد:"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    SpecialEffect =0
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3152
                    Top =3630
                    Width =1836
                    Height =315
                    TabIndex =7
                    Name ="CONDITIONS"
                    ControlSource ="CONDITIONS"
                    Format ="#,##0.00;-#,##0.00"
                    OnKeyPress ="[Event Procedure]"
                    Begin
                        Begin Label
                            TextFontCharSet =0
                            TextAlign =1
                            Left =5081
                            Top =3630
                            Width =1260
                            Height =300
                            Name ="Label21"
                            Caption ="شرايط محيط كار:"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    SpecialEffect =0
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3152
                    Top =4042
                    Width =1836
                    Height =315
                    TabIndex =8
                    Name ="BON"
                    ControlSource ="BON"
                    Format ="#,##0.00;-#,##0.00"
                    OnKeyPress ="[Event Procedure]"
                    Begin
                        Begin Label
                            TextFontCharSet =0
                            TextAlign =1
                            Left =5071
                            Top =4035
                            Width =690
                            Height =300
                            Name ="Label24"
                            Caption ="حق نهار:"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    SpecialEffect =0
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3152
                    Top =4447
                    Width =1836
                    Height =315
                    TabIndex =9
                    Name ="JAZB"
                    ControlSource ="JAZB"
                    Format ="#,##0.00;-#,##0.00"
                    OnKeyPress ="[Event Procedure]"
                    Begin
                        Begin Label
                            TextFontCharSet =0
                            TextAlign =1
                            Left =5070
                            Top =4440
                            Width =1005
                            Height =300
                            Name ="Label25"
                            Caption ="حق جذب:"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    SpecialEffect =0
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3152
                    Top =4852
                    Width =1836
                    Height =315
                    TabIndex =10
                    Name ="SAYER"
                    ControlSource ="SAYER"
                    Format ="#,##0.00;-#,##0.00"
                    OnKeyPress ="[Event Procedure]"
                    Begin
                        Begin Label
                            TextFontCharSet =0
                            TextAlign =1
                            Left =5070
                            Top =4845
                            Width =765
                            Height =285
                            Name ="Label27"
                            Caption ="ساير:"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    SpecialEffect =0
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3152
                    Top =2415
                    Width =1836
                    Height =315
                    ColumnWidth =1560
                    TabIndex =4
                    BackColor =13434828
                    ForeColor =10040115
                    Name ="SALARY_DAYLY"
                    ControlSource ="SALARY_DAYLY"
                    Format ="#,##0.00;-#,##0.00"
                    OnKeyPress ="[Event Procedure]"
                    Begin
                        Begin Label
                            TextFontCharSet =0
                            TextAlign =1
                            Left =5078
                            Top =2422
                            Width =1515
                            Height =300
                            Name ="Label61"
                            Caption ="حقوق روزانه اسمي:"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    SpecialEffect =0
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3152
                    Top =1587
                    Width =1836
                    Height =315
                    TabIndex =2
                    Name ="HDATE"
                    ControlSource ="HDATE"
                    BeforeUpdate ="[Event Procedure]"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            TextFontCharSet =0
                            TextAlign =1
                            Left =5083
                            Top =1594
                            Width =1515
                            Height =300
                            Name ="Label28"
                            Caption ="تاريخ حكم:"
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    DecimalPlaces =0
                    SpecialEffect =0
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3180
                    Top =6894
                    Width =1761
                    Height =315
                    TabIndex =14
                    BackColor =13434828
                    ForeColor =10040115
                    Name ="jam"
                    ControlSource ="=([SAYER]+[SALARY_DAYLY]+[CONDITIONS]+[BON]+[JAZB])*30+[HOME]+[CHILDREN]"
                    Format ="#,##0.00;-#,##0.00"
                    Begin
                        Begin Label
                            TextFontCharSet =0
                            TextAlign =1
                            Left =5039
                            Top =6885
                            Width =1650
                            Height =240
                            Name ="Label32"
                            Caption ="حقوق ماهيانه  اسمي:"
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    DecimalPlaces =0
                    SpecialEffect =0
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3182
                    Top =6474
                    Width =1761
                    Height =315
                    TabIndex =15
                    BackColor =16777164
                    ForeColor =10040115
                    Name ="Text33"
                    ControlSource ="=([SAYER]+[SALARY_DAYLYB]+[CONDITIONS]+[BON]+[JAZB])*30+[HOME]+[CHILDREN]"
                    Format ="#,##0.00;-#,##0.00"
                    Begin
                        Begin Label
                            TextFontCharSet =0
                            TextAlign =1
                            Left =5041
                            Top =6465
                            Width =1620
                            Height =240
                            Name ="Label34"
                            Caption ="حقوق ماهيانه رسمي:"
                        End
                    End
                End
                Begin ComboBox
                    SpecialEffect =0
                    OldBorderStyle =0
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    DecimalPlaces =0
                    NumeralShapes =2
                    ColumnCount =2
                    Left =61
                    Top =1602
                    Height =315
                    TabIndex =12
                    Name ="MADRAK"
                    ControlSource ="MADRAK"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TCODE_MADRAK.CODE, TCODE_MADRAK.MADRAK_NAME FROM TCODE_MADRAK GROUP BY TC"
                        "ODE_MADRAK.CODE, TCODE_MADRAK.MADRAK_NAME ORDER BY TCODE_MADRAK.MADRAK_NAME;"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            TextFontCharSet =0
                            TextAlign =3
                            Left =1752
                            Top =1609
                            Width =1095
                            Height =300
                            Name ="Label38"
                            Caption ="مدرك تحصيلي:"
                        End
                    End
                End
                Begin OptionGroup
                    SpecialEffect =0
                    Left =120
                    Top =2085
                    Width =2312
                    Height =553
                    TabIndex =13
                    Name ="Frame41"
                    ControlSource ="TAHOL"
                    DefaultValue ="2"
                    Begin
                        Begin Label
                            BackStyle =1
                            TextFontCharSet =0
                            TextAlign =2
                            Left =1762
                            Top =1932
                            Width =450
                            Height =300
                            Name ="Label42"
                            Caption ="تاهل"
                        End
                        Begin OptionButton
                            SpecialEffect =0
                            ReadingOrder =1
                            Left =735
                            Top =2263
                            OptionValue =2
                            Name ="Option46"
                            Begin
                                Begin Label
                                    TextFontCharSet =0
                                    TextAlign =3
                                    Left =315
                                    Top =2235
                                    Width =450
                                    Height =300
                                    Name ="Label47"
                                    Caption ="مجرد"
                                End
                            End
                        End
                        Begin OptionButton
                            SpecialEffect =0
                            ReadingOrder =1
                            Left =1745
                            Top =2286
                            OptionValue =1
                            Name ="Option44"
                            Begin
                                Begin Label
                                    TextFontCharSet =0
                                    TextAlign =3
                                    Left =1235
                                    Top =2258
                                    Width =555
                                    Height =300
                                    Name ="Label45"
                                    Caption ="متاهل"
                                End
                            End
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    SpecialEffect =0
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =42
                    Top =2949
                    Width =1371
                    Height =315
                    TabIndex =16
                    Name ="MABEZ"
                    ControlSource ="MABEZ"
                    Format ="#,##0.00;-#,##0.00"
                    OnKeyPress ="[Event Procedure]"
                    Begin
                        Begin Label
                            TextFontCharSet =0
                            TextAlign =3
                            Left =1440
                            Top =2955
                            Width =1020
                            Height =240
                            Name ="Label49"
                            Caption ="نرخ اضافه كار:"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    SpecialEffect =0
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =42
                    Top =3384
                    Width =1371
                    Height =315
                    TabIndex =17
                    Name ="MABEZT"
                    ControlSource ="MABEZT"
                    Format ="#,##0.00;-#,##0.00"
                    Begin
                        Begin Label
                            TextFontCharSet =0
                            TextAlign =3
                            Left =1440
                            Top =3390
                            Width =1560
                            Height =240
                            Name ="Label51"
                            Caption ="نرخ اضافه كارتعطيلي:"
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3151
                    Top =5250
                    Width =1836
                    Height =315
                    TabIndex =11
                    Name ="HSHIFT"
                    ControlSource ="HSHIFT"
                    Format ="#,##0.00;-#,##0.00"
                    Begin
                        Begin Label
                            TextFontCharSet =0
                            TextAlign =1
                            Left =5071
                            Top =5250
                            Width =1500
                            Height =480
                            Name ="Label53"
                            Caption ="درصد حق شيفت،شب كاري:"
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
                    TabIndex =18
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
                    TabIndex =19
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
                    TabIndex =20
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
                    TabIndex =21
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
                Begin Line
                    LineSlant = NotDefault
                    Top =6375
                    Width =6746
                    Name ="Line59"
                End
                Begin Label
                    TextFontCharSet =0
                    TextAlign =1
                    Left =4769
                    Top =7785
                    Width =915
                    Height =240
                    Name ="Label60"
                    Caption ="امضاء كارگر:"
                End
                Begin Label
                    TextFontCharSet =0
                    TextAlign =1
                    Left =630
                    Top =7785
                    Width =1020
                    Height =240
                    Name ="Label62"
                    Caption ="امضاء كارفرما:"
                End
                Begin TextBox
                    DecimalPlaces =0
                    SpecialEffect =0
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3151
                    Top =5850
                    Width =1836
                    Height =315
                    TabIndex =22
                    Name ="bonj"
                    ControlSource ="bonj"
                    Format ="#,##0.00;-#,##0.00"
                    Begin
                        Begin Label
                            TextFontCharSet =0
                            TextAlign =1
                            Left =5071
                            Top =5850
                            Width =300
                            Height =240
                            Name ="Label64"
                            Caption ="بن:"
                        End
                    End
                End
            End
        End
        Begin FormFooter
            Height =0
            BackColor =-2147483633
            Name ="FormFooter"
        End
    End
End
CodeBehindForm
' See "RPHOKM.cls"
