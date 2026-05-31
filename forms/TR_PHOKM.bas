Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    AllowEdits = NotDefault
    DefaultView =0
    AllowUpdating =4
    ScrollBars =0
    TabularCharSet =178
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =12245
    RowHeight =300
    ItemSuffix =57
    Left =1995
    Top =795
    Right =13110
    Bottom =6660
    DatasheetForeColor =33554432
    DatasheetGridlinesColor =12632256
    MaxRecords =10000000
    RecSrcDt = Begin
        0x54478d044f28e440
    End
    RecordSource ="TR_PHOKM"
    Caption ="حكم كارگزيني پرسنل"
    DatasheetFontName ="Tahoma"
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
        Begin FormHeader
            Height =690
            BackColor =8421376
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    TextFontFamily =10
                    Left =5294
                    Top =15
                    Width =3000
                    Height =675
                    FontSize =18
                    ForeColor =16777215
                    Name ="Label29"
                    Caption ="صدور حكم كارگزيني"
                    FontName ="Titr"
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    TextFontFamily =10
                    Left =5294
                    Width =3000
                    Height =675
                    FontSize =18
                    Name ="Label30"
                    Caption ="صدور حكم كارگزيني"
                    FontName ="Titr"
                End
            End
        End
        Begin Section
            SpecialEffect =1
            Height =6406
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =8347
                    Top =113
                    Width =906
                    Height =315
                    ColumnWidth =945
                    ColumnOrder =0
                    Name ="HONUM"
                    ControlSource ="HONUM"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =9669
                            Top =113
                            Width =870
                            Height =300
                            Name ="Label1"
                            Caption ="شماره حكم"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =2879
                    Top =120
                    Width =4761
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
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =7690
                            Top =120
                            Width =525
                            Height =300
                            Name ="Label22"
                            Caption ="پرسنل"
                        End
                    End
                End
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =2
                    Left =2773
                    Top =3525
                    Width =3675
                    Height =405
                    Name ="Label23"
                    Caption ="توضيح : تمام مبالغ  بجز حق خواربار و مسكن و حق اولاد بر اساس روز محاسبه مي گردد"
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =6600
                    Top =1110
                    Width =2016
                    Height =315
                    ColumnWidth =1590
                    TabIndex =3
                    BackColor =16777164
                    ForeColor =10040115
                    Name ="SALARY_DAYLYB"
                    ControlSource ="SALARY_DAYLYB"
                    Format ="#,##0.00;-#,##0.00"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =8709
                            Top =1117
                            Width =1515
                            Height =300
                            Name ="Label15"
                            Caption ="حقوق روزانه رسمي"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =6600
                    Top =1837
                    Width =2016
                    Height =315
                    ColumnWidth =1620
                    TabIndex =5
                    Name ="HOME"
                    ControlSource ="HOME"
                    Format ="#,##0.00;-#,##0.00"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =8711
                            Top =1837
                            Width =1545
                            Height =300
                            Name ="Label17"
                            Caption ="حق خواربار و مسكن"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =6600
                    Top =2197
                    Width =2016
                    Height =315
                    TabIndex =6
                    Name ="CHILDREN"
                    ControlSource ="CHILDREN"
                    Format ="#,##0.00;-#,##0.00"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =8711
                            Top =2197
                            Width =705
                            Height =300
                            Name ="Label19"
                            Caption ="حق اولاد"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =6600
                    Top =2557
                    Width =2016
                    Height =315
                    TabIndex =7
                    Name ="CONDITIONS"
                    ControlSource ="CONDITIONS"
                    Format ="#,##0.00;-#,##0.00"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =8709
                            Top =2557
                            Width =1260
                            Height =300
                            Name ="Label21"
                            Caption ="شرايط محيط كار"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =6600
                    Top =2917
                    Width =2016
                    Height =315
                    TabIndex =8
                    Name ="BON"
                    ControlSource ="BON"
                    Format ="#,##0.00;-#,##0.00"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =8699
                            Top =2910
                            Width =690
                            Height =300
                            Name ="Label24"
                            Caption ="حق نهار"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =6600
                    Top =3277
                    Width =2016
                    Height =315
                    TabIndex =9
                    Name ="JAZB"
                    ControlSource ="JAZB"
                    Format ="#,##0.00;-#,##0.00"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =8701
                            Top =3277
                            Width =690
                            Height =300
                            Name ="Label25"
                            Caption ="حق جذب"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =6600
                    Top =3637
                    Width =2016
                    Height =315
                    TabIndex =10
                    Name ="SAYER"
                    ControlSource ="SAYER"
                    Format ="#,##0.00;-#,##0.00"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =8701
                            Top =3637
                            Width =390
                            Height =300
                            Name ="Label27"
                            Caption ="ساير"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =6600
                    Top =1477
                    Width =2016
                    Height =315
                    ColumnWidth =1560
                    TabIndex =4
                    BackColor =13434828
                    ForeColor =10040115
                    Name ="SALARY_DAYLY"
                    ControlSource ="SALARY_DAYLY"
                    Format ="#,##0.00;-#,##0.00"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =8706
                            Top =1484
                            Width =1515
                            Height =300
                            Name ="Label61"
                            Caption ="حقوق روزانه اسمي"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =6600
                    Top =735
                    Width =2016
                    Height =315
                    TabIndex =2
                    Name ="HDATE"
                    ControlSource ="HDATE"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =8711
                            Top =742
                            Width =1515
                            Height =300
                            Name ="Label28"
                            Caption ="تاريخ حكم"
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    DecimalPlaces =0
                    SpecialEffect =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2976
                    Top =2224
                    Width =2016
                    Height =315
                    TabIndex =14
                    BackColor =13434828
                    ForeColor =10040115
                    Name ="jam"
                    ControlSource ="=[SAYER]+[SALARY_DAYLY]+[HOME]+[CHILDREN]+[CONDITIONS]+[BON]+[JAZB]"
                    Format ="#,##0.00;-#,##0.00"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =5091
                            Top =2220
                            Width =1260
                            Height =240
                            Name ="Label32"
                            Caption ="جمع حكم اسمي"
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    DecimalPlaces =0
                    SpecialEffect =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2976
                    Top =1879
                    Width =2016
                    Height =315
                    TabIndex =15
                    BackColor =16777164
                    ForeColor =10040115
                    Name ="Text33"
                    ControlSource ="=[SAYER]+[SALARY_DAYLYB]+[HOME]+[CHILDREN]+[CONDITIONS]+[BON]+[JAZB]"
                    Format ="#,##0.00;-#,##0.00"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =5091
                            Top =1875
                            Width =1275
                            Height =240
                            Name ="Label34"
                            Caption ="جمع حكم رسمي"
                        End
                    End
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    DecimalPlaces =0
                    NumeralShapes =2
                    ColumnCount =2
                    Left =2849
                    Top =750
                    Width =2376
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
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =5275
                            Top =757
                            Width =1200
                            Height =300
                            Name ="Label38"
                            Caption ="مدرك تحصيلي"
                        End
                    End
                End
                Begin OptionGroup
                    OverlapFlags =85
                    Left =2849
                    Top =1230
                    Width =2372
                    Height =553
                    TabIndex =13
                    Name ="Frame41"
                    ControlSource ="TAHOL"
                    DefaultValue ="2"
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =215
                            TextFontCharSet =0
                            Left =4551
                            Top =1077
                            Width =450
                            Height =300
                            BackColor =-2147483633
                            Name ="Label42"
                            Caption ="تاهل"
                        End
                        Begin OptionButton
                            OverlapFlags =87
                            ReadingOrder =1
                            Left =3524
                            Top =1408
                            OptionValue =2
                            Name ="Option46"
                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    TextFontCharSet =0
                                    Left =3104
                                    Top =1380
                                    Width =450
                                    Height =300
                                    Name ="Label47"
                                    Caption ="مجرد"
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =87
                            ReadingOrder =1
                            Left =4534
                            Top =1431
                            OptionValue =1
                            Name ="Option44"
                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    TextFontCharSet =0
                                    Left =4024
                                    Top =1403
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
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2969
                    Top =2664
                    Width =2016
                    Height =315
                    TabIndex =16
                    Name ="MABEZ"
                    ControlSource ="MABEZ"
                    Format ="#,##0.00;-#,##0.00"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =5076
                            Top =2670
                            Width =960
                            Height =240
                            Name ="Label49"
                            Caption ="نرخ اضافه كار"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2969
                    Top =3099
                    Width =2016
                    Height =315
                    TabIndex =17
                    Name ="MABEZT"
                    ControlSource ="MABEZT"
                    Format ="#,##0.00;-#,##0.00"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =5069
                            Top =3105
                            Width =1500
                            Height =240
                            Name ="Label51"
                            Caption ="نرخ اضافه كارتعطيلي"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =6599
                    Top =4005
                    Width =2016
                    Height =315
                    TabIndex =11
                    Name ="HSHIFT"
                    ControlSource ="HSHIFT"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =8699
                            Top =4005
                            Width =2370
                            Height =240
                            Name ="Label53"
                            Caption ="درصد حق شيفت،نوبت،شب كاري"
                        End
                    End
                End
                Begin CheckBox
                    Locked = NotDefault
                    OverlapFlags =85
                    ReadingOrder =1
                    Left =10823
                    Top =224
                    Width =340
                    Height =218
                    TabIndex =18
                    Name ="OKF"
                    ControlSource ="OKF"
                    DefaultValue ="False"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =11213
                            Top =170
                            Width =975
                            Height =240
                            ForeColor =255
                            Name ="Label146"
                            Caption ="تائيد حكم"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =1
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =6359
                    Top =5333
                    Width =1431
                    Height =345
                    FontSize =10
                    FontWeight =700
                    TabIndex =19
                    BackColor =13408767
                    Name ="UP_DATE"
                    ControlSource ="UP_DATE"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =7891
                            Top =5330
                            Width =900
                            Height =240
                            Name ="Label148"
                            Caption ="تاريخ اصلاح:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =1
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =255
                    Top =5955
                    Width =2841
                    Height =345
                    FontSize =10
                    FontWeight =700
                    TabIndex =20
                    BackColor =13408767
                    Name ="UP_TIME"
                    ControlSource ="UP_TIME"
                    Format ="General Date"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3212
                            Top =5957
                            Width =1050
                            Height =240
                            Name ="Label149"
                            Caption ="ساعت اصلاح:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =300
                    Top =5375
                    Height =345
                    TabIndex =21
                    BackColor =13408767
                    Name ="UP_USER_NAME"
                    ControlSource ="UP_USER_NAME"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2052
                            Top =5375
                            Width =1290
                            Height =240
                            Name ="Label150"
                            Caption ="كاربر اصلاح كننده:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =1
                    TextFontFamily =18
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =8953
                    Top =5380
                    Width =1431
                    Height =345
                    FontSize =10
                    FontWeight =700
                    TabIndex =22
                    BackColor =13408767
                    Name ="PC_NAME"
                    ControlSource ="PC_NAME"
                    FontName ="Times New Roman"
                    InputMask ="##/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =10484
                            Top =5375
                            Width =1755
                            Height =240
                            Name ="Label152"
                            Caption ="نام كامپيوتر اصلاح كننده:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =1
                    TextFontFamily =18
                    ReadingOrder =2
                    KeyboardLanguage =11
                    ScrollBarAlign =2
                    IMESentenceMode =3
                    Left =3810
                    Top =5325
                    Width =1929
                    Height =345
                    FontSize =12
                    FontWeight =700
                    TabIndex =23
                    BackColor =13408767
                    Name ="IPADD"
                    ControlSource ="IPADD"
                    FontName ="Times New Roman"
                    InputMask ="##/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3461
                            Top =5377
                            Width =285
                            Height =240
                            Name ="Label155"
                            Caption ="IP:"
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
' See "TR_PHOKM.cls"
