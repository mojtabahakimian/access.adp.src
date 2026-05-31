Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    DividingLines = NotDefault
    FilterOn = NotDefault
    DefaultView =0
    AllowUpdating =4
    ScrollBars =0
    TabularCharSet =178
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =11115
    RowHeight =300
    ItemSuffix =66
    Left =5055
    Top =510
    Right =15915
    Bottom =6810
    DatasheetForeColor =33554432
    DatasheetGridlinesColor =12632256
    Filter ="((CODE=\"99\"))"
    RecSrcDt = Begin
        0xc1538de40c8ce240
    End
    RecordSource ="PHOKM"
    Caption ="حكم كارگزيني پرسنل"
    OnCurrent ="[Event Procedure]"
    AfterUpdate ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
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
                    Left =4165
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
                    Left =4165
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
            Height =5612
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =7218
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
                            Left =8540
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
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =1750
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
                    OnNotInList ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =6561
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
                    Left =1650
                    Top =3975
                    Width =3675
                    Height =405
                    ForeColor =255
                    Name ="Label23"
                    Caption ="توضيح : تمام مبالغ  بجز حق خواربار و مسكن و حق اولاد و بن بر اساس روز محاسبه مي "
                        "گردد"
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =6586
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
                    OnKeyPress ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =8695
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
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =6586
                    Top =1837
                    Width =2016
                    Height =315
                    ColumnWidth =1620
                    TabIndex =5
                    Name ="HOME"
                    ControlSource ="HOME"
                    Format ="#,##0.00;-#,##0.00"
                    OnKeyPress ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =8697
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
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =6586
                    Top =2197
                    Width =2016
                    Height =315
                    TabIndex =6
                    Name ="CHILDREN"
                    ControlSource ="CHILDREN"
                    Format ="#,##0.00;-#,##0.00"
                    OnKeyPress ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =8697
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
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =6586
                    Top =2557
                    Width =2016
                    Height =315
                    TabIndex =7
                    Name ="CONDITIONS"
                    ControlSource ="CONDITIONS"
                    Format ="#,##0.00;-#,##0.00"
                    OnKeyPress ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =8695
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
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =6586
                    Top =2917
                    Width =2016
                    Height =315
                    TabIndex =8
                    Name ="BON"
                    ControlSource ="BON"
                    Format ="#,##0.00;-#,##0.00"
                    OnKeyPress ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =8685
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
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =6586
                    Top =3277
                    Width =2016
                    Height =315
                    TabIndex =9
                    Name ="JAZB"
                    ControlSource ="JAZB"
                    Format ="#,##0.00;-#,##0.00"
                    OnKeyPress ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =8687
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
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =6586
                    Top =3637
                    Width =2016
                    Height =315
                    TabIndex =10
                    Name ="SAYER"
                    ControlSource ="SAYER"
                    Format ="#,##0.00;-#,##0.00"
                    OnKeyPress ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =8687
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
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =6586
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
                    OnKeyPress ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =8692
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
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =6586
                    Top =735
                    Width =2016
                    Height =315
                    TabIndex =2
                    Name ="HDATE"
                    ControlSource ="HDATE"
                    BeforeUpdate ="[Event Procedure]"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =8697
                            Top =742
                            Width =1515
                            Height =300
                            Name ="Label28"
                            Caption ="تاريخ حكم"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    Enabled = NotDefault
                    Locked = NotDefault
                    DecimalPlaces =0
                    SpecialEffect =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1853
                    Top =2674
                    Width =2016
                    Height =315
                    TabIndex =15
                    BackColor =13434828
                    ForeColor =10040115
                    Name ="jam"
                    ControlSource ="=[SAYER]+[SALARY_DAYLY]+[HOME]+[CHILDREN]+[CONDITIONS]+[BON]+[JAZB]"
                    Format ="#,##0.00;-#,##0.00"
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =3968
                            Top =2670
                            Width =1260
                            Height =240
                            Name ="Label32"
                            Caption ="جمع حكم اسمي"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    Enabled = NotDefault
                    Locked = NotDefault
                    DecimalPlaces =0
                    SpecialEffect =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1853
                    Top =2329
                    Width =2016
                    Height =315
                    TabIndex =16
                    BackColor =16777164
                    ForeColor =10040115
                    Name ="Text33"
                    ControlSource ="=[SAYER]+[SALARY_DAYLYB]+[HOME]+[CHILDREN]+[CONDITIONS]+[BON]+[JAZB]"
                    Format ="#,##0.00;-#,##0.00"
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =3968
                            Top =2325
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
                    ScrollBarAlign =2
                    DecimalPlaces =0
                    NumeralShapes =2
                    ColumnCount =2
                    Left =1720
                    Top =750
                    Width =2376
                    Height =315
                    TabIndex =13
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
                            Left =4146
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
                    Left =1720
                    Top =1230
                    Width =2372
                    Height =553
                    TabIndex =14
                    Name ="Frame41"
                    ControlSource ="TAHOL"
                    DefaultValue ="2"
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =215
                            TextFontCharSet =0
                            Left =3422
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
                            Left =2395
                            Top =1408
                            OptionValue =2
                            Name ="Option46"
                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    TextFontCharSet =0
                                    Left =1975
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
                            Left =3405
                            Top =1431
                            OptionValue =1
                            Name ="Option44"
                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    TextFontCharSet =0
                                    Left =2895
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
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1846
                    Top =3114
                    Width =2016
                    Height =315
                    TabIndex =17
                    Name ="MABEZ"
                    ControlSource ="MABEZ"
                    Format ="#,##0.00;-#,##0.00"
                    OnKeyPress ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =3953
                            Top =3120
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
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1846
                    Top =3549
                    Width =2016
                    Height =315
                    TabIndex =18
                    Name ="MABEZT"
                    ControlSource ="MABEZT"
                    Format ="#,##0.00;-#,##0.00"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =3946
                            Top =3555
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
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =6585
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
                            Left =8685
                            Top =4005
                            Width =2370
                            Height =240
                            Name ="Label53"
                            Caption ="درصد حق شيفت،نوبت،شب كاري"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =1680
                    Top =4927
                    Width =1425
                    Height =405
                    TabIndex =19
                    Name ="Command54"
                    Caption ="چاپ حكم"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    Left =3103
                    Top =4920
                    Width =1426
                    Height =405
                    TabIndex =20
                    Name ="ESLAH"
                    Caption ="اصلاح حكم"
                    OnClick ="[Event Procedure]"
                End
                Begin CheckBox
                    Locked = NotDefault
                    OverlapFlags =85
                    ReadingOrder =1
                    Left =9694
                    Top =224
                    Width =340
                    Height =218
                    TabIndex =21
                    Name ="OKF"
                    ControlSource ="OKF"
                    DefaultValue ="False"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =10084
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
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =6590
                    Top =4372
                    Width =2016
                    Height =315
                    TabIndex =12
                    Name ="bonj"
                    ControlSource ="bonj"
                    Format ="#,##0.00;-#,##0.00"
                    OnKeyPress ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =8689
                            Top =4365
                            Width =690
                            Height =300
                            Name ="Label58"
                            Caption ="بن"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    ReadingOrder =1
                    Left =3838
                    Top =1885
                    ColumnWidth =2640
                    TabIndex =22
                    Name ="ISCEO"
                    ControlSource ="ISCEO"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1245
                            Top =1860
                            Width =2535
                            Height =240
                            Name ="Label60"
                            Caption ="مدير  غير مشمول  3 % بيمه بيکاري"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =56
                    Top =3007
                    ColumnWidth =2070
                    TabIndex =23
                    Name ="CRT"
                    ControlSource ="CRT"
                    DefaultValue ="Now()"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =2261
                            Top =3004
                            Width =1515
                            Height =240
                            Name ="Label35"
                            Caption ="تاريخ ايجاد"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Top =3287
                    TabIndex =24
                    Name ="UID"
                    ControlSource ="UID"
                    DefaultValue ="[Forms]![BASEKNOW]![USERCOD]"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =2205
                            Top =3289
                            Width =1365
                            Height =240
                            Name ="Label36"
                            Caption ="کد کاربر ايجاد کننده"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =6585
                    Top =4717
                    Width =2016
                    Height =315
                    TabIndex =25
                    BackColor =10079487
                    Name ="HAGHTAHOL"
                    ControlSource ="HAGHTAHOL"
                    Format ="#,##0.00;-#,##0.00"
                    OnKeyPress ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =8684
                            Top =4710
                            Width =690
                            Height =300
                            Name ="Label65"
                            Caption ="حق تاهل"
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
' See "PHOKM.cls"
