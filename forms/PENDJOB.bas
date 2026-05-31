Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    DividingLines = NotDefault
    DefaultView =0
    AllowUpdating =4
    ScrollBars =0
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =11622
    RowHeight =285
    ItemSuffix =49
    Left =8460
    Top =2130
    Right =20085
    Bottom =8715
    DatasheetGridlinesColor =12632256
    MaxRecords =100000
    RecSrcDt = Begin
        0xcf643a366df0e240
    End
    RecordSource ="PENDJOB"
    Caption ="تسويه پرسنل"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Tahoma"
    OnError ="[Event Procedure]"
    RecordSourceQualifier ="dbo"
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
        End
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
            Width =850
            Height =850
        End
        Begin Line
            Width =1701
        End
        Begin CommandButton
            Width =1701
            Height =283
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Tahoma"
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
        Begin ComboBox
            SpecialEffect =2
            Width =1701
            LabelX =-1701
            FontName ="Tahoma"
        End
        Begin FormHeader
            Height =793
            BackColor =8421376
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =93
                    TextFontCharSet =178
                    TextAlign =2
                    TextFontFamily =2
                    Left =4886
                    Top =15
                    Width =1665
                    Height =675
                    FontSize =18
                    FontWeight =700
                    ForeColor =16777215
                    Name ="Label33"
                    Caption ="تسويه حساب"
                    FontName ="Titr"
                End
                Begin Label
                    OverlapFlags =215
                    TextFontCharSet =178
                    TextAlign =2
                    TextFontFamily =2
                    Left =4886
                    Width =1665
                    Height =675
                    FontSize =18
                    FontWeight =700
                    ForeColor =128
                    Name ="Label30"
                    Caption ="تسويه حساب"
                    FontName ="Titr"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =1757
                    Top =226
                    Height =402
                    Name ="Command41"
                    Caption ="محاسبه و انتقال مبالغ"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =45
                    Top =225
                    Height =402
                    TabIndex =1
                    Name ="Command44"
                    Caption ="چاپ برگه تسويه"
                    OnClick ="[Event Procedure]"
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6750
                    Top =56
                    Width =1418
                    Height =280
                    TabIndex =2
                    Name ="MINSAL"
                    Format ="#,###"
                    DefaultValue ="73200"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =8231
                            Top =60
                            Width =2925
                            Height =240
                            Name ="Label4"
                            Caption ="حداقل دستمزد طبق قانون كار را وارد كنيد:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6746
                    Top =452
                    Width =1418
                    Height =280
                    TabIndex =3
                    Name ="MAXEYDY"
                    ControlSource ="=[MINSAL]*90"
                    Format ="#,###"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =8223
                            Top =446
                            Width =1530
                            Height =300
                            Name ="Label6"
                            Caption ="حد اكثر عيدي پاداش:"
                        End
                    End
                End
            End
        End
        Begin Section
            Height =5555
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1065
                    Top =390
                    Height =284
                    ColumnWidth =975
                    TabIndex =1
                    BackColor =13434828
                    Name ="DATE_TS"
                    ControlSource ="DATE_TS"
                    DefaultValue ="=FARSIDATE(Now())"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =3166
                            Top =390
                            Width =870
                            Height =240
                            Name ="Label3"
                            Caption ="تاريخ تسويه"
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
                    Left =7053
                    Top =1247
                    Height =284
                    ColumnWidth =1365
                    TabIndex =2
                    Name ="DATESTA"
                    ControlSource ="DATESTA"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =9154
                            Top =1247
                            Width =1515
                            Height =284
                            Name ="Label5"
                            Caption ="تاريخ شروع به كار"
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
                    Left =1044
                    Top =1247
                    Height =284
                    ColumnWidth =1050
                    TabIndex =3
                    Name ="DATEEND"
                    ControlSource ="DATEEND"
                    AfterUpdate ="[Event Procedure]"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3145
                            Top =1247
                            Width =945
                            Height =240
                            Name ="Label7"
                            Caption ="تاريخ پايان كار"
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
                    Left =1044
                    Top =1700
                    Height =284
                    TabIndex =5
                    Name ="LASTSAB"
                    ControlSource ="LASTSAB"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3145
                            Top =1700
                            Width =1485
                            Height =284
                            Name ="Label9"
                            Caption ="سابقه تسويه قبلي"
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
                    Left =7052
                    Top =2160
                    Height =284
                    TabIndex =6
                    Name ="DASTM"
                    ControlSource ="DASTM"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =9153
                            Top =2160
                            Width =1350
                            Height =284
                            Name ="Label11"
                            Caption ="دستمزد"
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
                    Left =1050
                    Top =2130
                    Height =284
                    TabIndex =7
                    Name ="MANDMORA"
                    ControlSource ="MANDMORA"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3151
                            Top =2130
                            Width =1725
                            Height =284
                            Name ="Label13"
                            Caption ="مانده مرخصي (دقيقه)"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =93
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =7053
                    Top =2948
                    Height =284
                    TabIndex =8
                    BackColor =16777164
                    Name ="EYDY"
                    ControlSource ="EYDY"
                    Format ="#,##0.00\" ريال \";-#,##0.00\" ريال \""
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =9154
                            Top =2948
                            Width =1215
                            Height =284
                            Name ="Label15"
                            Caption ="عيدي"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =93
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =7053
                    Top =3288
                    Height =284
                    ColumnWidth =1275
                    TabIndex =9
                    BackColor =16777164
                    Name ="BON"
                    ControlSource ="BON"
                    Format ="#,##0.00\" ريال \";-#,##0.00\" ريال \""
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =9154
                            Top =3288
                            Width =1155
                            Height =284
                            Name ="Label17"
                            Caption ="بن"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =93
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =7053
                    Top =3628
                    Height =284
                    ColumnWidth =1620
                    TabIndex =10
                    BackColor =16777164
                    Name ="MANDMO"
                    ControlSource ="MANDMO"
                    Format ="#,##0.00\" ريال \";-#,##0.00\" ريال \""
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =9154
                            Top =3628
                            Width =1515
                            Height =284
                            Name ="Label19"
                            Caption ="مانده مرخصي(ريالي)"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =93
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =7053
                    Top =3968
                    Height =284
                    ColumnWidth =1425
                    TabIndex =11
                    BackColor =16777164
                    Name ="HSANAVAT"
                    ControlSource ="HSANAVAT"
                    Format ="#,##0.00\" ريال \";-#,##0.00\" ريال \""
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =9154
                            Top =3968
                            Width =1635
                            Height =284
                            Name ="Label21"
                            Caption ="حق سنوات"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =93
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =7053
                    Top =4309
                    Height =284
                    TabIndex =12
                    BackColor =16777164
                    Name ="BEST"
                    ControlSource ="BEST"
                    Format ="#,##0.00\" ريال \";-#,##0.00\" ريال \""
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =9154
                            Top =4309
                            Width =1200
                            Height =284
                            Name ="Label23"
                            Caption ="بستانكاري قبلي"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =93
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1879
                    Top =3004
                    Height =284
                    TabIndex =13
                    BackColor =16777164
                    Name ="BEDH"
                    ControlSource ="BEDH"
                    Format ="#,##0.00\" ريال \";-#,##0.00\" ريال \""
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =3980
                            Top =3004
                            Width =1230
                            Height =284
                            Name ="Label25"
                            Caption ="بدهكاري قبلي"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =93
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1875
                    Top =3375
                    Height =284
                    TabIndex =14
                    BackColor =16777164
                    Name ="MALIATE"
                    ControlSource ="MALIATE"
                    Format ="#,##0.00\" ريال \";-#,##0.00\" ريال \""
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =3976
                            Top =3375
                            Width =1485
                            Height =284
                            Name ="Label27"
                            Caption ="ماليات عيدي"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =93
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1875
                    Top =3780
                    Height =284
                    TabIndex =15
                    BackColor =16777164
                    Name ="VAM"
                    ControlSource ="VAM"
                    Format ="#,##0.00\" ريال \";-#,##0.00\" ريال \""
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =3976
                            Top =3780
                            Width =1155
                            Height =284
                            Name ="Label29"
                            Caption ="مانده وام"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =5749
                    Top =396
                    Width =4761
                    Height =315
                    ColumnWidth =3150
                    BackColor =13434828
                    Name ="CODE"
                    ControlSource ="CODE"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT CODE, PNAME + N' ' + PFAMILY + N' ' + RTRIM(CAST(CODE AS NVARCHAR)) AS PE"
                        "R FROM PERSONEL ORDER BY PNAME + N' ' + PFAMILY + N' ' + RTRIM(CAST(CODE AS NVAR"
                        "CHAR))"
                    ColumnWidths ="0"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =10560
                            Top =396
                            Width =525
                            Height =300
                            Name ="Label22"
                            Caption ="پرسنل"
                        End
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =7067
                    Top =1695
                    Height =284
                    ColumnWidth =930
                    TabIndex =4
                    Name ="DAYS"
                    ControlSource ="DAYS"
                    DefaultValue ="=diff([datesta],[dateend])"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =9168
                            Top =1695
                            Width =1485
                            Height =284
                            Name ="Label32"
                            Caption ="سابقه(روز)"
                        End
                    End
                End
                Begin Rectangle
                    OverlapFlags =255
                    Left =930
                    Top =2834
                    Width =9921
                    Height =2266
                    Name ="Box34"
                End
                Begin Line
                    LineSlant = NotDefault
                    OverlapFlags =127
                    Left =915
                    Top =4650
                    Width =9921
                    Name ="Line35"
                End
                Begin Line
                    OverlapFlags =119
                    Left =6017
                    Top =2834
                    Width =0
                    Height =2268
                    Name ="Line36"
                End
                Begin TextBox
                    TabStop = NotDefault
                    DecimalPlaces =0
                    SpecialEffect =0
                    OverlapFlags =247
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =7053
                    Top =4713
                    Height =284
                    TabIndex =16
                    BackColor =10079487
                    Name ="DAR"
                    ControlSource ="=[EYDY]+[BON]+[BEST]+[MANDMO]+[HSANAVAT]"
                    Format ="#,##0.00\" ريال \";-#,##0.00\" ريال \""
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =9154
                            Top =4713
                            Width =1200
                            Height =284
                            Name ="Label38"
                            Caption ="جمع درآمدها"
                        End
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    DecimalPlaces =0
                    SpecialEffect =0
                    OverlapFlags =247
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1879
                    Top =4705
                    Height =284
                    TabIndex =17
                    BackColor =10079487
                    Name ="KOSU"
                    ControlSource ="=[BEDH]+[MALIATE]+[VAM]"
                    Format ="#,##0.00\" ريال \";-#,##0.00\" ريال \""
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =3980
                            Top =4705
                            Width =1200
                            Height =284
                            Name ="Label40"
                            Caption ="جمع كسورات"
                        End
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    DecimalPlaces =0
                    SpecialEffect =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1875
                    Top =5175
                    Height =284
                    ColumnWidth =1575
                    TabIndex =18
                    BackColor =52479
                    Name ="Text42"
                    ControlSource ="=[DAR]-[kosu]"
                    Format ="#,##0.00\" ريال \";-#,##0.00\" ريال \""
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3976
                            Top =5175
                            Width =1200
                            Height =284
                            Name ="Label43"
                            Caption ="خالص تسويه"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =2
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1050
                    Top =2490
                    Height =284
                    TabIndex =19
                    BackColor =12632256
                    Name ="Text45"
                    ControlSource ="=[MANDMORA]/440"
                    Format ="Standard"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3151
                            Top =2490
                            Width =1725
                            Height =284
                            Name ="Label46"
                            Caption ="مانده مرخصي (روز)"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    FELineBreak = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =178
                    IMESentenceMode =3
                    Left =56
                    Top =3
                    ColumnWidth =2070
                    TabIndex =20
                    Name ="CRT"
                    ControlSource ="CRT"
                    DefaultValue ="Now()"
                    AsianLineBreak =0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =178
                            Left =2261
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
                    FELineBreak = NotDefault
                    OverlapFlags =247
                    TextFontCharSet =178
                    IMESentenceMode =3
                    Top =283
                    TabIndex =21
                    Name ="UID"
                    ControlSource ="UID"
                    DefaultValue ="[Forms]![BASEKNOW]![USERCOD]"
                    AsianLineBreak =0
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextFontCharSet =178
                            Left =2205
                            Top =285
                            Width =1365
                            Height =240
                            Name ="Label36"
                            Caption ="کد کاربر ايجاد کننده"
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
' See "PENDJOB.cls"
