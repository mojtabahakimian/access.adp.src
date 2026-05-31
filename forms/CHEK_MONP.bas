Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    DefaultView =0
    AllowUpdating =4
    ScrollBars =0
    ViewsAllowed =1
    TabularCharSet =178
    TabularFamily =42
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =15023
    DatasheetFontHeight =10
    ItemSuffix =50
    Left =1770
    Right =16545
    Bottom =8670
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xf7065f484e9fe240
    End
    Caption ="دفتر ماهانه چكهاي پرداختي"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial (Arabic)"
    OnError ="[Event Procedure]"
    AllowDatasheetView =0
    Begin
        Begin Label
            BackStyle =0
            TextFontCharSet =178
            FontName ="Tahoma"
        End
        Begin OptionButton
            SpecialEffect =2
            LabelX =230
            LabelY =-30
        End
        Begin OptionGroup
            SpecialEffect =3
            Width =1701
            Height =1701
        End
        Begin TextBox
            SpecialEffect =2
            OldBorderStyle =0
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
        Begin ToggleButton
            TextFontCharSet =178
            Width =283
            Height =283
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Tahoma"
        End
        Begin Tab
            TextFontCharSet =178
            Width =5103
            Height =3402
            FontName ="Tahoma"
        End
        Begin Page
            Width =1701
            Height =1701
        End
        Begin Section
            CanGrow = NotDefault
            Height =8674
            BackColor =10079487
            Name ="Detail"
            Begin
                Begin OptionGroup
                    OverlapFlags =93
                    Left =112
                    Top =45
                    Width =14907
                    Height =591
                    Name ="Frame13"
                    DefaultValue ="1"
                    OnClick ="[Event Procedure]"
                    Begin
                        Begin ToggleButton
                            OverlapFlags =87
                            TextFontCharSet =0
                            Left =150
                            Top =120
                            Width =1140
                            Height =405
                            FontWeight =700
                            OptionValue =1
                            ForeColor =32768
                            Name ="Toggle16"
                            Caption ="فروردين"
                        End
                        Begin ToggleButton
                            OverlapFlags =87
                            TextFontCharSet =0
                            Left =1290
                            Top =120
                            Width =1215
                            Height =405
                            FontWeight =700
                            OptionValue =2
                            ForeColor =32768
                            Name ="Toggle17"
                            Caption ="ارديبهشت"
                        End
                        Begin ToggleButton
                            OverlapFlags =87
                            TextFontCharSet =0
                            Left =12819
                            Top =120
                            Width =1140
                            Height =405
                            FontWeight =700
                            OptionValue =12
                            ForeColor =16711680
                            Name ="Toggle27"
                            Caption ="اسفند"
                        End
                        Begin ToggleButton
                            OverlapFlags =87
                            TextFontCharSet =0
                            Left =2505
                            Top =124
                            Width =1140
                            Height =405
                            FontWeight =700
                            OptionValue =3
                            ForeColor =32768
                            Name ="Toggle18"
                            Caption ="خرداد"
                        End
                        Begin ToggleButton
                            OverlapFlags =87
                            TextFontCharSet =0
                            Left =3645
                            Top =124
                            Width =1140
                            Height =405
                            FontWeight =700
                            OptionValue =4
                            ForeColor =13209
                            Name ="Toggle19"
                            Caption ="تير"
                        End
                        Begin ToggleButton
                            OverlapFlags =87
                            TextFontCharSet =0
                            Left =4785
                            Top =124
                            Width =1140
                            Height =405
                            FontWeight =700
                            OptionValue =5
                            ForeColor =13209
                            Name ="Toggle20"
                            Caption ="مرداد"
                        End
                        Begin ToggleButton
                            OverlapFlags =87
                            TextFontCharSet =0
                            Left =5925
                            Top =124
                            Width =1140
                            Height =405
                            FontWeight =700
                            OptionValue =6
                            ForeColor =13209
                            Name ="Toggle21"
                            Caption ="شهريور"
                        End
                        Begin ToggleButton
                            OverlapFlags =87
                            TextFontCharSet =0
                            Left =7080
                            Top =124
                            Width =1140
                            Height =405
                            FontWeight =700
                            OptionValue =7
                            ForeColor =39423
                            Name ="Toggle22"
                            Caption ="مهر"
                        End
                        Begin ToggleButton
                            OverlapFlags =87
                            TextFontCharSet =0
                            Left =8220
                            Top =124
                            Width =1140
                            Height =405
                            FontWeight =700
                            OptionValue =8
                            ForeColor =39423
                            Name ="Toggle23"
                            Caption ="آبان"
                        End
                        Begin ToggleButton
                            OverlapFlags =87
                            TextFontCharSet =0
                            Left =9375
                            Top =124
                            Width =1140
                            Height =405
                            FontWeight =700
                            OptionValue =9
                            ForeColor =39423
                            Name ="Toggle24"
                            Caption ="آذر"
                        End
                        Begin ToggleButton
                            OverlapFlags =87
                            TextFontCharSet =0
                            Left =10515
                            Top =124
                            Width =1140
                            Height =405
                            FontWeight =700
                            OptionValue =10
                            ForeColor =16711680
                            Name ="Toggle25"
                            Caption ="دي"
                        End
                        Begin ToggleButton
                            OverlapFlags =87
                            TextFontCharSet =0
                            Left =11670
                            Top =124
                            Width =1140
                            Height =405
                            FontWeight =700
                            OptionValue =11
                            ForeColor =16711680
                            Name ="Toggle26"
                            Caption ="بهمن"
                        End
                    End
                End
                Begin Subform
                    OverlapFlags =85
                    Left =114
                    Top =737
                    Width =14820
                    Height =7425
                    TabIndex =1
                    Name ="PAY_GETP_mon"
                    SourceObject ="Form.PAY_GETP_mon"
                End
                Begin TextBox
                    OverlapFlags =215
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =14002
                    Top =180
                    Width =563
                    Height =330
                    TabIndex =2
                    Name ="Text35"
                    Format ="0000"
                    ValidationRule ="Is Not Null"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="=Left(Trim(Str(FARSIDATE(Date()))),4)"
                    InputMask ="####"
                    Begin
                        Begin Label
                            OverlapFlags =223
                            TextFontCharSet =0
                            TextAlign =2
                            Left =14566
                            Top =180
                            Width =450
                            Height =300
                            Name ="Label36"
                            Caption ="سال:"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =960
                    Top =8276
                    Width =2691
                    Height =315
                    TabIndex =3
                    BackColor =13434828
                    Name ="JJKOL"
                    ControlSource ="=IIf(IsError([PAY_GETP_mon].[Form]![SMAB]),0,[PAY_GETP_mon].[Form]![SMAB])"
                    Format ="#,##0.00\" ريال \";-#,##0.00\" ريال \""
                    Tag ="=IIf(IsError([PAY_GETD_mon].[Form]![SMAB]);0;[PAY_GETD_mon].[Form]![SMAB])"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =3750
                            Top =8276
                            Width =1260
                            Height =300
                            Name ="Label61"
                            Caption ="جمع كل :"
                        End
                    End
                End
                Begin OptionGroup
                    OverlapFlags =85
                    Left =7583
                    Top =8220
                    Width =6977
                    Height =380
                    TabIndex =4
                    Name ="grp"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="1"
                    Begin
                        Begin OptionButton
                            OverlapFlags =87
                            ReadingOrder =1
                            Left =9958
                            Top =8299
                            OptionValue =5
                            Name ="Option48"
                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    TextFontCharSet =0
                                    TextAlign =3
                                    Left =8830
                                    Top =8275
                                    Width =1140
                                    Height =240
                                    ForeColor =16737843
                                    Name ="Label49"
                                    Caption ="برگشت شده"
                                End
                            End
                        End
                        Begin OptionButton
                            Visible = NotDefault
                            OverlapFlags =87
                            ReadingOrder =1
                            Left =8423
                            Top =8303
                            OptionValue =4
                            Name ="Option46"
                            Begin
                                Begin Label
                                    Visible = NotDefault
                                    OverlapFlags =119
                                    TextFontCharSet =0
                                    Left =7583
                                    Top =8275
                                    Width =840
                                    Height =240
                                    ForeColor =16737843
                                    Name ="Label47"
                                    Caption ="واگذار شده"
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =87
                            ReadingOrder =1
                            Left =14192
                            Top =8308
                            OptionValue =1
                            Name ="Option40"
                            Begin
                                Begin Label
                                    OverlapFlags =119
                                    TextFontCharSet =0
                                    Left =13427
                                    Top =8280
                                    Width =765
                                    Height =240
                                    Name ="Label41"
                                    Caption ="همه چكها"
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =87
                            ReadingOrder =1
                            Left =11477
                            Top =8323
                            OptionValue =3
                            Name ="Option44"
                            Begin
                                Begin Label
                                    OverlapFlags =119
                                    TextFontCharSet =0
                                    Left =10637
                                    Top =8295
                                    Width =840
                                    Height =240
                                    ForeColor =16711680
                                    Name ="Label45"
                                    Caption ="وصول شده"
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =87
                            ReadingOrder =1
                            Left =12917
                            Top =8323
                            OptionValue =2
                            Name ="Option42"
                            Begin
                                Begin Label
                                    OverlapFlags =119
                                    TextFontCharSet =0
                                    Left =12017
                                    Top =8295
                                    Width =900
                                    Height =240
                                    ForeColor =10040115
                                    Name ="Label43"
                                    Caption ="وصول نشده"
                                End
                            End
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "CHEK_MONP.cls"
