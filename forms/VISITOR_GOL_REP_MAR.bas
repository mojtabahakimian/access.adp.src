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
    Width =15030
    DatasheetFontHeight =10
    ItemSuffix =56
    Left =4365
    Top =210
    Right =19395
    Bottom =9825
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xf7065f484e9fe240
    End
    Caption ="عملكرد هر ويزيتور به تفكيك ماه و كالا"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial (Arabic)"
    AllowDatasheetView =0
    Begin
        Begin Label
            BackStyle =0
            TextFontCharSet =178
            FontName ="Tahoma"
        End
        Begin CommandButton
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
            Height =9637
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin OptionGroup
                    OverlapFlags =85
                    Left =105
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
                            Left =143
                            Top =120
                            Width =1230
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
                            Left =1373
                            Top =120
                            Width =1305
                            Height =405
                            FontWeight =700
                            OptionValue =2
                            ForeColor =32768
                            Name ="Toggle17"
                            Caption ="ارديبهشت"
                        End
                        Begin ToggleButton
                            OverlapFlags =95
                            TextFontCharSet =0
                            Left =13705
                            Top =120
                            Width =1230
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
                            Left =2678
                            Top =124
                            Width =1230
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
                            Left =3908
                            Top =124
                            Width =1230
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
                            Left =5138
                            Top =124
                            Width =1230
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
                            Left =6368
                            Top =124
                            Width =1230
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
                            Left =7598
                            Top =124
                            Width =1230
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
                            Left =8828
                            Top =124
                            Width =1230
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
                            Left =10058
                            Top =124
                            Width =1230
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
                            Left =11288
                            Top =124
                            Width =1230
                            Height =405
                            FontWeight =700
                            OptionValue =10
                            ForeColor =16711680
                            Name ="Toggle25"
                            Caption ="دي"
                        End
                        Begin ToggleButton
                            OverlapFlags =247
                            TextFontCharSet =0
                            Left =12520
                            Top =124
                            Width =1230
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
                    Left =113
                    Top =1193
                    Width =14790
                    Height =8265
                    TabIndex =1
                    Name ="Visit_gol_dtl_REP"
                    SourceObject ="Form.Visit_gol_dtl_REP"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =8721
                    Top =720
                    Width =5556
                    Height =317
                    TabIndex =2
                    BackColor =13434828
                    ForeColor =8388608
                    Name ="HES"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT visitgol_head.HES, CUST_HESAB.NAME FROM visitgol_head INNER JOIN CUST_HES"
                        "AB ON visitgol_head.HES = CUST_HESAB.hes GROUP BY visitgol_head.HES, CUST_HESAB."
                        "NAME"
                    ColumnWidths ="0;3969"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =14361
                            Top =720
                            Width =510
                            Height =240
                            ForeColor =8388608
                            Name ="Label19"
                            Caption ="ويزيتور"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    AutoExpand = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =215
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5895
                    Top =720
                    Width =2766
                    Height =317
                    TabIndex =3
                    BackColor =13434828
                    ForeColor =8388608
                    Name ="HES2"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT HES FROM visitgol_head GROUP BY HES"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =93
                            TextFontCharSet =0
                            Left =7305
                            Top =720
                            Width =735
                            Height =317
                            ForeColor =8388608
                            Name ="Label18"
                            Caption ="كد ويزيتور"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =120
                    Top =705
                    Width =1530
                    Height =403
                    TabIndex =4
                    Name ="Command55"
                    Caption ="چاپ"
                    OnClick ="[Event Procedure]"
                End
            End
        End
    End
End
CodeBehindForm
' See "VISITOR_GOL_REP_MAR.cls"
