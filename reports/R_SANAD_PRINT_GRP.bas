Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    FilterOn = NotDefault
    DefaultView =0
    AllowUpdating =4
    TabularCharSet =178
    TabularFamily =42
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =10845
    DatasheetFontHeight =10
    ItemSuffix =231
    Left =1260
    Top =810
    DatasheetGridlinesColor =12632256
    OnPage ="[Event Procedure]"
    RecSrcDt = Begin
        0x86b8e9288e68e540
    End
    RecordSource ="DEAD_HEAD_PRINT3"
    Caption ="سند حسابداري"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial (Arabic)"
    OnError ="[Event Procedure]"
    Begin
        Begin Label
            BackStyle =0
            TextFontCharSet =178
            TextAlign =1
            TextFontFamily =42
            FontSize =9
            FontWeight =700
            ForeColor =128
            FontName ="Arial (Arabic)"
        End
        Begin Rectangle
            BackStyle =0
            BorderWidth =1
            Width =850
            Height =850
        End
        Begin Line
            Width =1701
            BorderColor =128
        End
        Begin Image
            OldBorderStyle =0
            PictureAlignment =2
            Width =1701
            Height =1701
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
            LabelX =230
            LabelY =-30
        End
        Begin CheckBox
            LabelX =230
            LabelY =-30
        End
        Begin OptionGroup
            BackStyle =1
            Width =1701
            Height =1701
        End
        Begin BoundObjectFrame
            Width =4536
            Height =2835
            LabelX =-1701
        End
        Begin TextBox
            OldBorderStyle =0
            TextFontCharSet =178
            TextFontFamily =42
            BackStyle =0
            Width =1701
            LabelX =-1701
            FontSize =9
            FontName ="Arial (Arabic)"
        End
        Begin ListBox
            TextFontCharSet =178
            TextFontFamily =42
            OldBorderStyle =0
            Width =1701
            Height =1417
            LabelX =-1701
            FontSize =9
            FontName ="Arial (Arabic)"
        End
        Begin ComboBox
            OldBorderStyle =0
            TextFontCharSet =178
            TextFontFamily =42
            BackStyle =0
            Width =1701
            LabelX =-1701
            FontSize =9
            FontName ="Arial (Arabic)"
        End
        Begin Subform
            OldBorderStyle =0
            Width =1701
            Height =1701
        End
        Begin UnboundObjectFrame
            Width =4536
            Height =2835
        End
        Begin Tab
            TextFontCharSet =178
            TextFontFamily =2
            Width =5103
            Height =3402
            FontName ="Arial (Arabic)"
        End
        Begin BreakLevel
            GroupFooter = NotDefault
            ControlSource ="N_S"
        End
        Begin BreakLevel
            SortOrder = NotDefault
            GroupHeader = NotDefault
            ControlSource ="grpas"
        End
        Begin BreakLevel
            SortOrder = NotDefault
            ControlSource ="joz"
        End
        Begin PageHeader
            Height =1724
            Name ="PageHeader"
            Begin
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =9630
                    Top =975
                    Width =1035
                    Height =345
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="N_S Label"
                    Caption ="شماره حساب "
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="N_S_Label"
                End
                Begin Line
                    BorderWidth =2
                    Left =15
                    Top =945
                    Width =10808
                    BorderColor =0
                    Name ="Line26"
                End
                Begin Line
                    Left =15
                    Top =1380
                    Width =10808
                    Name ="Line27"
                End
                Begin Line
                    BorderWidth =1
                    Left =9510
                    Top =979
                    Width =0
                    Height =404
                    BorderColor =0
                    Name ="Line29"
                End
                Begin Line
                    BorderWidth =2
                    Left =15
                    Top =979
                    Width =0
                    Height =404
                    BorderColor =0
                    Name ="Line36"
                End
                Begin Line
                    BorderWidth =2
                    Left =10800
                    Top =979
                    Width =0
                    Height =404
                    BorderColor =0
                    Name ="Line37"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =4351
                    Top =431
                    Width =1710
                    Height =540
                    FontSize =14
                    FontWeight =400
                    ForeColor =0
                    Name ="Label22"
                    Caption ="سند حسابداري "
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =2310
                    Top =960
                    Width =915
                    Height =345
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="Label80"
                    Caption ="مبلغ بدهكار"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =1
                    Left =3735
                    Top =979
                    Width =0
                    Height =404
                    BorderColor =0
                    Name ="Line81"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =6525
                    Top =975
                    Width =1725
                    Height =375
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="Label90"
                    Caption ="نام حساب  /   شرح عمليات "
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextFontFamily =2
                    Left =9808
                    Top =495
                    Width =975
                    Height =360
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="Label92"
                    Caption ="تاريخ سنـد:"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =1
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =8325
                    Top =495
                    Width =1446
                    Height =420
                    FontSize =10
                    Name ="Text93"
                    ControlSource ="DATE_S"
                    FontName ="WeblogmaYekan"
                    InputMask ="####/##/##"
                End
                Begin Label
                    TextFontFamily =2
                    Left =9810
                    Top =90
                    Width =975
                    Height =360
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="Label97"
                    Caption ="شماره سند:"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =1
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =8397
                    Top =95
                    Width =1371
                    Height =420
                    FontSize =10
                    TabIndex =1
                    Name ="Text98"
                    ControlSource ="N_S"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =362
                    Top =960
                    Width =930
                    Height =360
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="Label101"
                    Caption ="مبلغ بستانكار"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =1
                    Left =1890
                    Top =979
                    Width =0
                    Height =404
                    BorderColor =0
                    Name ="Line102"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =4237
                    Top =975
                    Width =600
                    Height =375
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="Label124"
                    Caption ="مبلغ جزء"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =1
                    Left =5446
                    Top =979
                    Width =0
                    Height =404
                    BorderColor =0
                    Name ="Line125"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1800
                    Width =6198
                    Height =540
                    FontSize =14
                    TabIndex =2
                    BorderColor =32768
                    Name ="Text90"
                    ControlSource ="=[Forms]![Baseknow]![WIDTH_D]"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =30
                    Top =600
                    Width =2711
                    Height =300
                    FontSize =10
                    TabIndex =3
                    Name ="Text24"
                    ControlSource ="=\"صفحه \" & [Page] & \" از \" & [Pages]"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    BorderWidth =2
                    Left =15
                    Top =1380
                    Width =0
                    Height =344
                    BorderColor =0
                    Name ="Line220"
                End
                Begin Line
                    BorderWidth =2
                    Left =10800
                    Top =1380
                    Width =0
                    Height =344
                    BorderColor =0
                    Name ="Line221"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =6720
                    Top =1400
                    Width =1140
                    Height =300
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="Label222"
                    Caption ="جمع صفحات قبل"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =1
                    KeyboardLanguage =15
                    ScrollBarAlign =1
                    NumeralShapes =2
                    Left =1905
                    Top =1380
                    Width =1799
                    Height =300
                    FontSize =10
                    TabIndex =4
                    Name ="Text223"
                    ControlSource ="=Format(Str([san]-[SumOfBED]),\"#,###\") & \"  \""
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Left =3735
                    Top =1380
                    Width =0
                    Height =344
                    BorderColor =0
                    Name ="Line224"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =1
                    KeyboardLanguage =15
                    ScrollBarAlign =1
                    NumeralShapes =2
                    Left =71
                    Top =1380
                    Width =1799
                    Height =300
                    FontSize =10
                    TabIndex =5
                    Name ="Text226"
                    ControlSource ="=Format(Str([sas]-[SumOfBES]),\"#,###\") & \"  \""
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Left =1890
                    Top =1380
                    Width =0
                    Height =344
                    BorderColor =0
                    Name ="Line227"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Left =5445
                    Top =1380
                    Width =0
                    Height =344
                    BorderColor =0
                    Name ="Line228"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Left =9510
                    Top =1380
                    Width =0
                    Height =344
                    BorderColor =0
                    Name ="Line229"
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =377
            BreakLevel =1
            OnFormat ="[Event Procedure]"
            Name ="GroupHeader0"
            Begin
                Begin Line
                    BorderWidth =1
                    Left =9510
                    Width =0
                    Height =377
                    BorderColor =0
                    Name ="Line127"
                End
                Begin Line
                    BorderWidth =2
                    Left =15
                    Width =0
                    Height =377
                    BorderColor =0
                    Name ="Line128"
                End
                Begin Line
                    BorderWidth =2
                    Left =10800
                    Width =0
                    Height =377
                    BorderColor =0
                    Name ="Line129"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =1
                    KeyboardLanguage =15
                    ScrollBarAlign =1
                    NumeralShapes =2
                    Left =1905
                    Width =1799
                    Height =330
                    FontSize =10
                    Name ="SumOfBED1"
                    ControlSource ="=Format(Str([SumOfBED]),\"#,###\") & \"  \""
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    BorderWidth =1
                    Left =3735
                    Width =0
                    Height =377
                    BorderColor =0
                    Name ="Line132"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =1
                    KeyboardLanguage =15
                    ScrollBarAlign =1
                    NumeralShapes =2
                    Left =60
                    Width =1799
                    Height =330
                    FontSize =10
                    TabIndex =1
                    Name ="SumOfBES1"
                    ControlSource ="=Format(Str([SumOfBES]),\"#,###\") & \"  \""
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    BorderWidth =1
                    Left =1890
                    Width =0
                    Height =377
                    BorderColor =0
                    Name ="Line134"
                End
                Begin TextBox
                    CanShrink = NotDefault
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5460
                    Width =4026
                    Height =330
                    FontSize =10
                    FontWeight =700
                    TabIndex =2
                    Name ="KOLNAME"
                    ControlSource ="KNAME"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    BorderWidth =1
                    Left =5445
                    Width =0
                    Height =377
                    BorderColor =0
                    Name ="Line137"
                End
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =2
                    ReadingOrder =1
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =9570
                    Width =1191
                    Height =345
                    FontSize =10
                    TabIndex =3
                    Name ="Text139"
                    ControlSource ="N_KOL"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    RunningSum =1
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =1
                    KeyboardLanguage =15
                    ScrollBarAlign =1
                    NumeralShapes =2
                    Left =4729
                    Width =1799
                    Height =330
                    FontSize =10
                    TabIndex =4
                    Name ="SAN"
                    ControlSource ="SumOfBED"
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    RunningSum =1
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =1
                    KeyboardLanguage =15
                    ScrollBarAlign =1
                    NumeralShapes =2
                    Left =3019
                    Width =1799
                    Height =330
                    FontSize =10
                    TabIndex =5
                    Name ="SAS"
                    ControlSource ="SumOfBES"
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    TextFontFamily =2
                    Left =6966
                    Top =56
                    TabIndex =6
                    Name ="SumOfBED"
                    ControlSource ="SumOfBED"
                    FontName ="WeblogmaYekan"
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            CanGrow = NotDefault
            Height =1020
            Name ="Detail"
            Begin
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TextAlign =1
                    TextFontFamily =2
                    ReadingOrder =1
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =9570
                    Width =1217
                    Height =660
                    FontSize =10
                    Name ="HES_K"
                    ControlSource ="HES"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    BorderWidth =1
                    Left =9510
                    Width =0
                    Height =1020
                    BorderColor =0
                    Name ="Line39"
                End
                Begin Line
                    BorderWidth =2
                    Left =15
                    Width =0
                    Height =1020
                    BorderColor =0
                    Name ="Line46"
                End
                Begin Line
                    BorderWidth =2
                    Left =10800
                    Width =0
                    Height =1020
                    BorderColor =0
                    Name ="Line47"
                End
                Begin TextBox
                    CanShrink = NotDefault
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5460
                    Top =345
                    Width =4011
                    Height =653
                    FontSize =10
                    TabIndex =1
                    Name ="SHARH"
                    ControlSource ="SHARH"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    BorderWidth =1
                    Left =3735
                    Width =0
                    Height =1020
                    BorderColor =0
                    Name ="Line83"
                End
                Begin Line
                    BorderWidth =1
                    Left =1890
                    Width =0
                    Height =1020
                    BorderColor =0
                    Name ="Line104"
                End
                Begin TextBox
                    CanShrink = NotDefault
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5460
                    Width =4026
                    Height =330
                    FontSize =10
                    TabIndex =2
                    Name ="JOZNAME"
                    ControlSource ="TNAME"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    BorderWidth =1
                    Left =5445
                    Width =0
                    Height =1020
                    BorderColor =0
                    Name ="Line117"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =1
                    KeyboardLanguage =15
                    ScrollBarAlign =1
                    NumeralShapes =2
                    Left =3750
                    Width =1664
                    Height =330
                    FontSize =10
                    TabIndex =3
                    Name ="Text143"
                    ControlSource ="=Format(Str([joz]),\"#,###\") & \"  \""
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    TextFontFamily =2
                    Width =771
                    TabIndex =4
                    Name ="Text185"
                    ControlSource ="=[Pages]"
                    FontName ="WeblogmaYekan"
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            ForceNewPage =2
            Height =1616
            OnFormat ="[Event Procedure]"
            Name ="GroupFooter2"
            Begin
                Begin Line
                    BorderWidth =2
                    Left =15
                    Width =0
                    Height =419
                    BorderColor =0
                    Name ="Line64"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =2
                    Left =15
                    Top =419
                    Width =10808
                    BorderColor =4210752
                    Name ="Line65"
                End
                Begin Line
                    BorderWidth =2
                    Left =10800
                    Width =0
                    Height =419
                    BorderColor =0
                    Name ="Line67"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =10368
                    Width =435
                    Height =375
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="Label68"
                    Caption ="جمع :"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =1
                    KeyboardLanguage =15
                    ScrollBarAlign =1
                    NumeralShapes =2
                    Left =1905
                    Width =1799
                    Height =375
                    FontSize =10
                    Name ="SSAN"
                    ControlSource ="=Format(Str([san]),\"#,###\") & \"  \""
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Left =3735
                    Width =0
                    Height =419
                    BorderColor =0
                    Name ="Line85"
                End
                Begin Line
                    BorderWidth =1
                    Left =15
                    Width =10808
                    Name ="Line96"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3765
                    Width =6606
                    Height =375
                    FontSize =10
                    TabIndex =1
                    Name ="Text105"
                    ControlSource ="=ALPHANUM([SSAN]) & \" \" & \"ريال\""
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =1
                    KeyboardLanguage =15
                    ScrollBarAlign =1
                    NumeralShapes =2
                    Left =71
                    Width =1799
                    Height =375
                    FontSize =10
                    TabIndex =2
                    Name ="SSAN2"
                    ControlSource ="=Format(Str([sas]),\"#,###\") & \"  \""
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Left =1890
                    Width =0
                    Height =419
                    BorderColor =0
                    Name ="Line115"
                End
                Begin BoundObjectFrame
                    OldBorderStyle =0
                    Left =60
                    Top =450
                    Width =10715
                    Height =1166
                    TabIndex =3
                    Name ="OLEBound157"
                    ControlSource ="=[Forms]![Baseknow]![EMZA]"
                End
            End
        End
        Begin PageFooter
            Height =419
            OnFormat ="[Event Procedure]"
            Name ="PageFooter"
            Begin
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =2
                    Width =10808
                    BorderColor =4210752
                    Name ="Line179"
                End
                Begin Line
                    BorderWidth =2
                    Left =15
                    Width =0
                    Height =419
                    BorderColor =0
                    Name ="Line205"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =2
                    Top =419
                    Width =10808
                    BorderColor =4210752
                    Name ="Line206"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =8730
                    Width =1455
                    Height =375
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="Label208"
                    Caption ="جمع نقل به صفحه بعد :"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =1
                    KeyboardLanguage =15
                    ScrollBarAlign =1
                    NumeralShapes =2
                    Left =1920
                    Width =1799
                    Height =375
                    FontSize =10
                    Name ="Text209"
                    ControlSource ="=Format(Str([san]),\"#,###\") & \"  \""
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Left =3750
                    Width =0
                    Height =419
                    BorderColor =0
                    Name ="Line210"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =1
                    KeyboardLanguage =15
                    ScrollBarAlign =1
                    NumeralShapes =2
                    Left =86
                    Width =1799
                    Height =375
                    FontSize =10
                    TabIndex =1
                    Name ="Text212"
                    ControlSource ="=Format(Str([sas]),\"#,###\") & \"  \""
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Left =1905
                    Width =0
                    Height =419
                    BorderColor =0
                    Name ="Line213"
                End
                Begin Line
                    BorderWidth =2
                    Left =10785
                    Width =0
                    Height =419
                    BorderColor =0
                    Name ="Line230"
                End
            End
        End
    End
End
CodeBehindForm
' See "R_SANAD_PRINT_GRP.cls"
