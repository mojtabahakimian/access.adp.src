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
    ItemSuffix =237
    Left =270
    Top =825
    DatasheetGridlinesColor =12632256
    OnPage ="[Event Procedure]"
    RecSrcDt = Begin
        0x13c19b278e68e540
    End
    RecordSource ="DEAD_HEAD_PRINT3"
    Caption ="سند حسابداري"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial (Arabic)"
    RecordSourceQualifier ="dbo"
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
            Height =1670
            OnFormat ="[Event Procedure]"
            Name ="PageHeader"
            Begin
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =9630
                    Top =975
                    Width =1035
                    Height =345
                    FontSize =8
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
                    LineSlant = NotDefault
                    BorderWidth =1
                    Top =1350
                    Width =10808
                    BorderColor =0
                    Name ="Line27"
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
                    Left =2070
                    Top =960
                    Width =915
                    Height =345
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="Label80"
                    Caption ="مبلغ بدهكار"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =6525
                    Top =975
                    Width =1920
                    Height =375
                    FontSize =8
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
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="Label101"
                    Caption ="مبلغ بستانكار"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =3757
                    Top =975
                    Width =600
                    Height =375
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="Label124"
                    Caption ="مبلغ جزء"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2205
                    Width =6033
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
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =6807
                    Top =1350
                    Width =1140
                    Height =300
                    FontSize =8
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
                    Left =1669
                    Top =1370
                    Width =1559
                    Height =300
                    FontSize =8
                    TabIndex =4
                    Name ="Text223"
                    ControlSource ="=Format(Str([san]),\"#,###\") & \"  \""
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                    Tag ="=Format(Str([san]-[SumOfBED]);\"#,###\") & \"  \""
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
                    Top =1370
                    Width =1574
                    Height =300
                    FontSize =8
                    TabIndex =5
                    Name ="Text226"
                    ControlSource ="=Format(Str([sas]),\"#,###\") & \"  \""
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =331
            BreakLevel =1
            OnFormat ="[Event Procedure]"
            Name ="GroupHeader0"
            Begin
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =1
                    KeyboardLanguage =15
                    ScrollBarAlign =1
                    NumeralShapes =2
                    Left =1665
                    Width =1559
                    Height =330
                    FontSize =8
                    Name ="SumOfBED1"
                    ControlSource ="=Format(Str([SumOfBED]),\"#,###\") & \"  \""
                    Format ="#,###"
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
                    Left =60
                    Width =1559
                    Height =330
                    FontSize =8
                    TabIndex =1
                    Name ="SumOfBES1"
                    ControlSource ="=Format(Str([SumOfBES]),\"#,###\") & \"  \""
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    CanShrink = NotDefault
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =1
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4811
                    Width =4671
                    Height =330
                    FontSize =8
                    TabIndex =2
                    Name ="KOLNAME"
                    ControlSource ="KNAME"
                    FontName ="WeblogmaYekan"
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
                    Height =331
                    FontSize =8
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
                    Left =5220
                    Width =599
                    Height =330
                    FontSize =8
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
                    Left =5880
                    Width =1799
                    Height =330
                    FontSize =8
                    TabIndex =5
                    Name ="SAS"
                    ControlSource ="SumOfBES"
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    TextFontFamily =2
                    Left =6737
                    Top =56
                    FontSize =8
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
            Height =676
            Name ="Detail"
            Begin
                Begin TextBox
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
                    FontSize =8
                    Name ="HES_K"
                    ControlSource ="HES"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    BorderWidth =1
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4755
                    Top =345
                    Width =4776
                    Height =331
                    FontSize =8
                    TabIndex =1
                    Name ="SHARH"
                    ControlSource ="SHARH"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    CanShrink = NotDefault
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4755
                    Width =4716
                    Height =330
                    FontSize =8
                    TabIndex =2
                    Name ="JOZNAME"
                    ControlSource ="TNAME"
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
                    Left =3270
                    Width =1424
                    Height =330
                    FontSize =8
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
                    FontSize =8
                    TabIndex =4
                    Name ="Text185"
                    ControlSource ="=[Pages]"
                    FontName ="WeblogmaYekan"
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =56
            Name ="GroupFooter0"
            Begin
                Begin Line
                    BorderWidth =1
                    Top =15
                    Width =10830
                    BorderColor =0
                    Name ="Line236"
                End
            End
        End
        Begin PageFooter
            Height =2281
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
                    LineSlant = NotDefault
                    BorderWidth =2
                    Top =419
                    Width =10808
                    BorderColor =4210752
                    Name ="Line206"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =2
                    Left =9297
                    Width =1455
                    Height =375
                    FontSize =8
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
                    Left =1680
                    Width =1560
                    Height =375
                    FontSize =8
                    Name ="ssan"
                    ControlSource ="=Format(Str([san]),\"#,###\") & \"  \""
                    Format ="#,###"
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
                    Left =86
                    Width =1575
                    Height =375
                    FontSize =8
                    TabIndex =1
                    Name ="Text212"
                    ControlSource ="=Format(Str([sas]),\"#,###\") & \"  \""
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3300
                    Width =7071
                    Height =375
                    FontSize =8
                    TabIndex =2
                    Name ="Text105"
                    ControlSource ="=ALPHANUM([SSAN]) & \" \" & \"ريال\""
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin BoundObjectFrame
                    OldBorderStyle =0
                    Left =7937
                    Top =520
                    Width =1701
                    Height =1701
                    TabIndex =3
                    Name ="EMZA1"
                    ControlSource ="EMZA1"
                End
                Begin BoundObjectFrame
                    OldBorderStyle =0
                    Left =5280
                    Top =517
                    Width =1701
                    Height =1701
                    TabIndex =4
                    Name ="EMZA2"
                    ControlSource ="EMZA2"
                End
                Begin BoundObjectFrame
                    OldBorderStyle =0
                    Left =2655
                    Top =517
                    Width =1701
                    Height =1701
                    TabIndex =5
                    Name ="EMZA3"
                    ControlSource ="EMZA3"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =8772
                    Top =843
                    Width =1725
                    Height =285
                    FontSize =8
                    FontWeight =400
                    BackColor =128
                    ForeColor =0
                    Name ="semat1"
                    Caption ="."
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =6180
                    Top =850
                    Width =1725
                    Height =285
                    FontSize =8
                    FontWeight =400
                    BackColor =128
                    ForeColor =0
                    Name ="semat2"
                    Caption ="."
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =3525
                    Top =851
                    Width =1725
                    Height =285
                    FontSize =8
                    FontWeight =400
                    BackColor =128
                    ForeColor =0
                    Name ="semat3"
                    Caption ="."
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =8769
                    Top =510
                    Width =1725
                    Height =285
                    FontSize =8
                    FontWeight =400
                    BackColor =128
                    ForeColor =0
                    Name ="nemz1"
                    Caption ="."
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =6182
                    Top =517
                    Width =1725
                    Height =285
                    FontSize =8
                    FontWeight =400
                    BackColor =128
                    ForeColor =0
                    Name ="nemz2"
                    Caption ="."
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =3510
                    Top =518
                    Width =1725
                    Height =285
                    FontSize =8
                    FontWeight =400
                    BackColor =128
                    ForeColor =0
                    Name ="nemz3"
                    Caption ="."
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Top =631
                    TabIndex =6
                    Name ="SGN1"
                    ControlSource ="SGN1"
                End
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Top =950
                    TabIndex =7
                    Name ="SGN2"
                    ControlSource ="SGN2"
                End
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Top =1269
                    TabIndex =8
                    Name ="SGN3"
                    ControlSource ="SGN3"
                End
            End
        End
    End
End
CodeBehindForm
' See "R_SANAD_PRINT_B.cls"
