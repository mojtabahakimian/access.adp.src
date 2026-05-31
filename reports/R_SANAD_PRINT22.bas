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
    Width =10830
    DatasheetFontHeight =10
    ItemSuffix =242
    Left =270
    Top =1605
    DatasheetGridlinesColor =12632256
    OnPage ="[Event Procedure]"
    RecSrcDt = Begin
        0xcb465d1fce64e340
    End
    RecordSource ="SELECT DEED_HED.N_S, DEED_HED.DATE_S, DEED_HED.SHARH_S, DEED_DTL.HES_K, TOTA_HES"
        ".NAME, SUM(DEED_DTL.BED) AS sbed, SUM(DEED_DTL.BES) AS sbes, DEED_HED.BASE, dbo."
        "UIIF(DEED_DTL.BED, N'>', 0, 1, 0) AS gr FROM DEED_DTL INNER JOIN DEED_HED ON DEE"
        "D_DTL.N_S = DEED_HED.N_S INNER JOIN TOTA_HES ON DEED_DTL.HES_K = TOTA_HES.NUMBER"
        " GROUP BY DEED_HED.N_S, DEED_HED.DATE_S, DEED_HED.SHARH_S, DEED_DTL.HES_K, TOTA_"
        "HES.NAME, DEED_HED.BASE, dbo.UIIF(DEED_DTL.BED, N'>', 0, 1, 0)"
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
            ControlSource ="gr"
        End
        Begin BreakLevel
            ControlSource ="HES_K"
        End
        Begin PageHeader
            Height =2090
            OnFormat ="[Event Procedure]"
            Name ="PageHeader"
            Begin
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =10215
                    Top =1395
                    Width =555
                    Height =345
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="N_S Label"
                    Caption =" حساب "
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="N_S_Label"
                End
                Begin Line
                    BorderWidth =2
                    Left =15
                    Top =1410
                    Width =10808
                    BorderColor =0
                    Name ="Line26"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Top =1770
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
                    Top =1380
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
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =6510
                    Top =1395
                    Width =1755
                    Height =375
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="Label90"
                    Caption ="نام حساب  "
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
                    Top =1380
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
                    Top =1770
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
                    Left =1669
                    Top =1790
                    Width =1559
                    Height =300
                    FontSize =10
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
                    Top =1790
                    Width =1574
                    Height =300
                    FontSize =10
                    TabIndex =5
                    Name ="Text226"
                    ControlSource ="=Format(Str([sas]),\"#,###\") & \"  \""
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =56
                    Top =170
                    Width =1071
                    Height =375
                    FontSize =10
                    TabIndex =6
                    Name ="BASE"
                    ControlSource ="BASE"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextAlign =0
                            TextFontFamily =2
                            Left =1140
                            Top =165
                            Width =1035
                            Height =375
                            FontSize =10
                            Name ="Label231"
                            Caption ="شماره مبنا :"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    BorderWidth =1
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =1
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =116
                    Top =963
                    Width =9651
                    Height =360
                    FontSize =10
                    TabIndex =7
                    Name ="SHARH_S"
                    ControlSource ="SHARH_S"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextFontFamily =2
                    Left =9808
                    Top =975
                    Width =975
                    Height =360
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="Label239"
                    Caption ="شرح سنـد:"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            CanGrow = NotDefault
            Height =375
            Name ="Detail"
            Begin
                Begin TextBox
                    CanShrink = NotDefault
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =1
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =10200
                    Width =587
                    Height =375
                    FontSize =10
                    Name ="HES_K"
                    ControlSource ="HES_K"
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
                    Left =3402
                    Width =6756
                    Height =375
                    FontSize =10
                    TabIndex =1
                    Name ="kolNAME"
                    ControlSource ="NAME"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    TextFontFamily =2
                    Left =105
                    Width =771
                    TabIndex =2
                    Name ="Text185"
                    ControlSource ="=[Pages]"
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
                    Left =1710
                    Width =1559
                    Height =330
                    FontSize =10
                    TabIndex =3
                    Name ="SumOfBED1"
                    ControlSource ="=Format(Str([SBED]),\"#,###\") & \"  \""
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
                    Left =105
                    Width =1559
                    Height =330
                    FontSize =10
                    TabIndex =4
                    Name ="SumOfBES1"
                    ControlSource ="=Format(Str([SBES]),\"#,###\") & \"  \""
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    RunningSum =1
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =3685
                    Top =56
                    TabIndex =5
                    Name ="san"
                    ControlSource ="sbed"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    RunningSum =1
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =3911
                    Top =113
                    TabIndex =6
                    Name ="sas"
                    ControlSource ="sbes"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =5896
                    TabIndex =7
                    Name ="sbed"
                    ControlSource ="sbed"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =6122
                    Top =57
                    TabIndex =8
                    Name ="sbes"
                    ControlSource ="sbes"
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
            Height =1631
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
                    Left =1680
                    Width =1560
                    Height =375
                    FontSize =10
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
                    FontSize =10
                    TabIndex =1
                    Name ="Text212"
                    ControlSource ="=Format(Str([sas]),\"#,###\") & \"  \""
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                End
                Begin BoundObjectFrame
                    OldBorderStyle =0
                    OLETypeAllowed =0
                    Left =60
                    Top =465
                    Width =10715
                    Height =1166
                    TabIndex =2
                    Name ="OLEBound157"
                    ControlSource ="=[Forms]![Baseknow]![EMZA]"
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
                    FontSize =10
                    TabIndex =3
                    Name ="Text105"
                    ControlSource ="=ALPHANUM([SSAN]) & \" \" & \"ريال\""
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
            End
        End
    End
End
CodeBehindForm
' See "R_SANAD_PRINT22.cls"
