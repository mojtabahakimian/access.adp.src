Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    AllowUpdating =4
    TabularCharSet =178
    TabularFamily =48
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =10395
    DatasheetFontHeight =10
    ItemSuffix =50
    Left =270
    Top =210
    DatasheetGridlinesColor =12632256
    MaxRecords =100000
    RecSrcDt = Begin
        0xac0f2577919fe440
    End
    RecordSource ="SELECT PAY_GETP.N_SERI, PAY_GETP.BANK, PAY_GETP.DATE_S, PAY_GETP.DATE, PAY_GETP."
        "SHOBEH, PAY_GETP.MABL, PAY_GETP.N_HESAB, PAY_GETP.N_S, TCOD_BANKS.NAMES, PAY_GET"
        "P.NAME_TAH, PAY_GETP.RADIF, PAY_GETP.N_KOL, PAY_GETP.N_MOIN, PAY_GETP.N_TAF, PAY"
        "_GETP.N_KOL2, PAY_GETP.N_MOIN2, PAY_GETP.N_TAF2, PAY_GETP.N_KOL3, PAY_GETP.N_MOI"
        "N3, PAY_GETP.N_TAF3 FROM TCOD_BANKS INNER JOIN PAY_GETP ON TCOD_BANKS.CODE = PAY"
        "_GETP.BANK WHERE (PAY_GETP.N_KOL3 IS NULL) AND (PAY_GETP.N_KOL2 IS NULL) AND (PA"
        "Y_GETP.N_KOL <> 911)"
    Caption ="ليست چكهاي پرداختي"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Arial"
    OnError ="[Event Procedure]"
    RecordSourceQualifier ="dbo"
    Begin
        Begin Label
            FontItalic = NotDefault
            BackStyle =0
            TextAlign =1
            TextFontFamily =18
            FontSize =11
            FontWeight =700
            ForeColor =8388608
            FontName ="Times New Roman"
        End
        Begin Rectangle
            BackStyle =0
            BorderWidth =1
            Width =850
            Height =850
            BorderColor =8388608
        End
        Begin Line
            Width =1701
            BorderColor =8388608
        End
        Begin Image
            OldBorderStyle =0
            PictureAlignment =2
            Width =1701
            Height =1701
        End
        Begin CheckBox
            LabelX =230
            LabelY =-30
        End
        Begin BoundObjectFrame
            Width =4536
            Height =2835
            LabelX =-1701
        End
        Begin TextBox
            FELineBreak = NotDefault
            OldBorderStyle =0
            BackStyle =0
            Width =1701
            LabelX =-1701
            FontName ="Arial"
            AsianLineBreak =255
        End
        Begin ListBox
            OldBorderStyle =0
            Width =1701
            Height =1417
            LabelX =-1701
            FontName ="Arial"
        End
        Begin ComboBox
            OldBorderStyle =0
            BackStyle =0
            Width =1701
            LabelX =-1701
            FontName ="Arial"
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
        Begin BreakLevel
            ControlSource ="DATE_S"
        End
        Begin BreakLevel
            ControlSource ="N_SERI"
        End
        Begin BreakLevel
            ControlSource ="BANK"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =0
            Name ="ReportHeader"
        End
        Begin PageHeader
            Height =982
            Name ="PageHeaderSection"
            Begin
                Begin Label
                    FontItalic = NotDefault
                    TextFontCharSet =178
                    TextAlign =2
                    TextFontFamily =2
                    Left =4822
                    Top =639
                    Width =840
                    Height =285
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="DATE_S_Label"
                    Caption ="تاريخ سررسيد"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    FontItalic = NotDefault
                    TextFontCharSet =178
                    TextAlign =2
                    TextFontFamily =2
                    Left =8874
                    Top =639
                    Width =780
                    Height =285
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="N_SERI_Label"
                    Caption ="شماره سريال"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    FontItalic = NotDefault
                    TextFontCharSet =178
                    TextAlign =2
                    TextFontFamily =2
                    Left =6322
                    Top =639
                    Width =555
                    Height =285
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="SHOBEH_Label"
                    Caption ="نام شعبه"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    FontItalic = NotDefault
                    TextFontCharSet =178
                    TextAlign =2
                    TextFontFamily =2
                    Left =618
                    Top =639
                    Width =315
                    Height =285
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="MABL_Label"
                    Caption ="مبلغ"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    FontItalic = NotDefault
                    TextFontCharSet =178
                    TextAlign =2
                    TextFontFamily =2
                    Left =3790
                    Top =639
                    Width =765
                    Height =285
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="N_HESAB_Label"
                    Caption ="شماره حساب"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    FontItalic = NotDefault
                    TextFontCharSet =178
                    TextAlign =2
                    TextFontFamily =2
                    Left =7799
                    Top =639
                    Width =510
                    Height =285
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="NAMES_Label"
                    Caption ="نام بانك"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    FontItalic = NotDefault
                    TextFontCharSet =178
                    TextAlign =2
                    TextFontFamily =2
                    Left =9900
                    Top =645
                    Width =375
                    Height =285
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="RADIF_Label"
                    Caption ="رديف"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    FontItalic = NotDefault
                    BackStyle =1
                    TextFontCharSet =178
                    TextAlign =2
                    TextFontFamily =2
                    Left =3675
                    Width =3915
                    Height =540
                    FontSize =14
                    FontWeight =400
                    ForeColor =0
                    Name ="Label18"
                    Caption ="ليست چكهاي پرداختي سر رسيد شده"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    FontItalic = NotDefault
                    TextFontCharSet =178
                    TextAlign =2
                    TextFontFamily =2
                    Left =2070
                    Top =639
                    Width =1005
                    Height =285
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="Label23"
                    Caption ="نام دريافت كننده"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Rectangle
                    Top =585
                    Width =10395
                    Height =375
                    BorderColor =0
                    Name ="Box26"
                End
                Begin Line
                    BorderWidth =1
                    Left =1530
                    Top =585
                    Width =0
                    Height =397
                    BorderColor =0
                    Name ="Line37"
                End
                Begin Line
                    BorderWidth =1
                    Left =3628
                    Top =585
                    Width =0
                    Height =397
                    BorderColor =0
                    Name ="Line38"
                End
                Begin Line
                    BorderWidth =1
                    Left =4725
                    Top =585
                    Width =0
                    Height =397
                    BorderColor =0
                    Name ="Line39"
                End
                Begin Line
                    BorderWidth =1
                    Left =5822
                    Top =585
                    Width =0
                    Height =397
                    BorderColor =0
                    Name ="Line40"
                End
                Begin Line
                    BorderWidth =1
                    Left =7425
                    Top =585
                    Width =0
                    Height =397
                    BorderColor =0
                    Name ="Line41"
                End
                Begin Line
                    BorderWidth =1
                    Left =8730
                    Top =585
                    Width =0
                    Height =397
                    BorderColor =0
                    Name ="Line42"
                End
                Begin Line
                    BorderWidth =1
                    Left =9840
                    Top =585
                    Width =0
                    Height =397
                    BorderColor =0
                    Name ="Line43"
                End
                Begin TextBox
                    FELineBreak = NotDefault
                    TextFontCharSet =178
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Top =201
                    Width =1021
                    Height =360
                    FontSize =10
                    Name ="Text52"
                    ControlSource ="=[Forms]![F_MENU_CHEK]![DT2]"
                    FontName ="WeblogmaYekan"
                    InputMask ="####/##/##"
                    AsianLineBreak =0
                End
                Begin TextBox
                    FELineBreak = NotDefault
                    TextFontCharSet =178
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1752
                    Top =199
                    Width =1086
                    Height =360
                    FontSize =10
                    TabIndex =1
                    Name ="Text53"
                    ControlSource ="=[Forms]![F_MENU_CHEK]![DT1]"
                    FontName ="WeblogmaYekan"
                    InputMask ="####/##/##"
                    AsianLineBreak =0
                End
                Begin Label
                    FontItalic = NotDefault
                    TextFontCharSet =178
                    TextAlign =2
                    TextFontFamily =2
                    Left =1069
                    Top =195
                    Width =660
                    Height =360
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="Label54"
                    Caption ="تا تاريخ:"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    FontItalic = NotDefault
                    TextFontCharSet =178
                    TextAlign =2
                    TextFontFamily =2
                    Left =2872
                    Top =197
                    Width =690
                    Height =360
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="Label55"
                    Caption ="از تاريخ:"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =397
            Name ="Detail"
            Begin
                Begin TextBox
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =4729
                    Width =1056
                    Height =330
                    Name ="DATE_S"
                    ControlSource ="DATE_S"
                    FontName ="WeblogmaYekan"
                    InputMask ="####/##/##"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =8759
                    Width =1056
                    Height =330
                    TabIndex =1
                    Name ="N_SERI"
                    ControlSource ="N_SERI"
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =5809
                    Width =1611
                    Height =315
                    TabIndex =2
                    Name ="SHOBEH"
                    ControlSource ="SHOBEH"
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =81
                    Width =1431
                    Height =330
                    TabIndex =3
                    Name ="MABL"
                    ControlSource ="MABL"
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =3649
                    Width =1056
                    Height =330
                    TabIndex =4
                    Name ="N_HESAB"
                    ControlSource ="N_HESAB"
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =7411
                    Width =1311
                    Height =315
                    TabIndex =5
                    Name ="NAMES"
                    ControlSource ="NAMES"
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =0
                    RunningSum =2
                    TextFontCharSet =178
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =9866
                    Width =491
                    Height =315
                    TabIndex =6
                    Name ="RADIF"
                    ControlSource ="=1"
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =12
                    TextFontCharSet =178
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1530
                    Width =2091
                    Height =315
                    FontSize =7
                    TabIndex =7
                    Name ="NAME_TAH"
                    ControlSource ="NAME_TAH"
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    BorderWidth =1
                    Top =397
                    Width =10395
                    BorderColor =0
                    Name ="Line27"
                End
                Begin Line
                    BorderWidth =1
                    Left =1530
                    Width =0
                    Height =397
                    BorderColor =0
                    Name ="Line28"
                End
                Begin Line
                    BorderWidth =1
                    Left =3628
                    Width =0
                    Height =397
                    BorderColor =0
                    Name ="Line29"
                End
                Begin Line
                    BorderWidth =1
                    Left =4725
                    Width =0
                    Height =397
                    BorderColor =0
                    Name ="Line30"
                End
                Begin Line
                    BorderWidth =1
                    Left =5822
                    Width =0
                    Height =397
                    BorderColor =0
                    Name ="Line31"
                End
                Begin Line
                    BorderWidth =1
                    Left =7425
                    Width =0
                    Height =397
                    BorderColor =0
                    Name ="Line32"
                End
                Begin Line
                    BorderWidth =1
                    Left =8730
                    Width =0
                    Height =397
                    BorderColor =0
                    Name ="Line33"
                End
                Begin Line
                    BorderWidth =1
                    Left =9840
                    Width =0
                    Height =397
                    BorderColor =0
                    Name ="Line34"
                End
                Begin Line
                    BorderWidth =1
                    Left =10395
                    Width =0
                    Height =397
                    BorderColor =0
                    Name ="Line35"
                End
                Begin Line
                    BorderWidth =1
                    Width =0
                    Height =397
                    BorderColor =0
                    Name ="Line36"
                End
            End
        End
        Begin PageFooter
            Height =0
            Name ="PageFooterSection"
        End
        Begin FormFooter
            KeepTogether = NotDefault
            Height =428
            Name ="ReportFooter"
            Begin
                Begin TextBox
                    DecimalPlaces =0
                    RunningSum =2
                    TextFontCharSet =178
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =105
                    Top =60
                    Width =1401
                    Height =270
                    Name ="Text25"
                    ControlSource ="=Sum([MABL])"
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                End
                Begin Rectangle
                    Top =45
                    Width =10387
                    Height =375
                    BorderColor =0
                    Name ="Box47"
                End
                Begin Line
                    Left =1530
                    Top =31
                    Width =0
                    Height =397
                    BorderColor =0
                    Name ="Line48"
                End
                Begin Label
                    FontItalic = NotDefault
                    TextFontCharSet =178
                    TextAlign =2
                    TextFontFamily =2
                    Left =1927
                    Top =56
                    Width =555
                    Height =330
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="Label49"
                    Caption ="جمع كل"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
            End
        End
    End
End
CodeBehindForm
' See "CHEK_PLIST.cls"
