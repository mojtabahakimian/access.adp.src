Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    DefaultView =0
    AllowUpdating =4
    TabularCharSet =178
    TabularFamily =42
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =14663
    DatasheetFontHeight =10
    ItemSuffix =115
    Left =1560
    Top =2250
    DatasheetGridlinesColor =12632256
    Tag ="@Forms___F_MENU_KART___FT BIGINT =[Forms]![F_MENU_KART]![FT],@Forms___F_MENU_KAR"
        "T___ANBAR NVARCHAR(10) = [Forms]![F_MENU_KART]![ANBAR],@Forms___F_MENU_KART___DT"
        "2 bigint= [Forms]![F_MENU_KART]![DT2]"
    RecSrcDt = Begin
        0xf98bed054fefe540
    End
    RecordSource ="SELECT PM_MVAD.JOB_IDD, PM_MVAD.MV_ANBAR, PM_MVAD.MV_CODE, PM_MVAD.MV_MEGH, PM_M"
        "VAD.MV_MEGHk, PM_MVAD.MV_VAHED_K, PM_MVAD.MV_DESCRIPTION, PM_JOB.AS_IDD, PM_JOB."
        "JO_STDATE, PM_JOB.JO_JOB, PM_JOB.JO_PERSON FROM PM_JOB INNER JOIN PM_MVAD ON PM_"
        "JOB.IDD = PM_MVAD.JOB_IDD WHERE (PM_MVAD.JOB_IDD = 1)"
    Caption ="مواد لازم براي دستور کار"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Arial (Arabic)"
    OnError ="[Event Procedure]"
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
        Begin BreakLevel
            GroupFooter = NotDefault
            ControlSource ="JO_PERSON"
        End
        Begin BreakLevel
            ControlSource ="JO_STDATE"
        End
        Begin PageHeader
            Height =1214
            Name ="PageHeader"
            Begin
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =14114
                    Top =747
                    Width =510
                    Height =435
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="N_S Label"
                    Caption ="رديف"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="N_S_Label"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =5610
                    Top =753
                    Width =921
                    Height =435
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="BED Label"
                    Caption ="مقدار"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="BED_Label"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =4474
                    Top =750
                    Width =1116
                    Height =435
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="BES Label"
                    Caption =" واحد"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="BES_Label"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =3225
                    Top =750
                    Width =1177
                    Height =435
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="Expr1 Label"
                    Caption ="مقدار كل"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="Expr1_Label"
                End
                Begin Line
                    BorderWidth =2
                    Left =60
                    Top =735
                    Width =14575
                    Name ="Line26"
                End
                Begin Line
                    BorderWidth =2
                    Left =60
                    Top =1200
                    Width =14575
                    Name ="Line27"
                End
                Begin Line
                    Left =14115
                    Top =735
                    Width =0
                    Height =449
                    Name ="Line28"
                End
                Begin Line
                    Left =6540
                    Top =756
                    Width =0
                    Height =449
                    Name ="Line32"
                End
                Begin Line
                    Left =5599
                    Top =765
                    Width =0
                    Height =449
                    Name ="Line33"
                End
                Begin Line
                    Left =4424
                    Top =750
                    Width =0
                    Height =449
                    Name ="Line34"
                End
                Begin Line
                    Left =3180
                    Top =753
                    Width =0
                    Height =449
                    Name ="Line35"
                End
                Begin Line
                    BorderWidth =2
                    Left =60
                    Top =753
                    Width =0
                    Height =449
                    Name ="Line36"
                End
                Begin Line
                    BorderWidth =2
                    Left =14626
                    Top =753
                    Width =0
                    Height =449
                    Name ="Line37"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =11520
                    Top =753
                    Width =2541
                    Height =435
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="Label74"
                    Caption ="کار"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =10470
                    Top =753
                    Width =1041
                    Height =435
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="Label85"
                    Caption ="تاريخ"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    Left =11520
                    Top =756
                    Width =0
                    Height =449
                    Name ="Line86"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =6566
                    Top =753
                    Width =3861
                    Height =435
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="Label90"
                    Caption ="مواد/کالا"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    Left =10496
                    Top =756
                    Width =0
                    Height =449
                    Name ="Line91"
                End
                Begin Label
                    TextAlign =0
                    TextFontFamily =2
                    Left =10714
                    Top =113
                    Width =3855
                    Height =540
                    FontSize =14
                    FontWeight =400
                    ForeColor =0
                    Name ="Label110"
                    Caption ="مواد مورد نياز دستور کارWork Order"
                    FontName ="WeblogmaYekan"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    TextAlign =3
                    TextFontFamily =2
                    DecimalPlaces =0
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =56
                    Top =226
                    Width =3330
                    FontSize =8
                    ForeColor =8388608
                    Name ="JO_PERSON"
                    ControlSource ="JO_PERSON"
                    RowSourceType ="Value List"
                    ColumnWidths ="0"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextFontFamily =2
                            Left =3465
                            Top =228
                            Width =630
                            Height =285
                            FontSize =8
                            FontWeight =400
                            ForeColor =8388608
                            Name ="JO_PERSON_Label"
                            Caption ="مجري:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =60
                    Top =753
                    Width =3096
                    Height =435
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="Label113"
                    Caption ="دستگاه"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =419
            Name ="Detail"
            Begin
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5625
                    Width =921
                    Height =375
                    FontSize =8
                    Name ="mv_megh"
                    ControlSource ="mv_megh"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3240
                    Width =1176
                    Height =375
                    FontSize =8
                    TabIndex =2
                    Name ="MV_MEGHk"
                    ControlSource ="MV_MEGHk"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    RunningSum =1
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =14180
                    Width =471
                    Height =375
                    FontSize =8
                    TabIndex =3
                    Name ="N_S"
                    ControlSource ="=1"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    Left =14115
                    Width =0
                    Height =419
                    Name ="Line38"
                End
                Begin Line
                    Left =6555
                    Width =0
                    Height =419
                    Name ="Line42"
                End
                Begin Line
                    Left =5614
                    Width =0
                    Height =419
                    Name ="Line43"
                End
                Begin Line
                    Left =4439
                    Width =0
                    Height =419
                    Name ="Line44"
                End
                Begin Line
                    Left =3195
                    Width =0
                    Height =419
                    Name ="Line45"
                End
                Begin Line
                    BorderWidth =2
                    Left =75
                    Width =0
                    Height =419
                    Name ="Line46"
                End
                Begin Line
                    BorderWidth =2
                    Left =14641
                    Width =0
                    Height =419
                    Name ="Line47"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =75
                    Top =419
                    Width =14588
                    BorderColor =4210752
                    Name ="Line48"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =10485
                    Width =1041
                    Height =375
                    FontSize =8
                    TabIndex =4
                    Name ="JO_STDATE"
                    ControlSource ="JO_STDATE"
                    FontName ="WeblogmaYekan"
                    InputMask ="####/##/##"
                End
                Begin Line
                    Left =11535
                    Width =0
                    Height =419
                    Name ="Line88"
                End
                Begin Line
                    Left =10511
                    Width =0
                    Height =419
                    Name ="Line93"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    TextAlign =3
                    TextFontFamily =2
                    DecimalPlaces =0
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =11565
                    Width =2547
                    Height =390
                    FontSize =8
                    TabIndex =6
                    Name ="JO_JOB"
                    ControlSource ="JO_JOB"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT IDD, JOB_TITLE FROM PM_JOBTITLE"
                    ColumnWidths ="0"
                    FontName ="WeblogmaYekan"
                End
                Begin ComboBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    DecimalPlaces =0
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =6581
                    Width =3861
                    Height =375
                    FontSize =8
                    TabIndex =5
                    Name ="MV_CODE"
                    ControlSource ="MV_CODE"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT CODE, NAME FROM STUF_DEF"
                    ColumnWidths ="0"
                    FontName ="WeblogmaYekan"
                    Format ="Standard"
                End
                Begin ComboBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    DecimalPlaces =0
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =4470
                    Width =1117
                    Height =375
                    FontSize =8
                    TabIndex =1
                    Name ="MV_vahed_k"
                    ControlSource ="MV_vahed_k"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT CODE, NAMES FROM TCOD_VAHEDS"
                    ColumnWidths ="0"
                    FontName ="WeblogmaYekan"
                    Format ="Standard"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    TextAlign =3
                    TextFontFamily =2
                    DecimalPlaces =0
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =120
                    Width =3057
                    Height =375
                    FontSize =8
                    TabIndex =7
                    Name ="AS_IDD"
                    ControlSource ="AS_IDD"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT IDD, AS_NAME FROM PM_ASSETS"
                    ColumnWidths ="0"
                    FontName ="WeblogmaYekan"
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            ForceNewPage =2
            Height =0
            Name ="GroupFooter1"
            Begin
                Begin Line
                    BorderWidth =2
                    Left =75
                    Width =14575
                    Name ="Line111"
                End
            End
        End
        Begin PageFooter
            Height =737
            Name ="PageFooter"
            Begin
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =9865
                    Width =4421
                    Height =300
                    FontSize =10
                    Name ="Text23"
                    ControlSource ="=FARSIDATE(Now())"
                    FontName ="WeblogmaYekan"
                    InputMask ="####/##/##"
                End
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =56
                    Width =4421
                    Height =300
                    FontSize =10
                    TabIndex =1
                    Name ="Text24"
                    ControlSource ="=\"صفحه \" & [Page] & \" از \" & [Pages]"
                    FontName ="WeblogmaYekan"
                End
            End
        End
    End
End
CodeBehindForm
' See "PM_JOB_MAVAD_REPORT.cls"
