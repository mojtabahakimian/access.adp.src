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
    Width =10267
    DatasheetFontHeight =10
    ItemSuffix =61
    Top =210
    DatasheetGridlinesColor =12632256
    MaxRecords =100000
    RecSrcDt = Begin
        0x48cb1f3721dae340
    End
    RecordSource ="SELECT PAY_GETD.N_SERI, PAY_GETD.BANK, PAY_GETD.DATE_S, PAY_GETD.DATE, PAY_GETD."
        "SHOBEH, PAY_GETD.MABL, PAY_GETD.NAME_TAH, PAY_GETD.N_HESAB, PAY_GETD.N_S, TCOD_B"
        "ANKS.NAMES, PAY_GETD.RADIF, PAY_GETD.N_KOL, PAY_GETD.N_MOIN, PAY_GETD.N_KOL2, PA"
        "Y_GETD.N_MOIN2, PAY_GETD.N_KOL3, PAY_GETD.N_MOIN3, PAY_GETD.N_TAF, PAY_GETD.N_TA"
        "F2, PAY_GETD.N_TAF3, TDETA_HES.NAME, CHRE_LSPH.RADIF AS list_no FROM TCOD_BANKS "
        "INNER JOIN PAY_GETD ON TCOD_BANKS.CODE = PAY_GETD.BANK INNER JOIN CHRE_LSPH ON P"
        "AY_GETD.N_SERI = CHRE_LSPH.N_SERI AND PAY_GETD.BANK = CHRE_LSPH.BANK AND PAY_GET"
        "D.DATE_S = CHRE_LSPH.DATE_S LEFT OUTER JOIN TDETA_HES ON PAY_GETD.N_KOL = TDETA_"
        "HES.N_KOL AND PAY_GETD.N_MOIN = TDETA_HES.NUMBER AND PAY_GETD.N_TAF = TDETA_HES."
        "TNUMBER"
    Caption ="ليست چكهاي دريافتي"
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
            Height =1154
            Name ="PageHeaderSection"
            Begin
                Begin Label
                    FontItalic = NotDefault
                    TextFontCharSet =178
                    TextAlign =2
                    TextFontFamily =2
                    Left =3757
                    Top =811
                    Width =870
                    Height =330
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
                    Left =8310
                    Top =817
                    Width =1290
                    Height =330
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
                    Left =5272
                    Top =811
                    Width =1005
                    Height =330
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
                    Left =678
                    Top =811
                    Width =735
                    Height =330
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
                    Left =1967
                    Top =817
                    Width =1680
                    Height =330
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
                    Left =7102
                    Top =811
                    Width =810
                    Height =330
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
                    Left =9735
                    Top =817
                    Width =405
                    Height =330
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="RADIF_Label"
                    Caption ="رديف"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Rectangle
                    Top =757
                    Width =10246
                    Height =375
                    BorderColor =0
                    Name ="Box26"
                End
                Begin Line
                    BorderWidth =1
                    Left =1950
                    Top =757
                    Width =0
                    Height =397
                    BorderColor =0
                    Name ="Line38"
                End
                Begin Line
                    BorderWidth =1
                    Left =3675
                    Top =757
                    Width =0
                    Height =397
                    BorderColor =0
                    Name ="Line39"
                End
                Begin Line
                    BorderWidth =1
                    Left =4772
                    Top =757
                    Width =0
                    Height =397
                    BorderColor =0
                    Name ="Line40"
                End
                Begin Line
                    BorderWidth =1
                    Left =6885
                    Top =757
                    Width =0
                    Height =397
                    BorderColor =0
                    Name ="Line41"
                End
                Begin Line
                    BorderWidth =1
                    Left =8295
                    Top =757
                    Width =0
                    Height =397
                    BorderColor =0
                    Name ="Line42"
                End
                Begin Line
                    BorderWidth =1
                    Left =9690
                    Top =757
                    Width =0
                    Height =397
                    BorderColor =0
                    Name ="Line43"
                End
                Begin Label
                    FontItalic = NotDefault
                    TextFontCharSet =178
                    TextFontFamily =2
                    Left =8745
                    Top =337
                    Width =1470
                    Height =330
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="Label50"
                    Caption ="ليست چكهاي نزد حساب:"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =4815
                    Top =337
                    Width =3861
                    Height =315
                    Name ="NAME"
                    ControlSource ="NAME"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =120
                    Top =285
                    Width =906
                    Height =315
                    TabIndex =1
                    Name ="Text56"
                    ControlSource ="=[Forms]![vbp_check]![dts]"
                    FontName ="WeblogmaYekan"
                    InputMask ="####/##/##"
                End
                Begin Label
                    FontItalic = NotDefault
                    TextFontCharSet =178
                    TextFontFamily =2
                    Left =1065
                    Top =300
                    Width =1470
                    Height =330
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="Label57"
                    Caption ="تاريخ :"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
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
                    Left =2546
                    Width =5391
                    Height =390
                    FontSize =10
                    TabIndex =2
                    Name ="Text58"
                    ControlSource ="=[Forms]![BASEKNOW]![NAME]"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =120
                    Width =906
                    Height =315
                    TabIndex =3
                    Name ="Text59"
                    ControlSource ="=[Forms]![vbp_check]![list_no]"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    FontItalic = NotDefault
                    TextFontCharSet =178
                    TextFontFamily =2
                    Left =1065
                    Top =15
                    Width =1470
                    Height =330
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="Label60"
                    Caption ="شماره ليست :"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =284
            Name ="Detail"
            Begin
                Begin TextBox
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =3679
                    Width =1056
                    Height =284
                    Name ="DATE_S"
                    ControlSource ="DATE_S"
                    FontName ="WeblogmaYekan"
                    InputMask ="####/##/##"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =8309
                    Width =1341
                    Height =284
                    TabIndex =1
                    Name ="N_SERI"
                    ControlSource ="N_SERI"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =4759
                    Width =1986
                    Height =284
                    TabIndex =2
                    Name ="SHOBEH"
                    ControlSource ="SHOBEH"
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =45
                    Width =1851
                    Height =284
                    TabIndex =3
                    Name ="MABL"
                    ControlSource ="MABL"
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1971
                    Width =1671
                    Height =284
                    TabIndex =4
                    Name ="N_HESAB"
                    ControlSource ="N_HESAB"
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =6901
                    Width =1326
                    Height =284
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
                    Left =9716
                    Width =491
                    Height =284
                    TabIndex =6
                    Name ="RADIF"
                    ControlSource ="=1"
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    BorderWidth =1
                    Top =284
                    Width =10260
                    BorderColor =0
                    Name ="Line27"
                End
                Begin Line
                    BorderWidth =1
                    Left =1950
                    Width =0
                    Height =284
                    BorderColor =0
                    Name ="Line29"
                End
                Begin Line
                    BorderWidth =1
                    Left =3675
                    Width =0
                    Height =284
                    BorderColor =0
                    Name ="Line30"
                End
                Begin Line
                    BorderWidth =1
                    Left =4772
                    Width =0
                    Height =284
                    BorderColor =0
                    Name ="Line31"
                End
                Begin Line
                    BorderWidth =1
                    Left =6885
                    Width =0
                    Height =284
                    BorderColor =0
                    Name ="Line32"
                End
                Begin Line
                    BorderWidth =1
                    Left =8295
                    Width =0
                    Height =284
                    BorderColor =0
                    Name ="Line33"
                End
                Begin Line
                    BorderWidth =1
                    Left =9690
                    Width =0
                    Height =284
                    BorderColor =0
                    Name ="Line34"
                End
                Begin Line
                    BorderWidth =1
                    Left =10245
                    Width =0
                    Height =284
                    BorderColor =0
                    Name ="Line35"
                End
                Begin Line
                    BorderWidth =1
                    Width =0
                    Height =284
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
            Height =451
            Name ="ReportFooter"
            Begin
                Begin TextBox
                    DecimalPlaces =0
                    RunningSum =2
                    TextFontCharSet =178
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =60
                    Width =1806
                    Height =330
                    Name ="Text25"
                    ControlSource ="=Sum([MABL])"
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                End
                Begin Rectangle
                    Width =10267
                    Height =375
                    BorderColor =0
                    Name ="Box47"
                End
                Begin Line
                    BorderWidth =1
                    Left =1950
                    Width =0
                    Height =391
                    BorderColor =0
                    Name ="Line48"
                End
                Begin Label
                    FontItalic = NotDefault
                    TextFontCharSet =178
                    TextAlign =2
                    TextFontFamily =2
                    Left =1972
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
' See "VBP_CHEK_DLISTS.cls"
