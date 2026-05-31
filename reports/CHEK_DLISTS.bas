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
    Width =11167
    DatasheetFontHeight =10
    ItemSuffix =56
    Top =210
    DatasheetGridlinesColor =12632256
    MaxRecords =100000
    RecSrcDt = Begin
        0xf19cde9319cce240
    End
    RecordSource ="SELECT PAY_GETD.N_SERI, PAY_GETD.BANK, PAY_GETD.DATE_S, PAY_GETD.DATE, PAY_GETD."
        "SHOBEH, PAY_GETD.MABL, PAY_GETD.NAME_TAH, PAY_GETD.N_HESAB, PAY_GETD.N_S, TCOD_B"
        "ANKS.NAMES, PAY_GETD.RADIF, PAY_GETD.N_KOL, PAY_GETD.N_MOIN, PAY_GETD.N_KOL2, PA"
        "Y_GETD.N_MOIN2, PAY_GETD.N_KOL3, PAY_GETD.N_MOIN3, PAY_GETD.N_TAF, PAY_GETD.N_TA"
        "F2, PAY_GETD.N_TAF3, TDETA_HES.NAME FROM TCOD_BANKS INNER JOIN PAY_GETD ON TCOD_"
        "BANKS.CODE = PAY_GETD.BANK LEFT OUTER JOIN TDETA_HES ON PAY_GETD.N_KOL = TDETA_H"
        "ES.N_KOL AND PAY_GETD.N_MOIN = TDETA_HES.NUMBER AND PAY_GETD.N_TAF = TDETA_HES.T"
        "NUMBER WHERE (PAY_GETD.N_KOL = 122 OR PAY_GETD.N_KOL IS NULL) AND (PAY_GETD.N_KO"
        "L2 IS NULL) AND (PAY_GETD.N_KOL3 IS NULL)"
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
            Height =982
            Name ="PageHeaderSection"
            Begin
                Begin Label
                    FontItalic = NotDefault
                    TextFontCharSet =178
                    TextAlign =2
                    TextFontFamily =2
                    Left =5902
                    Top =639
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
                    Left =9705
                    Top =645
                    Width =810
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
                    Left =7417
                    Top =639
                    Width =555
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
                    Left =603
                    Top =639
                    Width =345
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
                    Left =4995
                    Top =645
                    Width =810
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
                    Left =8737
                    Top =639
                    Width =555
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
                    Left =10635
                    Top =645
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
                Begin Label
                    FontItalic = NotDefault
                    BackStyle =1
                    TextFontCharSet =178
                    TextAlign =2
                    TextFontFamily =2
                    Left =3685
                    Width =3885
                    Height =540
                    FontSize =14
                    FontWeight =400
                    ForeColor =0
                    Name ="Label18"
                    Caption ="ليست چكهاي دريافتي سر رسيد شده"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    FontItalic = NotDefault
                    TextFontCharSet =178
                    TextAlign =2
                    TextFontFamily =2
                    Left =3630
                    Top =630
                    Width =1095
                    Height =330
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="Label23"
                    Caption ="نام پرداخت كننده"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Rectangle
                    Top =585
                    Width =11146
                    Height =375
                    BorderColor =0
                    Name ="Box26"
                End
                Begin Line
                    BorderWidth =1
                    Left =1455
                    Top =585
                    Width =0
                    Height =397
                    BorderColor =0
                    Name ="Line37"
                End
                Begin Line
                    BorderWidth =1
                    Left =4978
                    Top =585
                    Width =0
                    Height =397
                    BorderColor =0
                    Name ="Line38"
                End
                Begin Line
                    BorderWidth =1
                    Left =5820
                    Top =585
                    Width =0
                    Height =397
                    BorderColor =0
                    Name ="Line39"
                End
                Begin Line
                    BorderWidth =1
                    Left =6917
                    Top =585
                    Width =0
                    Height =397
                    BorderColor =0
                    Name ="Line40"
                End
                Begin Line
                    BorderWidth =1
                    Left =8520
                    Top =585
                    Width =0
                    Height =397
                    BorderColor =0
                    Name ="Line41"
                End
                Begin Line
                    BorderWidth =1
                    Left =9690
                    Top =585
                    Width =0
                    Height =397
                    BorderColor =0
                    Name ="Line42"
                End
                Begin Line
                    BorderWidth =1
                    Left =10590
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
                Begin Label
                    FontItalic = NotDefault
                    TextFontCharSet =178
                    TextAlign =2
                    TextFontFamily =2
                    Left =1845
                    Top =630
                    Width =1215
                    Height =330
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="Label50"
                    Caption ="نزد حساب"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =1
                    Left =3298
                    Top =585
                    Width =0
                    Height =397
                    BorderColor =0
                    Name ="Line51"
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
                    Left =5824
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
                    Left =9704
                    Width =861
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
                    Left =6904
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
                    Left =4999
                    Width =801
                    Height =330
                    FontSize =5
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
                    Left =8536
                    Width =1146
                    Height =315
                    TabIndex =5
                    Name ="NAMES"
                    ControlSource ="NAMES"
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
                    Left =10616
                    Width =491
                    Height =315
                    TabIndex =6
                    Name ="RADIF"
                    ControlSource ="RADIF"
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
                    Left =3315
                    Width =1656
                    Height =315
                    FontSize =7
                    TabIndex =7
                    Name ="NAME_TAH"
                    ControlSource ="NAME_TAH"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    BorderWidth =1
                    Top =397
                    Width =11160
                    BorderColor =0
                    Name ="Line27"
                End
                Begin Line
                    BorderWidth =1
                    Left =1449
                    Width =0
                    Height =397
                    BorderColor =0
                    Name ="Line28"
                End
                Begin Line
                    BorderWidth =1
                    Left =4978
                    Width =0
                    Height =397
                    BorderColor =0
                    Name ="Line29"
                End
                Begin Line
                    BorderWidth =1
                    Left =5820
                    Width =0
                    Height =397
                    BorderColor =0
                    Name ="Line30"
                End
                Begin Line
                    BorderWidth =1
                    Left =6917
                    Width =0
                    Height =397
                    BorderColor =0
                    Name ="Line31"
                End
                Begin Line
                    BorderWidth =1
                    Left =8520
                    Width =0
                    Height =397
                    BorderColor =0
                    Name ="Line32"
                End
                Begin Line
                    BorderWidth =1
                    Left =9690
                    Width =0
                    Height =397
                    BorderColor =0
                    Name ="Line33"
                End
                Begin Line
                    BorderWidth =1
                    Left =10590
                    Width =0
                    Height =397
                    BorderColor =0
                    Name ="Line34"
                End
                Begin Line
                    BorderWidth =1
                    Left =11145
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
                Begin TextBox
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1485
                    Width =1806
                    Height =315
                    FontSize =7
                    TabIndex =8
                    Name ="NAME"
                    ControlSource ="NAME"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    BorderWidth =1
                    Left =3298
                    Width =0
                    Height =397
                    BorderColor =0
                    Name ="Line55"
                End
            End
        End
        Begin PageFooter
            Height =0
            Name ="PageFooterSection"
        End
        Begin FormFooter
            KeepTogether = NotDefault
            Height =442
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
                    Left =15
                    Top =59
                    Width =1401
                    Height =270
                    Name ="Text25"
                    ControlSource ="=Sum([MABL])"
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                End
                Begin Rectangle
                    Top =45
                    Width =11167
                    Height =375
                    BorderColor =0
                    Name ="Box47"
                End
                Begin Line
                    BorderWidth =1
                    Left =1455
                    Top =45
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
' See "CHEK_DLISTS.cls"
