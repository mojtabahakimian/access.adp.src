Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    DefaultView =0
    AllowUpdating =4
    TabularCharSet =178
    TabularFamily =42
    DateGrouping =1
    Orientation =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =64
    GridY =64
    Width =10753
    DatasheetFontHeight =10
    ItemSuffix =27
    Left =600
    Top =540
    DatasheetGridlinesColor =12632256
    Tag ="@Forms___CHK_V_PRINT___CHK_V_PRINTSUB___Form___N_SERI float = [Forms]![CHK_V_PRI"
        "NT]![CHK_V_PRINTSUB].[Form]![N_SERI] ,\015\012 @Forms___CHK_V_PRINT___CHK_V_PRIN"
        "TSUB___Form___DATE_S  = [Forms]![CHK_V_PRINT]![CHK_V_PRINTSUB].[Form]![DATE_S] ,"
        "\015\012 @Forms___CHK_V_PRINT___CHK_V_PRINTSUB___Form___BANK  = [Forms]![CHK_V_P"
        "RINT]![CHK_V_PRINTSUB].[Form]![BANK]"
    RecSrcDt = Begin
        0x53fa0807dca8e440
    End
    RecordSource ="SELECT PAY_GETD.SHOBEH, PAY_GETD.MABL, PAY_GETD.DATE_S AS Expr1, PAY_GETD.N_KOL,"
        " PAY_GETD.N_MOIN, PAY_GETD.N_TAF, PAY_GETD.N_S, CHRE_LSPH.BANK, CHRE_LSPH.DATE_S"
        ", CHRE_LSPH.N_SERI, CHRE_LSPH.DATE, CHRE_LSPH.RADIF, PAY_GETD.N_HESAB, PAY_GETD."
        "HES1 AS hes, CUST_HESAB.ECODE AS BKK, CUST_HESAB.TOZIH, CUST_HESAB.TEL AS CODE, "
        "TCOD_BANKS.NAMES + N' ' + PAY_GETD.SHOBEH AS shob, CUST_HESAB.ADDRESS FROM PAY_G"
        "ETD INNER JOIN TCOD_BANKS ON PAY_GETD.BANK = TCOD_BANKS.CODE INNER JOIN CUST_HES"
        "AB ON PAY_GETD.HES1 = CUST_HESAB.hes RIGHT OUTER JOIN CHRE_LSPH ON PAY_GETD.N_SE"
        "RI = CHRE_LSPH.N_SERI AND PAY_GETD.BANK = CHRE_LSPH.BANK AND PAY_GETD.DATE_S = C"
        "HRE_LSPH.DATE_S WHERE (NOT (PAY_GETD.MABL IS NULL))"
    Caption ="چك قابل چاپ"
    DatasheetFontName ="Arial (Arabic)"
    RecordSourceQualifier ="dbo"
    Begin
        Begin Label
            BackStyle =0
            TextFontCharSet =178
            TextFontFamily =2
            FontName ="Arial (Arabic)"
        End
        Begin Rectangle
            Width =850
            Height =850
        End
        Begin Line
            Width =1701
        End
        Begin TextBox
            OldBorderStyle =0
            TextFontCharSet =178
            TextFontFamily =2
            Width =1701
            LabelX =-1701
            FontName ="Arial (Arabic)"
        End
        Begin ComboBox
            OldBorderStyle =0
            TextFontCharSet =178
            TextFontFamily =2
            Width =1701
            LabelX =-1701
            FontName ="Arial (Arabic)"
        End
        Begin Section
            KeepTogether = NotDefault
            Height =9017
            OnFormat ="[Event Procedure]"
            Name ="Detail"
            Begin
                Begin TextBox
                    BorderWidth =1
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =1
                    KeyboardLanguage =15
                    ScrollBarAlign =1
                    BackStyle =0
                    NumeralShapes =2
                    Left =7293
                    Width =2321
                    Height =495
                    FontSize =14
                    FontWeight =700
                    BorderColor =4210752
                    Name ="DATE"
                    ControlSource ="=FARSIDATE(Date())"
                    FontName ="Zar"
                    InputMask ="####/##/##"
                End
                Begin TextBox
                    BorderWidth =1
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    BackStyle =0
                    NumeralShapes =2
                    Left =1904
                    Top =2763
                    Width =7201
                    Height =495
                    FontSize =13
                    FontWeight =700
                    TabIndex =1
                    BorderColor =4210752
                    Name ="TEXTMABL"
                    ControlSource ="=\"*\" & ALPHANUM([MABL]) & \" ريال*************\""
                    FontName ="Zar"
                End
                Begin TextBox
                    BorderWidth =1
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    BackStyle =0
                    NumeralShapes =2
                    Left =1470
                    Top =1301
                    Width =8051
                    Height =495
                    FontSize =14
                    FontWeight =700
                    TabIndex =2
                    BorderColor =4210752
                    Name ="NAMEE"
                    ControlSource ="ADDRESS"
                    FontName ="Zar"
                End
                Begin TextBox
                    BorderWidth =1
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    BackStyle =0
                    NumeralShapes =2
                    Left =633
                    Top =3465
                    Width =2268
                    Height =495
                    FontSize =14
                    FontWeight =700
                    TabIndex =3
                    BorderColor =4210752
                    Name ="MABL"
                    ControlSource ="MABL"
                    Format ="\"-/\"#,###"
                    FontName ="Zar"
                End
                Begin TextBox
                    BorderWidth =1
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =1
                    KeyboardLanguage =15
                    ScrollBarAlign =1
                    BackStyle =0
                    NumeralShapes =2
                    Left =2568
                    Top =26
                    Width =3686
                    Height =495
                    FontSize =14
                    FontWeight =700
                    TabIndex =4
                    BorderColor =4210752
                    Name ="SHO"
                    ControlSource ="TOZIH"
                    FontName ="Zar"
                End
                Begin TextBox
                    DecimalPlaces =0
                    BorderWidth =1
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =1
                    KeyboardLanguage =15
                    ScrollBarAlign =1
                    BackStyle =0
                    NumeralShapes =2
                    Left =7293
                    Top =615
                    Width =2321
                    Height =495
                    FontSize =14
                    FontWeight =700
                    TabIndex =5
                    BorderColor =4210752
                    Name ="SHES"
                    ControlSource ="CODE"
                    Format ="Fixed"
                    FontName ="Zar"
                End
                Begin TextBox
                    BorderWidth =1
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =1
                    KeyboardLanguage =15
                    ScrollBarAlign =1
                    BackStyle =0
                    NumeralShapes =2
                    Left =8225
                    Top =2028
                    Width =2321
                    Height =495
                    FontSize =14
                    FontWeight =700
                    TabIndex =6
                    BorderColor =4210752
                    Name ="CDATE"
                    ControlSource ="DATE_S"
                    FontName ="Zar"
                    InputMask ="####/##/##"
                End
                Begin TextBox
                    DecimalPlaces =0
                    BorderWidth =1
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =1
                    KeyboardLanguage =15
                    ScrollBarAlign =1
                    BackStyle =0
                    NumeralShapes =2
                    Left =5833
                    Top =2028
                    Width =2321
                    Height =495
                    FontSize =14
                    FontWeight =700
                    TabIndex =7
                    BorderColor =4210752
                    Name ="SER"
                    ControlSource ="N_SERI"
                    Format ="Fixed"
                    FontName ="Zar"
                End
                Begin TextBox
                    DecimalPlaces =0
                    BorderWidth =1
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =1
                    KeyboardLanguage =15
                    ScrollBarAlign =1
                    BackStyle =0
                    NumeralShapes =2
                    Left =3483
                    Top =2028
                    Width =2321
                    Height =495
                    FontSize =14
                    FontWeight =700
                    TabIndex =8
                    BorderColor =4210752
                    Name ="CHES"
                    ControlSource ="N_HESAB"
                    Format ="Fixed"
                    FontName ="Zar"
                End
                Begin TextBox
                    BorderWidth =1
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =1
                    KeyboardLanguage =15
                    ScrollBarAlign =1
                    BackStyle =0
                    NumeralShapes =2
                    Left =1094
                    Top =2010
                    Width =2321
                    Height =495
                    FontSize =14
                    FontWeight =700
                    TabIndex =9
                    BorderColor =4210752
                    Name ="BSHO"
                    ControlSource ="shob"
                    FontName ="Zar"
                End
                Begin TextBox
                    BorderWidth =1
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =1
                    KeyboardLanguage =15
                    ScrollBarAlign =1
                    BackStyle =0
                    NumeralShapes =2
                    Top =2639
                    Width =2321
                    Height =495
                    FontSize =14
                    FontWeight =700
                    TabIndex =10
                    BorderColor =4210752
                    Name ="CITY"
                    ControlSource ="=[Forms]![CHKREC_HES]![CITY]"
                    FontName ="Zar"
                End
                Begin TextBox
                    BackStyle =0
                    IMESentenceMode =3
                    Left =673
                    Top =850
                    Width =126
                    TabIndex =11
                    ForeColor =16777215
                    Name ="BKK"
                    ControlSource ="BKK"
                End
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Left =3667
                    Top =4366
                    TabIndex =12
                    Name ="hes"
                    ControlSource ="hes"
                End
            End
        End
    End
End
CodeBehindForm
' See "P_HESAB_CHEKG.cls"
