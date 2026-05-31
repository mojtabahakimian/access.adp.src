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
    ItemSuffix =25
    Left =705
    Top =1035
    DatasheetGridlinesColor =12632256
    Tag ="@Forms___CHK_V_PRINT___CHK_V_PRINTSUB___Form___N_SERI float = [Forms]![CHK_V_PRI"
        "NT]![CHK_V_PRINTSUB].[Form]![N_SERI] ,\015\012 @Forms___CHK_V_PRINT___CHK_V_PRIN"
        "TSUB___Form___DATE_S  = [Forms]![CHK_V_PRINT]![CHK_V_PRINTSUB].[Form]![DATE_S] ,"
        "\015\012 @Forms___CHK_V_PRINT___CHK_V_PRINTSUB___Form___BANK  = [Forms]![CHK_V_P"
        "RINT]![CHK_V_PRINTSUB].[Form]![BANK]"
    RecSrcDt = Begin
        0x0f48d609fa6be340
    End
    InputParameters ="@Forms___CHK_V_PRINT___CHK_V_PRINTSUB___Form___N_SERI BIGINT = [Forms]![CHK_V_PR"
        "INT]![CHK_V_PRINTSUB].[Form]![N_SERI] ,\015\012 @Forms___CHK_V_PRINT___CHK_V_PRI"
        "NTSUB___Form___DATE_S  = [Forms]![CHK_V_PRINT]![CHK_V_PRINTSUB].[Form]![TEXT34] "
        ",\015\012 @Forms___CHK_V_PRINT___CHK_V_PRINTSUB___Form___BANK  = [Forms]![CHK_V_"
        "PRINT]![CHK_V_PRINTSUB].[Form]![BANK]"
    RecordSource ="CHEK_HES"
    Caption ="چك قابل چاپ"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial (Arabic)"
    UnknownProp = {258 ,66 ,12 ,4 ,690 } Begin
        0x400046006f0072006d0073005f005f005f00430048004b005f0056005f005000 ,
        0x520049004e0054005f005f005f00430048004b005f0056005f00500052004900 ,
        0x4e0054005300550042005f005f005f0046006f0072006d005f005f005f004e00 ,
        0x5f005300450052004900200042004900470049004e00540020003d0020005b00 ,
        0x46006f0072006d0073005d0021005b00430048004b005f0056005f0050005200 ,
        0x49004e0054005d0021005b00430048004b005f0056005f005000520049004e00 ,
        0x54005300550042005d002e005b0046006f0072006d005d0021005b004e005f00 ,
        0x53004500520049005d0020002c000d000a002000400046006f0072006d007300 ,
        0x5f005f005f00430048004b005f0056005f005000520049004e0054005f005f00 ,
        0x5f00430048004b005f0056005f005000520049004e0054005300550042005f00 ,
        0x5f005f0046006f0072006d005f005f005f0044004100540045005f0053002000 ,
        0x20003d0020005b0046006f0072006d0073005d0021005b00430048004b005f00 ,
        0x56005f005000520049004e0054005d0021005b00430048004b005f0056005f00 ,
        0x5000520049004e0054005300550042005d002e005b0046006f0072006d005d00 ,
        0x21005b005400450058005400330034005d0020002c000d000a00200040004600 ,
        0x6f0072006d0073005f005f005f00430048004b005f0056005f00500052004900 ,
        0x4e0054005f005f005f00430048004b005f0056005f005000520049004e005400 ,
        0x5300550042005f005f005f0046006f0072006d005f005f005f00420041004e00 ,
        0x4b00200020003d0020005b0046006f0072006d0073005d0021005b0043004800 ,
        0x4b005f0056005f005000520049004e0054005d0021005b00430048004b005f00 ,
        0x56005f005000520049004e0054005300550042005d002e005b0046006f007200 ,
        0x6d005d0021005b00420041004e004b005d00
    End
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
                    ControlSource ="=[Forms]![Baseknow]![MANAGER]"
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
                    ControlSource ="SHOBEH"
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
                    ControlSource ="sho"
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
                    ControlSource ="=[Forms]![CHK_V_PRINT]![CITY]"
                    FontName ="Zar"
                End
            End
        End
    End
End
CodeBehindForm
' See "P_HESAB_CHEK.cls"
