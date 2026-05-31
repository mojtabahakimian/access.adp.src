Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    AllowUpdating =4
    TabularCharSet =178
    TabularFamily =0
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =10125
    DatasheetFontHeight =10
    ItemSuffix =24
    Left =1815
    Top =7965
    DatasheetGridlinesColor =12632256
    MaxRecords =10000000
    RecSrcDt = Begin
        0x1716c98c7345e440
    End
    RecordSource ="SELECT TARAZ_DTL_1.N_KOL, TARAZ_DTL_1.NUMBER, TARAZ_DTL_1.TNUMBER, TARAZ_DTL_1.H"
        "ES_T2, SUM(TARAZ_DTL_1.BED) AS SumOfBED, SUM(TARAZ_DTL_1.BES) AS SumOfBES, TARAZ"
        "_DTL_1.ECODE, SUM(TARAZ_DTL_1.BED - TARAZ_DTL_1.BES) AS mand, CAST(TDETA_HES4.TN"
        "UMBER4 AS NVARCHAR) + N' - ' + CAST(TDETA_HES4.TNUMBER3 AS NVARCHAR) + N' - ' + "
        "CAST(TDETA_HES4.TNUMBER2 AS NVARCHAR) + N' - ' + CAST(TDETA_HES4.TNUMBER AS NVAR"
        "CHAR) + N' - ' + CAST(TDETA_HES4.NUMBER AS NVARCHAR) + N' - ' + CAST(TDETA_HES4."
        "N_KOL AS NVARCHAR) + N' - ' + TDETA_HES4.NAME AS t3, TARAZ_DTL_1.HES_T3, TARAZ_D"
        "TL_1.HES_T4 FROM TARAZ_DTL_1 INNER JOIN TDETA_HES4 ON TARAZ_DTL_1.N_KOL = TDETA_"
        "HES4.N_KOL AND TARAZ_DTL_1.NUMBER = TDETA_HES4.NUMBER AND TARAZ_DTL_1.TNUMBER = "
        "TDETA_HES4.TNUMBER AND TARAZ_DTL_1.HES_T2 = TDETA_HES4.TNUMBER2 AND TARAZ_DTL_1."
        "HES_T3 = TDETA_HES4.TNUMBER3 AND TARAZ_DTL_1.HES_T4 = TDETA_HES4.TNUMBER4 GROUP "
        "BY TARAZ_DTL_1.N_KOL, TARAZ_DTL_1.NUMBER, TARAZ_DTL_1.TNUMBER, TARAZ_DTL_1.ECODE"
        ", TARAZ_DTL_1.HES_T2, CAST(TDETA_HES4.TNUMBER4 AS NVARCHAR) + N' - ' + CAST(TDET"
        "A_HES4.TNUMBER3 AS NVARCHAR) + N' - ' + CAST(TDETA_HES4.TNUMBER2 AS NVARCHAR) + "
        "N' - ' + CAST(TDETA_HES4.TNUMBER AS NVARCHAR) + N' - ' + CAST(TDETA_HES4.NUMBER "
        "AS NVARCHAR) + N' - ' + CAST(TDETA_HES4.N_KOL AS NVARCHAR) + N' - ' + TDETA_HES4"
        ".NAME, TARAZ_DTL_1.HES_T3, TARAZ_DTL_1.HES_T4"
    DatasheetFontName ="Arial"
    RecordSourceQualifier ="dbo"
    Begin
        Begin Label
            BackStyle =0
            TextFontFamily =2
            FontName ="Arial"
        End
        Begin Line
            Width =1701
        End
        Begin TextBox
            FELineBreak = NotDefault
            OldBorderStyle =0
            TextFontFamily =2
            Width =1701
            LabelX =-1701
            FontName ="Arial"
            AsianLineBreak =255
        End
        Begin Subform
            Width =1701
            Height =1701
        End
        Begin FormHeader
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =288
            BackColor =12632256
            Name ="ReportHeader"
            Begin
                Begin Label
                    TextFontCharSet =178
                    Left =6045
                    Width =795
                    Height =285
                    FontWeight =700
                    Name ="Label16"
                    Caption ="مبالغ جزء..."
                    FontName ="Titr"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Left =5987
                    Width =0
                    Height =288
                    Name ="Line17"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Left =4504
                    Width =0
                    Height =288
                    Name ="Line18"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Left =3004
                    Width =0
                    Height =288
                    Name ="Line19"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Left =1519
                    Width =0
                    Height =288
                    Name ="Line20"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =2
                    Width =0
                    Height =288
                    Name ="Line21"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =2
                    Left =10093
                    Width =0
                    Height =288
                    Name ="Line22"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Left =9431
                    Width =0
                    Height =288
                    Name ="Line23"
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =426
            BackColor =12632256
            Name ="Detail"
            Begin
                Begin TextBox
                    FELineBreak = NotDefault
                    TextFontCharSet =178
                    TextAlign =1
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    BackStyle =0
                    NumeralShapes =2
                    Left =6018
                    Width =3399
                    Height =375
                    BackColor =10079487
                    BorderColor =12632256
                    Name ="t3"
                    ControlSource ="t3"
                    FontName ="Zar"
                    AsianLineBreak =0
                End
                Begin TextBox
                    FELineBreak = NotDefault
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    BackStyle =0
                    NumeralShapes =2
                    Left =4512
                    Width =1449
                    Height =375
                    FontSize =10
                    TabIndex =1
                    BackColor =10079487
                    BorderColor =12632256
                    Name ="SumOfBED"
                    ControlSource ="SumOfBED"
                    Format ="#,###"
                    FontName ="Zar"
                    AsianLineBreak =0
                End
                Begin TextBox
                    FELineBreak = NotDefault
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    BackStyle =0
                    NumeralShapes =2
                    Left =3023
                    Width =1449
                    Height =375
                    FontSize =10
                    TabIndex =2
                    BackColor =10079487
                    BorderColor =12632256
                    Name ="SumOfBES"
                    ControlSource ="SumOfBES"
                    Format ="#,###"
                    FontName ="Zar"
                    AsianLineBreak =0
                End
                Begin TextBox
                    FELineBreak = NotDefault
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    BackStyle =0
                    NumeralShapes =2
                    Left =1547
                    Width =1434
                    Height =375
                    FontSize =10
                    TabIndex =3
                    BackColor =10079487
                    BorderColor =12632256
                    Name ="bed"
                    ControlSource ="=IIf([mand]>0,[mand],0)"
                    Format ="#,###"
                    FontName ="Zar"
                    AsianLineBreak =0
                End
                Begin TextBox
                    FELineBreak = NotDefault
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    BackStyle =0
                    NumeralShapes =2
                    Left =44
                    Width =1446
                    Height =375
                    FontSize =10
                    TabIndex =4
                    BackColor =52479
                    BorderColor =12632256
                    Name ="bes"
                    ControlSource ="=IIf([mand]<0,Abs([mand]),0)"
                    Format ="#,###"
                    FontName ="Zar"
                    AsianLineBreak =0
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Left =5987
                    Top =3
                    Width =0
                    Height =423
                    Name ="Line24"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Left =4504
                    Top =3
                    Width =0
                    Height =423
                    Name ="Line25"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Left =3004
                    Top =3
                    Width =0
                    Height =423
                    Name ="Line26"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Left =1519
                    Top =3
                    Width =0
                    Height =423
                    Name ="Line27"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =2
                    Top =3
                    Width =0
                    Height =423
                    Name ="Line29"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =2
                    Left =10093
                    Top =3
                    Width =0
                    Height =423
                    Name ="Line31"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Left =9431
                    Top =3
                    Width =0
                    Height =423
                    Name ="Line76"
                End
                Begin TextBox
                    FELineBreak = NotDefault
                    RunningSum =2
                    TextFontCharSet =178
                    TextAlign =1
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    BackStyle =0
                    NumeralShapes =2
                    Left =9450
                    Top =3
                    Width =594
                    Height =313
                    FontSize =10
                    TabIndex =5
                    BackColor =10079487
                    BorderColor =12632256
                    Name ="Text106"
                    ControlSource ="=1"
                    FontName ="Zar"
                    AsianLineBreak =0
                End
            End
        End
        Begin FormFooter
            KeepTogether = NotDefault
            Height =0
            BackColor =52479
            Name ="ReportFooter"
        End
    End
End
CodeBehindForm
' See "R_TARAZ_TAFZ4_4.cls"
