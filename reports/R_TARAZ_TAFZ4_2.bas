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
    Left =330
    Top =4290
    DatasheetGridlinesColor =12632256
    MaxRecords =10000000
    RecSrcDt = Begin
        0xbb2cffef7245e440
    End
    RecordSource ="SELECT TARAZ_DTL_1.N_KOL, TARAZ_DTL_1.NUMBER, TARAZ_DTL_1.TNUMBER, TARAZ_DTL_1.H"
        "ES_T2, SUM(TARAZ_DTL_1.BED) AS SumOfBED, SUM(TARAZ_DTL_1.BES) AS SumOfBES, TARAZ"
        "_DTL_1.ECODE, SUM(TARAZ_DTL_1.BED - TARAZ_DTL_1.BES) AS mand, CAST(TDETA_HES2.TN"
        "UMBER2 AS NVARCHAR) + N' - ' + CAST(TDETA_HES2.TNUMBER AS NVARCHAR) + N' - ' + C"
        "AST(TDETA_HES2.NUMBER AS NVARCHAR) + N' - ' + CAST(TDETA_HES2.N_KOL AS NVARCHAR)"
        " + N' - ' + TDETA_HES2.NAME AS t2 FROM TDETA_HES2 INNER JOIN TARAZ_DTL_1 ON TDET"
        "A_HES2.N_KOL = TARAZ_DTL_1.N_KOL AND TDETA_HES2.NUMBER = TARAZ_DTL_1.NUMBER AND "
        "TDETA_HES2.TNUMBER = TARAZ_DTL_1.TNUMBER AND TDETA_HES2.TNUMBER2 = TARAZ_DTL_1.H"
        "ES_T2 GROUP BY TARAZ_DTL_1.N_KOL, TARAZ_DTL_1.NUMBER, TARAZ_DTL_1.TNUMBER, TARAZ"
        "_DTL_1.ECODE, TARAZ_DTL_1.HES_T2, CAST(TDETA_HES2.TNUMBER2 AS NVARCHAR) + N' - '"
        " + CAST(TDETA_HES2.TNUMBER AS NVARCHAR) + N' - ' + CAST(TDETA_HES2.NUMBER AS NVA"
        "RCHAR) + N' - ' + CAST(TDETA_HES2.N_KOL AS NVARCHAR) + N' - ' + TDETA_HES2.NAME"
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
            BackColor =52479
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
            BackColor =52479
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
                    Width =3339
                    Height =375
                    BackColor =10079487
                    BorderColor =12632256
                    Name ="t2"
                    ControlSource ="t2"
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
                Begin Subform
                    CanShrink = NotDefault
                    OldBorderStyle =0
                    Top =360
                    Width =10125
                    Height =63
                    TabIndex =6
                    Name ="R_TARAZ_TAFZ4_3"
                    SourceObject ="Report.R_TARAZ_TAFZ4_3"
                    LinkChildFields ="N_KOL;NUMBER;TNUMBER;HES_T2"
                    LinkMasterFields ="N_KOL;NUMBER;TNUMBER;HES_T2"
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
' See "R_TARAZ_TAFZ4_2.cls"
