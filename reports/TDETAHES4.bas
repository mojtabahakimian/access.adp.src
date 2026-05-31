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
    DatasheetFontHeight =10
    ItemSuffix =2
    Left =330
    Top =1290
    DatasheetGridlinesColor =12632256
    MaxRecords =10000000
    RecSrcDt = Begin
        0x93926dea733be440
    End
    RecordSource ="SELECT TDETA_HES4.*, RTRIM(CAST(N_KOL AS nvarchar)) + '-' + RTRIM(CAST(NUMBER AS"
        " nvarchar)) + '-' + RTRIM(CAST(TNUMBER AS nvarchar)) + '-' + RTRIM(CAST(TNUMBER2"
        " AS nvarchar)) + '-' + RTRIM(CAST(TNUMBER3 AS nvarchar)) + '-' + RTRIM(CAST(TNUM"
        "BER4 AS nvarchar)) AS HES FROM TDETA_HES4"
    DatasheetFontName ="Arial"
    RecordSourceQualifier ="dbo"
    Begin
        Begin Label
            BackStyle =0
            TextFontFamily =2
            FontName ="Arial"
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
        Begin Section
            KeepTogether = NotDefault
            Height =390
            Name ="Detail"
            Begin
                Begin TextBox
                    TextFontCharSet =178
                    TextAlign =1
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =5448
                    Height =390
                    FontSize =10
                    Name ="HES"
                    ControlSource ="HES"
                    FontName ="Titr"
                End
                Begin TextBox
                    TextFontCharSet =178
                    TextAlign =2
                    ReadingOrder =1
                    ScrollBarAlign =1
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =54
                    Width =5331
                    Height =390
                    FontSize =10
                    FontWeight =700
                    TabIndex =1
                    Name ="NAME"
                    ControlSource ="NAME"
                    FontName ="Titr"
                End
            End
        End
    End
End
