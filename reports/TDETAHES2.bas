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
    ItemSuffix =4
    Left =360
    Top =3675
    DatasheetGridlinesColor =12632256
    MaxRecords =10000000
    RecSrcDt = Begin
        0x18270d8f723be440
    End
    RecordSource ="SELECT TDETA_HES2.*, RTRIM(CAST(N_KOL AS nvarchar)) + '-' + RTRIM(CAST(NUMBER AS"
        " nvarchar)) + '-' + RTRIM(CAST(TNUMBER AS nvarchar)) + '-' + RTRIM(CAST(TNUMBER2"
        " AS nvarchar)) AS HES FROM TDETA_HES2"
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
        Begin Subform
            Width =1701
            Height =1701
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =300
            BackColor =13434828
            Name ="ReportHeader"
            Begin
                Begin Label
                    TextFontCharSet =178
                    Left =6121
                    Width =570
                    Height =300
                    FontSize =10
                    FontWeight =700
                    Name ="Label0"
                    Caption ="حساب"
                    FontName ="Titr"
                End
                Begin Label
                    TextFontCharSet =178
                    Left =3061
                    Width =390
                    Height =300
                    FontSize =10
                    FontWeight =700
                    Name ="Label1"
                    Caption ="نام "
                    FontName ="Titr"
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            CanGrow = NotDefault
            Height =449
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
                Begin Subform
                    OldBorderStyle =0
                    Top =405
                    Width =7200
                    Height =44
                    TabIndex =2
                    Name ="TDETAHES3"
                    SourceObject ="Report.TDETAHES3"
                    LinkChildFields ="N_KOL;NUMBER;TNUMBER;TNUMBER2"
                    LinkMasterFields ="N_KOL;NUMBER;TNUMBER;TNUMBER2"
                End
            End
        End
        Begin FormFooter
            KeepTogether = NotDefault
            Height =0
            Name ="ReportFooter"
        End
    End
End
