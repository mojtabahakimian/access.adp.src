Version =20
VersionRequired =20
Begin Form
    AllowDeletions = NotDefault
    AllowAdditions = NotDefault
    AllowEdits = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =2
    AllowUpdating =4
    ViewsAllowed =2
    TabularCharSet =178
    TabularFamily =2
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =5669
    RowHeight =360
    DatasheetFontHeight =12
    DatasheetFontWeight =700
    ItemSuffix =16
    Left =345
    Top =690
    Right =8955
    Bottom =6315
    DatasheetGridlinesColor =12632256
    MaxRecords =100000
    RecSrcDt = Begin
        0x942be3492e10e340
    End
    RecordSource ="SELECT  HES_K, HES_M, HES_T, BED, BES, MM, DATE_S, dbo.FISHSTR(ISNULL(SHARH, N'ب"
        "دون شرح')) AS SHR FROM dbo.SALARY_BEDS(12, 213,84) SALARY_BEDS WHERE  HES_T = 2 "
        "ORDER BY BES DESC, BED DESC"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="B Nazanin"
    AllowFormView =0
    Begin
        Begin Label
            BackStyle =0
            TextFontCharSet =178
            FontName ="Tahoma"
        End
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            OldBorderStyle =0
            TextFontCharSet =178
            Width =1701
            LabelX =-1701
            FontName ="Tahoma"
            AsianLineBreak =255
        End
        Begin FormHeader
            Height =0
            BackColor =-2147483633
            Name ="FormHeader"
        End
        Begin Section
            Height =3480
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    FELineBreak = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =42
                    BackStyle =0
                    NumeralShapes =2
                    Width =1155
                    Height =330
                    ColumnWidth =1800
                    ColumnOrder =1
                    FontSize =9
                    FontWeight =700
                    Name ="BES"
                    ControlSource ="BES"
                    Format ="#,###"
                    AsianLineBreak =0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2955
                            Top =1440
                            Width =600
                            Height =240
                            Name ="Label9"
                            Caption ="درآمدها"
                        End
                    End
                End
                Begin TextBox
                    FELineBreak = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =10
                    BackStyle =0
                    NumeralShapes =2
                    Left =1170
                    Width =3645
                    Height =330
                    ColumnWidth =4710
                    ColumnOrder =0
                    FontSize =9
                    FontWeight =700
                    TabIndex =1
                    Name ="VKI"
                    ControlSource ="SHR"
                    FontName ="Zar"
                    AsianLineBreak =0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2955
                            Top =1080
                            Width =420
                            Height =240
                            Name ="Label8"
                            Caption ="شرح"
                        End
                    End
                End
                Begin TextBox
                    FELineBreak = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =42
                    BackStyle =0
                    NumeralShapes =2
                    Top =1800
                    Width =1155
                    Height =330
                    ColumnWidth =1725
                    ColumnOrder =2
                    FontSize =9
                    FontWeight =700
                    TabIndex =2
                    Name ="BED"
                    ControlSource ="BED"
                    Format ="#,###"
                    AsianLineBreak =0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2955
                            Top =3240
                            Width =645
                            Height =240
                            Name ="Label13"
                            Caption ="كسورات"
                        End
                    End
                End
            End
        End
        Begin FormFooter
            Height =1680
            BackColor =-2147483633
            Name ="FormFooter"
            Begin
                Begin TextBox
                    FELineBreak = NotDefault
                    SpecialEffect =0
                    OverlapFlags =93
                    TextAlign =3
                    TextFontFamily =42
                    BackStyle =0
                    NumeralShapes =2
                    Width =1155
                    Height =330
                    FontSize =9
                    FontWeight =700
                    Name ="SBES"
                    ControlSource ="=Sum([BES])"
                    Format ="#,###"
                    AsianLineBreak =0
                End
                Begin TextBox
                    FELineBreak = NotDefault
                    SpecialEffect =0
                    OverlapFlags =247
                    TextAlign =3
                    TextFontFamily =42
                    BackStyle =0
                    NumeralShapes =2
                    Left =1133
                    Width =1155
                    Height =330
                    FontSize =9
                    FontWeight =700
                    TabIndex =1
                    Name ="SBED"
                    ControlSource ="=Sum([BED])"
                    Format ="#,###"
                    AsianLineBreak =0
                End
            End
        End
    End
End
CodeBehindForm
' See "SALARY_BES_SUB.cls"
