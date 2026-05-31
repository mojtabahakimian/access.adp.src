Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    AllowEdits = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    AllowUpdating =4
    ScrollBars =0
    ViewsAllowed =1
    TabularCharSet =178
    TabularFamily =0
    BorderStyle =3
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =8957
    DatasheetFontHeight =10
    ItemSuffix =5
    Left =525
    Right =11070
    Bottom =6330
    DatasheetGridlinesColor =12632256
    MaxRecords =100000
    RecSrcDt = Begin
        0x3797570d2e10e340
    End
    Caption ="محاسبات حقوق"
    DatasheetFontName ="Arial"
    AllowDatasheetView =0
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
        Begin Subform
            SpecialEffect =2
            Width =1701
            Height =1701
        End
        Begin Section
            CanGrow = NotDefault
            Height =6576
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin Subform
                    OverlapFlags =85
                    Top =60
                    Width =8880
                    Height =5895
                    Name ="SALARY_BES_SUB"
                    SourceObject ="Form.SALARY_BES_SUB"
                End
                Begin TextBox
                    FELineBreak = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =42
                    NumeralShapes =2
                    Left =2130
                    Top =6064
                    Width =1770
                    Height =330
                    FontSize =9
                    FontWeight =700
                    TabIndex =1
                    Name ="BES"
                    ControlSource ="=[SALARY_BES_SUB].[Form]![SBES]"
                    Format ="#,###"
                    AsianLineBreak =0
                End
                Begin TextBox
                    FELineBreak = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =42
                    NumeralShapes =2
                    Left =329
                    Top =6064
                    Width =1770
                    Height =330
                    FontSize =9
                    FontWeight =700
                    TabIndex =2
                    Name ="BED"
                    ControlSource ="=[SALARY_BES_SUB].[Form]![SBED]"
                    Format ="#,###"
                    AsianLineBreak =0
                End
                Begin TextBox
                    FELineBreak = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =42
                    NumeralShapes =2
                    Left =5745
                    Top =6047
                    Width =1770
                    Height =330
                    FontSize =9
                    FontWeight =700
                    TabIndex =3
                    Name ="Text4"
                    ControlSource ="=[SALARY_BES_SUB].[Form]![SBES]-[SALARY_BES_SUB].[Form]![SBED]"
                    Format ="#,###"
                    AsianLineBreak =0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =7565
                            Top =6045
                            Width =1275
                            Height =240
                            Name ="Label13"
                            Caption ="حقوق پرداختني :"
                        End
                    End
                End
            End
        End
    End
End
