Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    OrderByOn = NotDefault
    DefaultView =0
    AllowUpdating =4
    TabularCharSet =178
    TabularFamily =42
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =60
    GridY =60
    Width =4830
    DatasheetFontHeight =10
    ItemSuffix =13
    Left =5655
    Top =2385
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xbcbd2fbfd2b9e240
    End
    RecordSource ="SALARY_BESTT"
    Caption ="VAMS1"
    DatasheetFontName ="Arial (Arabic)"
    RecordSourceQualifier ="dbo"
    Begin
        Begin Label
            BackStyle =0
            TextFontCharSet =178
            TextAlign =1
            TextFontFamily =42
            FontSize =9
            FontWeight =700
            FontName ="Arial (Arabic)"
        End
        Begin Rectangle
            BackStyle =0
            Width =850
            Height =850
        End
        Begin Line
            Width =1701
        End
        Begin Image
            OldBorderStyle =0
            PictureAlignment =2
            Width =1701
            Height =1701
        End
        Begin CommandButton
            TextFontCharSet =178
            Width =1701
            Height =283
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Tahoma"
        End
        Begin OptionButton
            LabelX =230
            LabelY =-30
        End
        Begin CheckBox
            LabelX =230
            LabelY =-30
        End
        Begin OptionGroup
            BackStyle =1
            Width =1701
            Height =1701
        End
        Begin BoundObjectFrame
            Width =4536
            Height =2835
            LabelX =-1701
        End
        Begin TextBox
            OldBorderStyle =0
            TextFontCharSet =178
            TextFontFamily =42
            BackStyle =0
            Width =1701
            LabelX =-1701
            FontName ="Arial (Arabic)"
        End
        Begin ListBox
            TextFontCharSet =178
            TextFontFamily =42
            OldBorderStyle =0
            Width =1701
            Height =1417
            LabelX =-1701
            FontName ="Arial (Arabic)"
        End
        Begin ComboBox
            OldBorderStyle =0
            TextFontCharSet =178
            TextFontFamily =42
            BackStyle =0
            Width =1701
            LabelX =-1701
            FontName ="Arial (Arabic)"
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
            SortOrder = NotDefault
            ControlSource ="BES"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Visible = NotDefault
            Height =0
            Name ="ReportHeader"
        End
        Begin Section
            KeepTogether = NotDefault
            Height =330
            Name ="Detail"
            Begin
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    NumeralShapes =2
                    Left =15
                    Width =1155
                    Height =330
                    FontSize =9
                    Name ="BES"
                    ControlSource ="BES"
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    NumeralShapes =2
                    Left =1185
                    Width =3645
                    Height =330
                    FontSize =9
                    TabIndex =1
                    Name ="VKI"
                    ControlSource ="SHR"
                    FontName ="WeblogmaYekan"
                End
            End
        End
        Begin FormFooter
            KeepTogether = NotDefault
            Visible = NotDefault
            Height =311
            Name ="ReportFooter"
            Begin
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    NumeralShapes =2
                    Width =4665
                    Height =255
                    FontSize =9
                    Name ="SVAM"
                    ControlSource ="=Sum([BES])"
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                End
            End
        End
    End
End
CodeBehindForm
' See "SALARYBEST.cls"
