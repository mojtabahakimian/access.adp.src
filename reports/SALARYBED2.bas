Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
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
    Width =5119
    DatasheetFontHeight =10
    ItemSuffix =15
    Left =345
    Top =2265
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x04e0017aa242e340
    End
    RecordSource ="SALARY_BEDTT"
    Caption ="VAMS1"
    DatasheetFontName ="Arial (Arabic)"
    OnError ="[Event Procedure]"
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
            ControlSource ="BED"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Visible = NotDefault
            Height =0
            Name ="ReportHeader"
        End
        Begin Section
            KeepTogether = NotDefault
            Height =348
            Name ="Detail"
            Begin
                Begin TextBox
                    RunningSum =2
                    TextAlign =2
                    NumeralShapes =2
                    Left =4789
                    Width =330
                    Height =330
                    FontSize =9
                    FontWeight =700
                    Name ="C"
                    ControlSource ="=1"
                    FontName ="Tahoma"
                End
                Begin TextBox
                    TextAlign =3
                    NumeralShapes =2
                    Width =1005
                    Height =330
                    FontSize =9
                    FontWeight =700
                    TabIndex =1
                    Name ="BED"
                    ControlSource ="BED"
                    Format ="#,###"
                    FontName ="Tahoma"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Left =1035
                    Width =0
                    Height =348
                    Name ="Line13"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =10
                    Left =1077
                    Width =3646
                    Height =331
                    FontSize =9
                    FontWeight =700
                    TabIndex =2
                    Name ="VKI"
                    ControlSource ="SHR"
                    FontName ="Zar"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Left =4755
                    Width =0
                    Height =348
                    Name ="Line14"
                End
            End
        End
        Begin FormFooter
            KeepTogether = NotDefault
            Visible = NotDefault
            Height =331
            Name ="ReportFooter"
            Begin
                Begin TextBox
                    TextAlign =3
                    NumeralShapes =2
                    Left =606
                    Width =1860
                    Height =255
                    FontSize =9
                    FontWeight =700
                    Name ="SVAM"
                    ControlSource ="=Sum([BED])"
                    Format ="#,###"
                    FontName ="Tahoma"
                End
            End
        End
    End
End
CodeBehindForm
' See "SALARYBED2.cls"
