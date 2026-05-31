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
    Width =5345
    DatasheetFontHeight =10
    ItemSuffix =12
    Left =5595
    Top =2265
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x9fe4bf6ea242e340
    End
    RecordSource ="SALARY_BESTT"
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
            Height =348
            Name ="Detail"
            Begin
                Begin TextBox
                    TextAlign =3
                    NumeralShapes =2
                    Left =15
                    Width =1170
                    Height =330
                    FontSize =9
                    FontWeight =700
                    Name ="BES"
                    ControlSource ="BES"
                    Format ="#,###"
                    FontName ="Tahoma"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =10
                    NumeralShapes =2
                    Left =1230
                    Width =3735
                    Height =330
                    FontSize =9
                    FontWeight =700
                    TabIndex =1
                    Name ="VKI"
                    ControlSource ="SHR"
                    FontName ="Zar"
                End
                Begin TextBox
                    RunningSum =2
                    TextAlign =2
                    NumeralShapes =2
                    Left =5015
                    Width =330
                    Height =330
                    FontSize =9
                    FontWeight =700
                    TabIndex =2
                    Name ="C"
                    ControlSource ="=1"
                    FontName ="Tahoma"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Left =1218
                    Width =0
                    Height =348
                    Name ="Line13"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Left =4980
                    Width =0
                    Height =348
                    Name ="Line14"
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
                    NumeralShapes =2
                    Width =5070
                    Height =255
                    FontSize =9
                    FontWeight =700
                    Name ="SVAM"
                    ControlSource ="=Sum([BES])"
                    Format ="#,###"
                    FontName ="Tahoma"
                End
            End
        End
    End
End
CodeBehindForm
' See "SALARYBES2.cls"
