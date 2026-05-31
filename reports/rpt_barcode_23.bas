Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    AllowUpdating =4
    DateGrouping =1
    GridX =16
    GridY =16
    Width =2551
    ItemSuffix =16
    Left =390
    Top =1125
    RecSrcDt = Begin
        0x6ab78456eddae340
    End
    RecordSource ="tbl_BarCodes"
    DatasheetFontName ="Arial"
    RecordSourceQualifier ="dbo"
    Begin
        Begin Label
            TextFontFamily =2
            FontName ="Arial"
        End
        Begin Rectangle
            BackStyle =0
            Width =850
            Height =850
        End
        Begin TextBox
            OldBorderStyle =0
            TextFontFamily =2
            FontName ="Arial"
        End
        Begin PageHeader
            Height =0
            Name ="PageHeader0"
        End
        Begin Section
            KeepTogether = NotDefault
            Height =1531
            OnFormat ="[Event Procedure]"
            OnPrint ="[Event Procedure]"
            Name ="Detail1"
            Begin
                Begin TextBox
                    TextFontCharSet =178
                    TextAlign =2
                    ReadingOrder =2
                    BackStyle =0
                    Left =30
                    Top =525
                    Width =2486
                    Height =405
                    FontSize =12
                    FontWeight =700
                    Name ="BarcodeContent"
                    ControlSource ="=[Forms]![stuf_def]![B_SEF] & \" - \" & [Barcode]"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =2
                    BackStyle =0
                    NumeralShapes =1
                    Left =30
                    Top =915
                    Width =2456
                    Height =405
                    FontSize =12
                    FontWeight =700
                    TabIndex =1
                    Name ="nam"
                    ControlSource ="=[Forms]![stuf_def]![NAM]"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    TextFontCharSet =178
                    TextAlign =2
                    Top =90
                    Width =2478
                    Height =450
                    TabIndex =2
                    Name ="Barcode"
                    ControlSource ="Barcode"
                    FontName ="WeblogmaYekan"
                End
            End
        End
        Begin PageFooter
            Height =0
            Name ="PageFooter2"
        End
    End
End
CodeBehindForm
' See "rpt_barcode_23.cls"
