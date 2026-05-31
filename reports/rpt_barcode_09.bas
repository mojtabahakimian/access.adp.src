Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    AllowUpdating =4
    DateGrouping =1
    GridX =16
    GridY =16
    Width =1901
    ItemSuffix =13
    Left =390
    Top =345
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
            Height =780
            OnPrint ="[Event Procedure]"
            Name ="Detail1"
            Begin
                Begin TextBox
                    Visible = NotDefault
                    TextFontCharSet =178
                    TextAlign =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    BackStyle =0
                    NumeralShapes =1
                    Left =45
                    Width =1781
                    Height =390
                    Name ="Barcode"
                    ControlSource ="Barcode"
                    FontName ="Badr"
                End
                Begin TextBox
                    TextFontCharSet =178
                    TextAlign =2
                    ReadingOrder =2
                    BackStyle =0
                    Left =30
                    Top =315
                    Width =1871
                    Height =285
                    TabIndex =1
                    Name ="BarcodeContent"
                    ControlSource ="=[Forms]![stuf_def]![B_SEF] & \" - \" & [Barcode]"
                    FontName ="Koodak"
                End
                Begin TextBox
                    TextFontCharSet =178
                    TextAlign =2
                    BackStyle =0
                    Top =435
                    Width =1871
                    Height =315
                    FontSize =7
                    TabIndex =2
                    Name ="Text10"
                    ControlSource ="=[Forms]![stuf_def]![NAM]"
                    FontName ="Koodak"
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
' See "rpt_barcode_09.cls"
