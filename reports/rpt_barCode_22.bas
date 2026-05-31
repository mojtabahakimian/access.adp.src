Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    AllowUpdating =4
    DateGrouping =1
    GridX =16
    GridY =16
    Width =1948
    ItemSuffix =18
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
            Height =1531
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
                    Left =15
                    Top =345
                    Width =1781
                    Height =390
                    Name ="Barcode"
                    ControlSource ="Barcode"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    TextFontCharSet =178
                    TextAlign =2
                    ReadingOrder =2
                    BackStyle =0
                    Left =30
                    Top =750
                    Width =1871
                    Height =285
                    FontSize =10
                    TabIndex =1
                    Name ="BarcodeContent"
                    ControlSource ="=[Forms]![stuf_def]![B_SEF] & \" - \" & [Barcode]"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    OverlapFlags =4
                    TextFontCharSet =178
                    TextAlign =2
                    BackStyle =0
                    Top =1020
                    Width =1946
                    Height =330
                    FontSize =10
                    TabIndex =2
                    Name ="Text10"
                    ControlSource ="=[Forms]![stuf_def]![NAM]"
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
' See "rpt_barCode_22.cls"
