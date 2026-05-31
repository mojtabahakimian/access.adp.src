Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    AllowUpdating =4
    DateGrouping =1
    GridX =16
    GridY =16
    Width =2343
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
            Height =1455
            OnFormat ="[Event Procedure]"
            OnPrint ="[Event Procedure]"
            Name ="Detail1"
            Begin
                Begin TextBox
                    TextFontCharSet =178
                    TextAlign =2
                    BackStyle =0
                    Top =1050
                    Width =2291
                    Height =405
                    FontSize =11
                    FontWeight =700
                    Name ="nam"
                    ControlSource ="nam"
                    FontName ="Titr"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =2
                    BackStyle =0
                    NumeralShapes =1
                    Left =35
                    Width =2306
                    Height =405
                    FontSize =10
                    FontWeight =700
                    TabIndex =1
                    Name ="Text15"
                    ControlSource ="=\"قيمت با تخفيف:  \" & [mab]"
                    Format ="Standard"
                    FontName ="Titr"
                End
                Begin TextBox
                    Visible = NotDefault
                    TextAlign =2
                    TextFontFamily =34
                    Left =15
                    Top =390
                    Width =2298
                    Height =645
                    TabIndex =2
                    Name ="Barcode"
                    ControlSource ="Barcode"
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
' See "rpt_barcode_11.cls"
