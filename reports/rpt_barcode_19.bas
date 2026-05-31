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
            Height =1701
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
                    FontSize =11
                    Name ="BarcodeContent"
                    ControlSource ="=[code] & \" - \" & [Barcode]"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    TextFontCharSet =178
                    TextAlign =2
                    BackStyle =0
                    Top =1245
                    Width =2486
                    Height =405
                    FontSize =11
                    TabIndex =1
                    Name ="nam"
                    ControlSource ="nam"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    BackStyle =0
                    NumeralShapes =2
                    Left =30
                    Top =915
                    Width =2456
                    Height =405
                    FontSize =12
                    TabIndex =2
                    Name ="Text15"
                    ControlSource ="mab"
                    Format ="#,###\" ريال\""
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    TextFontCharSet =178
                    TextAlign =2
                    Top =90
                    Width =2268
                    Height =450
                    TabIndex =3
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
' See "rpt_barcode_19.cls"
