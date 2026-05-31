Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    AllowUpdating =4
    DateGrouping =1
    GridX =16
    GridY =16
    Width =2731
    ItemSuffix =13
    Left =795
    Top =1110
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
            Height =1381
            OnPrint ="[Event Procedure]"
            Name ="Detail1"
            Begin
                Begin TextBox
                    Visible = NotDefault
                    TextAlign =2
                    TextFontFamily =34
                    Left =130
                    Top =135
                    Width =2478
                    Height =450
                    Name ="Barcode"
                    ControlSource ="Barcode"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =18
                    ReadingOrder =2
                    Left =465
                    Top =485
                    Width =1860
                    Height =225
                    FontSize =12
                    FontWeight =900
                    TabIndex =1
                    Name ="BarcodeContent"
                    ControlSource ="Barcode"
                    FontName ="Times New Roman"
                End
                Begin TextBox
                    TextFontCharSet =178
                    TextAlign =1
                    Left =135
                    Top =665
                    Width =675
                    Height =300
                    FontWeight =900
                    TabIndex =2
                    Name ="Text8"
                    ControlSource ="=[Forms]![stuf_def]![code]"
                    FontName ="Titr"
                End
                Begin Label
                    TextFontCharSet =178
                    Left =975
                    Top =680
                    Width =1635
                    Height =285
                    FontWeight =900
                    Name ="Label9"
                    Caption ="شركت تعاوني يزد باف"
                    FontName ="Sina"
                End
                Begin TextBox
                    TextFontCharSet =178
                    TextAlign =3
                    Left =90
                    Top =965
                    Width =2520
                    Height =300
                    FontSize =7
                    TabIndex =3
                    Name ="Text10"
                    ControlSource ="=[Forms]![stuf_def]![NAM]"
                    FontName ="Koodak"
                End
                Begin Rectangle
                    BorderWidth =2
                    Width =2731
                    Height =1346
                    Name ="Box12"
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
' See "rpt_BarCode1.cls"
