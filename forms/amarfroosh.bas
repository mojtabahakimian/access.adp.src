Version =20
VersionRequired =20
Begin Form
    OrderByOn = NotDefault
    DefaultView =2
    AllowUpdating =4
    TabularCharSet =178
    TabularFamily =0
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =4950
    DatasheetFontHeight =10
    ItemSuffix =18
    Left =2100
    Top =1470
    Right =9615
    Bottom =5205
    DatasheetGridlinesColor =12632256
    MaxRecords =100000
    OrderBy ="View60.mab DESC"
    RecSrcDt = Begin
        0x6ae5d54a1a2ce340
    End
    RecordSource ="View60"
    Caption ="آمار فروش و تخفيفات"
    DatasheetFontName ="Arial"
    RecordSourceQualifier ="dbo"
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
        Begin Section
            Height =4485
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =3
                    IMESentenceMode =3
                    Left =1927
                    Top =113
                    Width =2490
                    ColumnWidth =1230
                    ColumnOrder =0
                    Name ="CODE"
                    ControlSource ="CODE"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =113
                            Width =510
                            Height =240
                            Name ="Label1"
                            Caption ="كد كالا"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =3
                    IMESentenceMode =3
                    Left =1927
                    Top =453
                    Width =2490
                    ColumnWidth =2715
                    ColumnOrder =1
                    TabIndex =1
                    Name ="NAME"
                    ControlSource ="NAME"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =453
                            Width =495
                            Height =240
                            Name ="Label3"
                            Caption ="NAME"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =3
                    IMESentenceMode =3
                    Left =1927
                    Top =793
                    ColumnWidth =705
                    ColumnOrder =2
                    TabIndex =2
                    Name ="meg"
                    ControlSource ="meg"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =793
                            Width =375
                            Height =240
                            Name ="Label5"
                            Caption ="meg"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =3
                    IMESentenceMode =3
                    Left =1927
                    Top =1133
                    ColumnOrder =4
                    TabIndex =3
                    Name ="mab"
                    ControlSource ="mab"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =1133
                            Width =375
                            Height =240
                            Name ="Label7"
                            Caption ="mab"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =3
                    IMESentenceMode =3
                    Left =1927
                    Top =1474
                    ColumnWidth =1125
                    ColumnOrder =5
                    TabIndex =4
                    Name ="taf"
                    ControlSource ="taf"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =1474
                            Width =285
                            Height =240
                            Name ="Label9"
                            Caption ="taf"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =2
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =3
                    IMESentenceMode =3
                    Left =1927
                    Top =1834
                    Width =2496
                    ColumnWidth =1860
                    ColumnOrder =8
                    TabIndex =5
                    Name ="Text10"
                    ControlSource ="=([mbb])/[mab]*100"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =1830
                            Width =930
                            Height =240
                            Name ="Label11"
                            Caption ="درصد تخفيف"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1933
                    Top =2154
                    ColumnWidth =1020
                    ColumnOrder =6
                    TabIndex =6
                    Name ="mbb"
                    ControlSource ="mbb"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =90
                            Top =2205
                            Width =945
                            Height =240
                            Name ="Label12"
                            Caption ="مبلغ برگشت"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1933
                    Top =2565
                    ColumnWidth =1230
                    ColumnOrder =7
                    TabIndex =7
                    Name ="Text13"
                    ControlSource ="=[mbb]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =90
                            Top =2610
                            Width =1215
                            Height =240
                            Name ="Label14"
                            Caption ="برگشت و تخفيف"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =906
                    Top =3344
                    ColumnWidth =495
                    ColumnOrder =9
                    TabIndex =8
                    Name ="RADAH"
                    ControlSource ="RADAH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3648
                            Top =3344
                            Width =660
                            Height =240
                            Name ="Label15"
                            Caption ="گروه"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =226
                    Top =3798
                    ColumnWidth =1140
                    ColumnOrder =3
                    TabIndex =9
                    Name ="megmar"
                    ControlSource ="megmar"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2925
                            Top =3795
                            Width =1110
                            Height =240
                            Name ="Label16"
                            Caption ="مقدار مرجوعي"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1247
                    Top =4081
                    TabIndex =10
                    Name ="mnd"
                    ControlSource ="mnd"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4215
                            Top =4080
                            Width =735
                            Height =405
                            Name ="Label17"
                            Caption ="قيمت تمام شده"
                        End
                    End
                End
            End
        End
    End
End
