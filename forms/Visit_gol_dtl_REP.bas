Version =20
VersionRequired =20
Begin Form
    DefaultView =2
    AllowUpdating =4
    ViewsAllowed =2
    TabularCharSet =178
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =8361
    RowHeight =330
    ItemSuffix =42
    Left =465
    Top =1830
    Right =14985
    Bottom =9825
    DatasheetForeColor =33554432
    DatasheetGridlinesColor =12632256
    MaxRecords =1000000
    RecSrcDt = Begin
        0xf84fad3ad162e440
    End
    DatasheetFontName ="Tahoma"
    AllowFormView =0
    Begin
        Begin Label
            BackStyle =0
            TextFontCharSet =178
            FontName ="Tahoma"
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
        Begin CheckBox
            SpecialEffect =2
            LabelX =230
            LabelY =-30
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
        Begin ComboBox
            SpecialEffect =2
            TextFontCharSet =178
            Width =1701
            LabelX =-1701
            FontName ="Tahoma"
        End
        Begin Section
            Height =8063
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =4081
                    Top =2381
                    ColumnWidth =810
                    ColumnOrder =0
                    Name ="CODE"
                    ControlSource ="CODE"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6913
                            Top =2381
                            Width =570
                            Height =240
                            Name ="Label25"
                            Caption ="كد كالا"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =4081
                    Top =3061
                    ColumnOrder =1
                    TabIndex =1
                    Name ="HES"
                    ControlSource ="HES"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =7065
                            Top =3060
                            Width =720
                            Height =240
                            Name ="Label27"
                            Caption ="كد ويزيتور"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =4081
                    Top =3401
                    ColumnWidth =630
                    ColumnOrder =2
                    TabIndex =2
                    Name ="MAH"
                    ControlSource ="MAH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =7018
                            Top =3401
                            Width =465
                            Height =240
                            Name ="Label28"
                            Caption ="ماه"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =4081
                    Top =3741
                    ColumnWidth =4065
                    ColumnOrder =3
                    TabIndex =3
                    Name ="kala"
                    ControlSource ="kala"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =7063
                            Top =3741
                            Width =420
                            Height =240
                            Name ="Label29"
                            Caption ="كالا"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =4081
                    Top =4081
                    ColumnOrder =9
                    TabIndex =4
                    Name ="MABL_K"
                    ControlSource ="MABL_K"
                    Format ="Standard"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6780
                            Top =4080
                            Width =840
                            Height =240
                            Name ="Label30"
                            Caption ="مبلغ فروش"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =4081
                    Top =4422
                    ColumnWidth =1140
                    ColumnOrder =12
                    TabIndex =5
                    Name ="MABMAR"
                    ControlSource ="MABMAR"
                    Format ="Standard"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6705
                            Top =4425
                            Width =1035
                            Height =240
                            Name ="Label31"
                            Caption ="مبلغ مرجوعي"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =4081
                    Top =5102
                    ColumnWidth =1215
                    ColumnOrder =13
                    TabIndex =6
                    Name ="MEGH_MAR"
                    ControlSource ="MEGH_MAR"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6510
                            Top =5100
                            Width =1110
                            Height =240
                            Name ="Label33"
                            Caption ="مقدار مرجوعي"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =4081
                    Top =5442
                    ColumnWidth =1155
                    ColumnOrder =7
                    TabIndex =7
                    Name ="MEGHk"
                    ControlSource ="MEGHk"
                    ConditionalFormat = Begin
                        0x0100000066000000010000000000000004000000000000000200000001000000 ,
                        0x00000000ffcc9900000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x300000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6855
                            Top =5445
                            Width =930
                            Height =240
                            Name ="Label34"
                            Caption ="فروخته شده"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =4081
                    Top =6122
                    ColumnOrder =15
                    TabIndex =8
                    Name ="VISITOR"
                    ControlSource ="VISITOR"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6733
                            Top =6122
                            Width =750
                            Height =240
                            Name ="Label36"
                            Caption ="ويزيتور"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =4081
                    Top =6803
                    ColumnWidth =1230
                    ColumnOrder =6
                    TabIndex =9
                    Name ="MEGHkGOL"
                    ControlSource ="MEGHkGOL"
                    ConditionalFormat = Begin
                        0x0100000066000000010000000000000004000000000000000200000001000000 ,
                        0x00000000ff99cc00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x300000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6463
                            Top =6803
                            Width =1020
                            Height =240
                            Name ="Label38"
                            Caption ="هدف"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =4081
                    Top =7143
                    ColumnWidth =1200
                    ColumnOrder =8
                    TabIndex =10
                    Name ="MANDMEGH"
                    ControlSource ="MANDMEGH"
                    ConditionalFormat = Begin
                        0x0100000066000000010000000000000005000000000000000200000001000000 ,
                        0x00000000ccffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x300000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6463
                            Top =7143
                            Width =1020
                            Height =240
                            Name ="Label39"
                            Caption ="مانده"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =4081
                    Top =7483
                    ColumnWidth =975
                    ColumnOrder =11
                    TabIndex =11
                    Name ="DARSADFR"
                    ControlSource ="DARSADFR"
                    Format ="Fixed"
                    ConditionalFormat = Begin
                        0x010000007c000000030000000000000005000000000000000300000001000000 ,
                        0x00000000ff000000000000000000000004000000070000000100000000000000 ,
                        0xffcc000000000000040000000a0000000d0000000100000000000000ccffcc00 ,
                        0x31003000000000003100300000003700300000003700300000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6538
                            Top =7483
                            Width =945
                            Height =240
                            Name ="Label40"
                            Caption ="درصد تحقق"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =4081
                    Top =7823
                    ColumnWidth =915
                    ColumnOrder =10
                    TabIndex =12
                    Name ="DAYMAND"
                    ControlSource ="DAYMAND"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6613
                            Top =7823
                            Width =870
                            Height =240
                            Name ="Label41"
                            Caption ="روز مانده"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "Visit_gol_dtl_REP.cls"
