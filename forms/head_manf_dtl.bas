Version =20
VersionRequired =20
Begin Form
    AutoResize = NotDefault
    RecordSelectors = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    OrderByOn = NotDefault
    AllowEdits = NotDefault
    DefaultView =2
    AllowUpdating =4
    ViewsAllowed =2
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =4940
    RowHeight =315
    ItemSuffix =38
    Left =705
    Top =645
    Right =8160
    Bottom =4380
    DatasheetGridlinesColor =12632256
    MaxRecords =100000
    Filter ="((head_manf_dtl.cokal=\"2888\"))"
    OrderBy ="head_manf_dtl.SMABL DESC"
    RecSrcDt = Begin
        0xe98bf9887024e340
    End
    RecordSource ="head_manf_dtl"
    Caption ="head_manf_dtl"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Tahoma"
    OnError ="[Event Procedure]"
    AllowFormView =0
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
            Height =6596
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1927
                    Top =113
                    ColumnWidth =1095
                    ColumnOrder =0
                    Name ="FNUMB"
                    ControlSource ="FNUMB"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =113
                            Width =990
                            Height =240
                            Name ="Label1"
                            Caption ="شماره فرمول"
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
                    Left =1927
                    Top =453
                    Width =2490
                    ColumnOrder =1
                    TabIndex =1
                    Name ="cokal"
                    ControlSource ="cokal"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =453
                            Width =555
                            Height =240
                            Name ="Label3"
                            Caption ="كد  كالا"
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
                    Left =1927
                    Top =793
                    Width =2490
                    ColumnWidth =4290
                    ColumnOrder =2
                    TabIndex =2
                    Name ="kalname"
                    ControlSource ="kalname"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =793
                            Width =660
                            Height =240
                            Name ="Label5"
                            Caption ="نام كالا"
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
                    Left =1927
                    Top =1133
                    ColumnOrder =4
                    TabIndex =3
                    Name ="IMBIBE_MANF"
                    ControlSource ="IMBIBE_MANF"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =1133
                            Width =1245
                            Height =240
                            Name ="Label7"
                            Caption ="جذب هزينه تبديل"
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
                    Left =1927
                    Top =1474
                    ColumnOrder =5
                    TabIndex =4
                    Name ="IMBIBE_SAR"
                    ControlSource ="IMBIBE_SAR"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =1474
                            Width =1275
                            Height =240
                            Name ="Label9"
                            Caption ="جذب هزينه سربار"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =215
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1927
                    Top =1814
                    ColumnWidth =1965
                    ColumnOrder =6
                    TabIndex =5
                    Name ="GHEYMAT"
                    ControlSource ="GHEYMAT"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =113
                            Top =1814
                            Width =1860
                            Height =240
                            Name ="Label11"
                            Caption ="قيمت تمام شده استاندارد"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =3
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1927
                    Top =2154
                    ColumnWidth =1245
                    ColumnOrder =7
                    TabIndex =6
                    Name ="SA_HOUR"
                    ControlSource ="SA_HOUR"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =2154
                            Width =1140
                            Height =240
                            Name ="Label13"
                            Caption ="جذب ساعت كار"
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
                    Left =1927
                    Top =2494
                    ColumnWidth =1815
                    ColumnOrder =8
                    TabIndex =7
                    Name ="SA_NHOU"
                    ControlSource ="SA_NHOU"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =2494
                            Width =1710
                            Height =240
                            Name ="Label15"
                            Caption ="نرخ ساعت كار استاندارد"
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
                    Left =1927
                    Top =2834
                    Width =2490
                    ColumnOrder =3
                    TabIndex =8
                    Name ="vahedkal"
                    ControlSource ="vahedkal"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =2834
                            Width =630
                            Height =240
                            Name ="Label17"
                            Caption ="نام واحد"
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
                    Left =1927
                    Top =3174
                    Width =2490
                    TabIndex =9
                    Name ="CODE"
                    ControlSource ="CODE"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =3174
                            Width =555
                            Height =240
                            Name ="Label19"
                            Caption ="كد  كالا"
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
                    Left =1927
                    Top =3514
                    Width =2490
                    ColumnWidth =4545
                    TabIndex =10
                    Name ="mavname"
                    ControlSource ="mavname"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =3514
                            Width =765
                            Height =240
                            Name ="Label21"
                            Caption ="نام ماده"
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
                    Left =1927
                    Top =3855
                    Width =2490
                    TabIndex =11
                    Name ="vahedmav"
                    ControlSource ="vahedmav"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =3855
                            Width =630
                            Height =240
                            Name ="Label23"
                            Caption ="نام واحد"
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
                    Left =1927
                    Top =4195
                    Width =2490
                    ColumnWidth =2385
                    TabIndex =12
                    Name ="namanbar"
                    ControlSource ="namanbar"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =4195
                            Width =630
                            Height =240
                            Name ="Label25"
                            Caption ="نام انبار"
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
                    Left =1927
                    Top =4535
                    TabIndex =13
                    Name ="ANBAR"
                    ControlSource ="ANBAR"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =4535
                            Width =330
                            Height =240
                            Name ="Label27"
                            Caption ="انبار"
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
                    Left =1927
                    Top =4875
                    TabIndex =14
                    Name ="MEGH"
                    ControlSource ="MEGH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =4875
                            Width =780
                            Height =240
                            Name ="Label29"
                            Caption ="مقدار مواد"
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
                    Left =1927
                    Top =5215
                    TabIndex =15
                    Name ="MEGHk"
                    ControlSource ="MEGHk"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =5215
                            Width =1020
                            Height =240
                            Name ="Label31"
                            Caption ="مقدار كل مواد"
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
                    Left =1927
                    Top =5555
                    TabIndex =16
                    Name ="PERT"
                    ControlSource ="PERT"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =5555
                            Width =330
                            Height =240
                            Name ="Label33"
                            Caption ="پرت"
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
                    Left =1927
                    Top =5896
                    TabIndex =17
                    Name ="SMABL"
                    ControlSource ="SMABL"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =5896
                            Width =1050
                            Height =240
                            Name ="Label35"
                            Caption ="مبلغ استاندارد"
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
                    Left =1927
                    Top =6236
                    TabIndex =18
                    Name ="MABLK"
                    ControlSource ="MABLK"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =6236
                            Width =615
                            Height =240
                            Name ="Label37"
                            Caption ="مبلغ كل"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "head_manf_dtl.cls"
