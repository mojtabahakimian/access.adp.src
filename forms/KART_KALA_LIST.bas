Version =20
VersionRequired =20
Begin Form
    DefaultView =0
    AllowUpdating =4
    TabularCharSet =178
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =5669
    RowHeight =330
    ItemSuffix =20
    Left =600
    Top =150
    Right =28050
    Bottom =12330
    DatasheetForeColor =33554432
    DatasheetGridlinesColor =12632256
    MaxRecords =1000000
    RecSrcDt = Begin
        0xc9d5b5739874e640
    End
    RecordSource ="exec dbo.usp_KART_ANBAR_TOTAL_CLEAN_V3 2000"
    Caption ="ليست كارت انبار كالا"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Tahoma"
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
            Height =7156
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
                    Left =2834
                    Top =1133
                    ColumnWidth =645
                    ColumnOrder =14
                    Name ="ANBAR"
                    ControlSource ="ANBAR"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1133
                            Top =1133
                            Width =600
                            Height =240
                            Name ="Label0"
                            Caption ="كد انبار"
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
                    Left =2834
                    Top =1474
                    ColumnWidth =855
                    ColumnOrder =10
                    TabIndex =1
                    Name ="CODE"
                    ControlSource ="CODE"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1133
                            Top =1474
                            Width =570
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
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2834
                    Top =1814
                    ColumnWidth =1020
                    ColumnOrder =3
                    TabIndex =2
                    Name ="DATE_N"
                    ControlSource ="DATE_N"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1133
                            Top =1814
                            Width =720
                            Height =240
                            Name ="Label2"
                            Caption ="تاريخ"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2834
                    Top =2154
                    ColumnWidth =1050
                    ColumnOrder =4
                    TabIndex =3
                    BackColor =13434828
                    Name ="runingmeg"
                    ControlSource ="Running_MEG"
                    Format ="Fixed"
                    ConditionalFormat = Begin
                        0x010000006c000000020000000000000005000000000000000200000001000000 ,
                        0x00000000ff000000000000000600000003000000050000000100000000000000 ,
                        0xccffcc0000000000000000000000000000000000000000000000000000000000 ,
                        0x300000000000300000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1140
                            Top =2160
                            Width =645
                            Height =240
                            Name ="Label3"
                            Caption ="موجودي"
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
                    Left =2834
                    Top =3514
                    ColumnWidth =870
                    ColumnOrder =7
                    TabIndex =4
                    Name ="MEGK"
                    ControlSource ="MEGK"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1133
                            Top =3514
                            Width =1065
                            Height =240
                            Name ="Label7"
                            Caption ="مقدار"
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
                    Left =2834
                    Top =3855
                    ColumnWidth =1320
                    ColumnOrder =1
                    TabIndex =5
                    Name ="BARGAH"
                    ControlSource ="BARGAH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1133
                            Top =3855
                            Width =750
                            Height =240
                            Name ="Label8"
                            Caption ="نوع برگه"
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
                    Left =2834
                    Top =4195
                    ColumnWidth =3210
                    ColumnOrder =11
                    TabIndex =6
                    Name ="NAME"
                    ControlSource ="NAME"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1133
                            Top =4195
                            Width =555
                            Height =240
                            Name ="Label9"
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
                    Left =2834
                    Top =4875
                    ColumnWidth =900
                    ColumnOrder =0
                    TabIndex =7
                    Name ="NUMBER"
                    ControlSource ="NUMBER"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1140
                            Top =4875
                            Width =855
                            Height =240
                            Name ="Label11"
                            Caption ="شماره برگه"
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
                    Left =2834
                    Top =6576
                    ColumnOrder =12
                    TabIndex =8
                    Name ="FNUMCO"
                    ControlSource ="FNUMCO"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1140
                            Top =6570
                            Width =1020
                            Height =240
                            Name ="Label16"
                            Caption ="شماره داخلي"
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
                    Left =2834
                    Top =6916
                    ColumnOrder =13
                    TabIndex =9
                    Name ="N_FANI"
                    ControlSource ="N_FANI"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1140
                            Top =6915
                            Width =870
                            Height =240
                            Name ="Label17"
                            Caption ="شماره فني"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "KART_KALA_LIST.cls"
