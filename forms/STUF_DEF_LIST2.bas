Version =20
VersionRequired =20
Begin Form
    AllowDeletions = NotDefault
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
    Width =6283
    RowHeight =300
    ItemSuffix =32
    Left =1245
    Top =2280
    Right =9585
    Bottom =7140
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x9827eaf751ffe340
    End
    RecordSource ="SELECT STUF_DEF.CODE, STUF_DEF.NAME, STUF_DEF.N_FANI, STUF_DEF.TOZIH, STUF_DEF.V"
        "AHED, STUF_DEF.B_SEF, STUF_DEF.N_SEF, STUF_DEF.MIN_M, STUF_DEF.MAX_M, STUF_DEF.R"
        "ADAH, STUF_DEF.KINDK, STUF_DEF.MABL_F, TCOD_STUFGROUP.NAMES, TCOD_VAHEDS.NAMES A"
        "S tvnames, CAST(STUF_DEF.CODE AS BIGINT) AS VCOD, PRICE_LIST.KH, PRICE_LIST.FR, "
        "MOGUDI_KOL_ANBARHA.MANDAH FROM STUF_DEF LEFT OUTER JOIN MOGUDI_KOL_ANBARHA ON ST"
        "UF_DEF.CODE = MOGUDI_KOL_ANBARHA.CODE LEFT OUTER JOIN TCOD_STUFGROUP ON STUF_DEF"
        ".RADAH = TCOD_STUFGROUP.CODE LEFT OUTER JOIN TCOD_VAHEDS ON STUF_DEF.VAHED = TCO"
        "D_VAHEDS.CODE LEFT OUTER JOIN PRICE_LIST ON STUF_DEF.CODE = PRICE_LIST.CODE ORDE"
        "R BY CAST(STUF_DEF.CODE AS BIGINT)"
    Caption ="مشخصات كالاها"
    DatasheetFontName ="Tahoma"
    OnLoad ="[Event Procedure]"
    OnError ="[Event Procedure]"
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
        Begin OptionButton
            SpecialEffect =2
            LabelX =230
            LabelY =-30
        End
        Begin CheckBox
            SpecialEffect =2
            LabelX =230
            LabelY =-30
        End
        Begin OptionGroup
            SpecialEffect =3
            Width =1701
            Height =1701
        End
        Begin BoundObjectFrame
            SpecialEffect =2
            Width =4536
            Height =2835
            LabelX =-1701
        End
        Begin TextBox
            SpecialEffect =2
            OldBorderStyle =0
            TextFontCharSet =178
            Width =1701
            LabelX =-1701
            FontName ="Tahoma"
        End
        Begin ListBox
            SpecialEffect =2
            TextFontCharSet =178
            Width =1701
            Height =1417
            LabelX =-1701
            FontName ="Tahoma"
        End
        Begin ComboBox
            SpecialEffect =2
            TextFontCharSet =178
            Width =1701
            LabelX =-1701
            FontName ="Tahoma"
        End
        Begin Section
            Height =4215
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3600
                    Top =56
                    Width =1605
                    Height =315
                    ColumnOrder =0
                    FontWeight =700
                    Name ="CODE"
                    ControlSource ="CODE"
                    FontName ="Arabic Transparent"
                    OnClick ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =5368
                            Top =63
                            Width =600
                            Height =300
                            Name ="Label1"
                            Caption ="كد كالا"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3690
                    Top =736
                    Width =1515
                    Height =315
                    ColumnWidth =1417
                    ColumnOrder =1
                    FontWeight =700
                    TabIndex =1
                    Name ="N_FANI"
                    ControlSource ="N_FANI"
                    FontName ="Arabic Transparent"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =5368
                            Top =736
                            Width =915
                            Height =300
                            Name ="Label5"
                            Caption ="شماره فني"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2820
                    Top =1076
                    Width =2385
                    Height =315
                    ColumnWidth =1417
                    ColumnOrder =9
                    FontWeight =700
                    TabIndex =2
                    Name ="TOZIH"
                    ControlSource ="TOZIH"
                    FontName ="Arabic Transparent"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =5368
                            Top =1076
                            Width =615
                            Height =300
                            Name ="Label7"
                            Caption ="توضيح"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Top =386
                    Width =1296
                    Height =315
                    ColumnWidth =1417
                    ColumnOrder =10
                    FontWeight =700
                    TabIndex =3
                    Name ="B_SEF"
                    ControlSource ="B_SEF"
                    FontName ="Arabic Transparent"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextFontCharSet =0
                            Left =1455
                            Top =390
                            Width =1080
                            Height =405
                            Name ="Label11"
                            Caption ="في خرده فروش"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =87
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1
                    Top =726
                    Width =1296
                    Height =315
                    ColumnWidth =1417
                    ColumnOrder =11
                    FontWeight =700
                    TabIndex =4
                    Name ="N_SEF"
                    ControlSource ="N_SEF"
                    FontName ="Arabic Transparent"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextFontCharSet =0
                            Left =1460
                            Top =726
                            Width =1080
                            Height =300
                            Name ="Label13"
                            Caption ="نقطه سفارش"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Top =1418
                    Width =1296
                    Height =315
                    ColumnWidth =1417
                    ColumnOrder =12
                    FontWeight =700
                    TabIndex =5
                    Name ="MIN_M"
                    ControlSource ="MIN_M"
                    FontName ="Arabic Transparent"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextFontCharSet =0
                            Left =1459
                            Top =1418
                            Width =1260
                            Height =300
                            Name ="Label15"
                            Caption ="حداقل موجودي"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1
                    Top =1066
                    Width =1296
                    Height =315
                    ColumnWidth =1417
                    ColumnOrder =13
                    FontWeight =700
                    TabIndex =6
                    Name ="MAX_M"
                    ControlSource ="MAX_M"
                    FontName ="Arabic Transparent"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =1460
                            Top =1066
                            Width =1320
                            Height =300
                            Name ="Label17"
                            Caption ="حد اكثر موجودي"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =247
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Top =1707
                    Width =1296
                    Height =315
                    ColumnWidth =1995
                    ColumnOrder =3
                    FontWeight =700
                    TabIndex =7
                    Name ="MABL_F"
                    ControlSource ="MABL_F"
                    Format ="Standard"
                    FontName ="Arabic Transparent"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextFontCharSet =0
                            Left =1455
                            Top =1710
                            Width =1200
                            Height =405
                            Name ="Label23"
                            Caption ="في عمده  فروش"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =87
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2820
                    Top =396
                    Width =2385
                    Height =315
                    ColumnWidth =3390
                    ColumnOrder =2
                    FontWeight =700
                    TabIndex =10
                    Name ="NAME"
                    ControlSource ="NAME"
                    FontName ="Arabic Transparent"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =5368
                            Top =396
                            Width =660
                            Height =300
                            Name ="Label3"
                            Caption ="نام كالا"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3339
                    Top =1428
                    Width =1866
                    Height =315
                    ColumnWidth =1417
                    ColumnOrder =7
                    FontWeight =700
                    TabIndex =8
                    Name ="VAHED"
                    ControlSource ="tvnames"
                    FontName ="Arabic Transparent"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =5368
                            Top =1428
                            Width =465
                            Height =300
                            Name ="Label9"
                            Caption ="واحد"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3343
                    Top =1768
                    Width =1866
                    Height =315
                    ColumnWidth =2325
                    ColumnOrder =8
                    FontWeight =700
                    TabIndex =9
                    Name ="RADAH"
                    ControlSource ="NAMES"
                    FontName ="Arabic Transparent"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =5372
                            Top =1768
                            Width =795
                            Height =300
                            Name ="Label19"
                            Caption ="گروه كالا"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Top =2211
                    Width =1296
                    Height =315
                    ColumnWidth =1650
                    ColumnOrder =5
                    FontWeight =700
                    TabIndex =11
                    Name ="KHAR"
                    ControlSource ="KH"
                    Format ="Standard"
                    FontName ="Arabic Transparent"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextFontCharSet =0
                            Left =1459
                            Top =2211
                            Width =825
                            Height =435
                            Name ="Label25"
                            Caption ="آخرين خريد ريال"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Top =2646
                    Width =1296
                    Height =315
                    ColumnOrder =6
                    FontWeight =700
                    TabIndex =12
                    Name ="FR"
                    ControlSource ="FR"
                    Format ="Standard"
                    FontName ="Arabic Transparent"
                    Begin
                        Begin Label
                            OverlapFlags =87
                            TextFontCharSet =0
                            Left =1455
                            Top =2640
                            Width =855
                            Height =405
                            Name ="Label27"
                            Caption ="آخرين فروش ريال"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3288
                    Top =3004
                    Width =1605
                    Height =315
                    ColumnOrder =14
                    FontWeight =700
                    TabIndex =13
                    Name ="VCOD"
                    ControlSource ="VCOD"
                    FontName ="Arabic Transparent"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =5056
                            Top =3011
                            Width =600
                            Height =300
                            Name ="Label29"
                            Caption ="كد كالا"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3288
                    Top =3439
                    Width =1605
                    Height =255
                    ColumnWidth =750
                    ColumnOrder =4
                    FontWeight =700
                    TabIndex =14
                    Name ="MANDAH"
                    ControlSource ="MANDAH"
                    OnDblClick ="[Event Procedure]"
                    FontName ="Arabic Transparent"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =5055
                            Top =3450
                            Width =645
                            Height =240
                            Name ="Label31"
                            Caption ="موجودي*"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "STUF_DEF_LIST2.cls"
