Version =20
VersionRequired =20
Begin Form
    AllowDeletions = NotDefault
    AllowAdditions = NotDefault
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
    ItemSuffix =42
    Left =2160
    Top =2670
    Right =10560
    Bottom =7530
    DatasheetGridlinesColor =12632256
    MaxRecords =1000000
    OrderBy ="tvnames"
    RecSrcDt = Begin
        0xaeb8fd80bf5ae340
    End
    RecordSource ="SELECT STUF_DEF.CODE, STUF_DEF.NAME, STUF_DEF.N_FANI, STUF_DEF.TOZIH, STUF_DEF.V"
        "AHED, STUF_DEF.B_SEF, STUF_DEF.N_SEF, STUF_DEF.MIN_M, STUF_DEF.MAX_M, STUF_DEF.R"
        "ADAH, STUF_DEF.KINDK, STUF_DEF.MABL_F, TCOD_STUFGROUP.NAMES, TCOD_VAHEDS.NAMES A"
        "S tvnames, CAST(STUF_DEF.CODE AS BIGINT) AS VCOD, MOGUDI_KOL_ANBARHA.MANDAH, MOG"
        "UDI_KOL_ANBARHA.MANDAH - STUF_DEF.B_SEF AS BSEF, MOGUDI_KOL_ANBARHA.MANDAH - STU"
        "F_DEF.N_SEF AS NSEF, MOGUDI_KOL_ANBARHA.MANDAH - STUF_DEF.MIN_M AS MINM, MOGUDI_"
        "KOL_ANBARHA.MANDAH - STUF_DEF.MAX_M AS MAXM FROM STUF_DEF LEFT OUTER JOIN MOGUDI"
        "_KOL_ANBARHA ON STUF_DEF.CODE = MOGUDI_KOL_ANBARHA.CODE LEFT OUTER JOIN TCOD_STU"
        "FGROUP ON STUF_DEF.RADAH = TCOD_STUFGROUP.CODE LEFT OUTER JOIN TCOD_VAHEDS ON ST"
        "UF_DEF.VAHED = TCOD_VAHEDS.CODE ORDER BY CAST(STUF_DEF.CODE AS BIGINT)"
    Caption ="مشخصات كالاها"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Tahoma"
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
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3600
                    Top =56
                    Width =1605
                    Height =315
                    ColumnWidth =705
                    ColumnOrder =16
                    FontWeight =700
                    Name ="CODE"
                    ControlSource ="CODE"
                    OnDblClick ="[Event Procedure]"
                    FontName ="Arabic Transparent"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =5368
                            Top =63
                            Width =600
                            Height =300
                            Name ="Label1"
                            Caption ="كد كالا*"
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
                    ColumnWidth =2850
                    ColumnOrder =7
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
                    ColumnWidth =615
                    ColumnOrder =8
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
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Top =113
                    Width =1296
                    Height =315
                    ColumnWidth =1260
                    ColumnOrder =3
                    FontWeight =700
                    TabIndex =3
                    Name ="B_SEF"
                    ControlSource ="B_SEF"
                    FontName ="Arabic Transparent"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =1455
                            Top =117
                            Width =1080
                            Height =405
                            Name ="Label11"
                            Caption ="في خرده فروش"
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
                    Top =623
                    Width =1296
                    Height =315
                    ColumnWidth =1215
                    ColumnOrder =10
                    FontWeight =700
                    TabIndex =4
                    Name ="N_SEF"
                    ControlSource ="N_SEF"
                    OnDblClick ="[Event Procedure]"
                    FontName ="Arabic Transparent"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextFontCharSet =0
                            Left =1459
                            Top =623
                            Width =1380
                            Height =300
                            Name ="Label13"
                            Caption ="نقطه سفارش*"
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
                    ColumnWidth =1230
                    ColumnOrder =11
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
                    ColumnWidth =1470
                    ColumnOrder =9
                    FontWeight =700
                    TabIndex =6
                    Name ="MAX_M"
                    ControlSource ="MAX_M"
                    FontName ="Arabic Transparent"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextFontCharSet =0
                            Left =1455
                            Top =1065
                            Width =1320
                            Height =405
                            Name ="Label17"
                            Caption ="قيمت مصرف كننده"
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
                    ColumnWidth =1485
                    ColumnOrder =4
                    FontWeight =700
                    TabIndex =7
                    Name ="MABL_F"
                    ControlSource ="MABL_F"
                    Format ="Standard"
                    OnDblClick ="[Event Procedure]"
                    FontName ="Arabic Transparent"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextFontCharSet =0
                            Left =1459
                            Top =1707
                            Width =1440
                            Height =300
                            Name ="Label23"
                            Caption ="في عمده  فروش *"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =215
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2820
                    Top =396
                    Width =2385
                    Height =315
                    ColumnWidth =4290
                    ColumnOrder =1
                    FontWeight =700
                    TabIndex =10
                    Name ="NAME"
                    ControlSource ="NAME"
                    OnDblClick ="[Event Procedure]"
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
                            Caption ="نام كالا*"
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
                    ColumnWidth =750
                    ColumnOrder =5
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
                    ColumnWidth =1020
                    ColumnOrder =6
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
                    ColumnWidth =765
                    ColumnOrder =0
                    FontWeight =700
                    TabIndex =11
                    Name ="VCOD"
                    ControlSource ="VCOD"
                    OnDblClick ="[Event Procedure]"
                    FontName ="Arabic Transparent"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =5055
                            Top =3015
                            Width =600
                            Height =405
                            Name ="Label29"
                            Caption ="كد كالا.*"
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
                    Top =2142
                    Width =1296
                    Height =315
                    ColumnWidth =1710
                    ColumnOrder =12
                    FontWeight =700
                    TabIndex =12
                    Name ="BSEF"
                    ControlSource ="BSEF"
                    FontName ="Arabic Transparent"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextFontCharSet =0
                            Left =1455
                            Top =2145
                            Width =2460
                            Height =405
                            Name ="Label31"
                            Caption ="فاصله تا بهينه سفارش"
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
                    Top =2482
                    Width =1296
                    Height =315
                    ColumnWidth =1710
                    ColumnOrder =13
                    FontWeight =700
                    TabIndex =13
                    Name ="NSEF"
                    ControlSource ="NSEF"
                    FontName ="Arabic Transparent"
                    Begin
                        Begin Label
                            OverlapFlags =223
                            TextFontCharSet =0
                            Left =1455
                            Top =2475
                            Width =1080
                            Height =405
                            Name ="Label33"
                            Caption ="فاصله تا نقطه سفارش"
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
                    Top =3231
                    Width =1296
                    Height =315
                    ColumnWidth =1815
                    ColumnOrder =14
                    FontWeight =700
                    TabIndex =14
                    Name ="MINM"
                    ControlSource ="MINM"
                    FontName ="Arabic Transparent"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextFontCharSet =0
                            Left =1455
                            Top =3225
                            Width =1260
                            Height =405
                            Name ="Label35"
                            Caption ="فاصله تا حداقل موجودي"
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
                    Top =2822
                    Width =1296
                    Height =315
                    ColumnWidth =1875
                    ColumnOrder =15
                    FontWeight =700
                    TabIndex =15
                    Name ="MAXM"
                    ControlSource ="MAXM"
                    FontName ="Arabic Transparent"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextFontCharSet =0
                            Left =1455
                            Top =2820
                            Width =1320
                            Height =405
                            Name ="Label37"
                            Caption ="فاصله تا حد اكثر موجودي"
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
                    Top =3756
                    Width =1296
                    Height =315
                    ColumnWidth =750
                    ColumnOrder =2
                    FontWeight =700
                    TabIndex =16
                    Name ="MANDAH"
                    ControlSource ="MANDAH"
                    FontName ="Arabic Transparent"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =1455
                            Top =3750
                            Width =1260
                            Height =405
                            Name ="Label41"
                            Caption ="موجودي"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "STUF_DEF_LIST.cls"
