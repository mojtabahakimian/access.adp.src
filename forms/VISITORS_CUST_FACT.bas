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
    Width =4940
    RowHeight =360
    ItemSuffix =37
    Left =345
    Top =7155
    Right =5325
    Bottom =8595
    DatasheetForeColor =33554432
    DatasheetGridlinesColor =12632256
    MaxRecords =1000000
    Tag ="@DT1 = FORMS![F_MENU_DATE]![DT1],@DT2= FORMS![F_MENU_DATE]![DT2],@VIS= FORMS![VI"
        "SITORS_KOL]![CUST_NO]"
    RecSrcDt = Begin
        0xc2bed8cd6c47e440
    End
    Caption ="عملكرد ويزيتورها به تفكيك فاكتور ها"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Tahoma"
    AllowFormView =0
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
            Height =6929
            BackColor =-2147483633
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
                    Left =1927
                    Top =113
                    Width =2490
                    ColumnWidth =825
                    ColumnOrder =14
                    Name ="CUST_NO"
                    ControlSource ="CUST_NO"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =113
                            Width =780
                            Height =240
                            Name ="Label1"
                            Caption ="كد ويزيتور"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1927
                    Top =453
                    ColumnOrder =2
                    TabIndex =1
                    BackColor =13434828
                    Name ="mabpur"
                    ControlSource ="mabpur"
                    Format ="Standard"
                    ConditionalFormat = Begin
                        0x0100000066000000010000000000000004000000000000000200000001000000 ,
                        0x00000000ccffcc00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x300000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =453
                            Width =615
                            Height =240
                            Name ="Label3"
                            Caption ="جمع پورسانت"
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
                    Left =1927
                    Top =793
                    ColumnOrder =5
                    TabIndex =2
                    Name ="MABL_K"
                    ControlSource ="MABL_K"
                    Format ="Standard"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =793
                            Width =645
                            Height =240
                            Name ="Label5"
                            Caption ="مبلغ كل فروش"
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
                    Left =1927
                    Top =1133
                    ColumnWidth =705
                    ColumnOrder =4
                    TabIndex =3
                    Name ="MEGHk"
                    ControlSource ="MEGHk"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =1133
                            Width =570
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
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1927
                    Top =1474
                    ColumnWidth =795
                    ColumnOrder =8
                    TabIndex =4
                    Name ="MEGH_MAR"
                    ControlSource ="MEGH_MAR"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =1474
                            Width =915
                            Height =240
                            Name ="Label9"
                            Caption ="مرجوعي"
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
                    Left =1927
                    Top =1814
                    ColumnWidth =975
                    ColumnOrder =6
                    TabIndex =5
                    Name ="TAKHF"
                    ControlSource ="TAKHF"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =1814
                            Width =555
                            Height =240
                            Name ="Label11"
                            Caption ="جمع تخفيف"
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
                    Left =1927
                    Top =2154
                    ColumnWidth =1005
                    ColumnOrder =7
                    TabIndex =6
                    Name ="MBAA"
                    ControlSource ="MBAA"
                    Format ="Standard"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =2154
                            Width =495
                            Height =240
                            Name ="Label13"
                            Caption ="ارزش افزوده"
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
                    Left =1927
                    Top =2494
                    ColumnWidth =1140
                    ColumnOrder =9
                    TabIndex =7
                    Name ="MABMAR"
                    ControlSource ="MABMAR"
                    Format ="Standard"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =2494
                            Width =720
                            Height =240
                            Name ="Label15"
                            Caption ="مبلغ مرجوعي"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1927
                    Top =2834
                    ColumnOrder =10
                    TabIndex =8
                    BackColor =10092543
                    Name ="GHABEL"
                    ControlSource ="GHABEL"
                    Format ="Standard"
                    ConditionalFormat = Begin
                        0x0100000066000000010000000000000004000000000000000200000001000000 ,
                        0x00000000ffff9900000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x300000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =2834
                            Width =645
                            Height =240
                            Name ="Label17"
                            Caption ="فروش خالص"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2040
                    Top =4648
                    Width =2490
                    ColumnWidth =2445
                    ColumnOrder =15
                    TabIndex =9
                    Name ="NAME"
                    ControlSource ="VISNAME"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =226
                            Top =4648
                            Width =1185
                            Height =240
                            Name ="Label19"
                            Caption ="نام ويزيتور"
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
                    Left =2040
                    Top =4988
                    Width =2490
                    ColumnWidth =3195
                    ColumnOrder =17
                    TabIndex =10
                    Name ="CUSTNAME"
                    ControlSource ="CUSTNAME"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =225
                            Top =4995
                            Width =630
                            Height =240
                            Name ="Label21"
                            Caption ="مشتري"
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
                    Left =2040
                    Top =5329
                    Width =2490
                    ColumnWidth =2055
                    ColumnOrder =12
                    TabIndex =11
                    Name ="TEL"
                    ControlSource ="TEL"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =226
                            Top =5329
                            Width =390
                            Height =240
                            Name ="Label23"
                            Caption ="تلفن"
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
                    Left =2040
                    Top =5669
                    Width =2490
                    ColumnWidth =3105
                    ColumnOrder =13
                    TabIndex =12
                    Name ="ADDRESS"
                    ControlSource ="ADDRESS"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =226
                            Top =5669
                            Width =540
                            Height =240
                            Name ="Label25"
                            Caption ="آدرس"
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
                    Left =2040
                    Top =6009
                    Width =2490
                    ColumnWidth =1905
                    ColumnOrder =11
                    TabIndex =13
                    Name ="MOBILE"
                    ControlSource ="MOBILE"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =226
                            Top =6009
                            Width =630
                            Height =240
                            Name ="Label27"
                            Caption ="موبايل"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1927
                    Top =3194
                    ColumnWidth =540
                    ColumnOrder =3
                    TabIndex =14
                    BackColor =16751052
                    Name ="DARSAD"
                    ControlSource ="=IIf(([GHABEL])<>0,([mabpur])/([GHABEL])*100,0)"
                    Format ="Standard"
                    ConditionalFormat = Begin
                        0x0100000066000000010000000000000004000000000000000200000001000000 ,
                        0x00000000cc99ff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x300000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =3194
                            Width =645
                            Height =240
                            Name ="Label29"
                            Caption ="درصد"
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
                    Left =1829
                    Top =3630
                    Width =2490
                    ColumnWidth =1020
                    ColumnOrder =16
                    TabIndex =15
                    Name ="CUSTOMER"
                    ControlSource ="CUSTOMER"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =15
                            Top =3630
                            Width =840
                            Height =240
                            Name ="Label31"
                            Caption ="كد مشتري"
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
                    Left =1934
                    Top =4095
                    ColumnWidth =1065
                    ColumnOrder =1
                    TabIndex =16
                    Name ="NUMBER"
                    ControlSource ="NUMBER"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =4095
                            Width =960
                            Height =240
                            Name ="Label33"
                            Caption ="شماره فاكتور"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =247
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1934
                    Top =4455
                    ColumnOrder =0
                    TabIndex =17
                    Name ="DATE_N"
                    ControlSource ="DATE_N"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =120
                            Top =4455
                            Width =960
                            Height =240
                            Name ="Label35"
                            Caption ="تاريخ فاكتور"
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
                    Left =793
                    Top =6349
                    TabIndex =18
                    Name ="TOZIH"
                    ControlSource ="TOZIH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3595
                            Top =6349
                            Width =600
                            Height =240
                            Name ="Label36"
                            Caption ="توضيح"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "VISITORS_CUST_FACT.cls"
