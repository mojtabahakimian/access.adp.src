Version =20
VersionRequired =20
Begin Form
    OrderByOn = NotDefault
    DefaultView =2
    AllowUpdating =4
    ViewsAllowed =2
    TabularCharSet =178
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =5564
    RowHeight =360
    ItemSuffix =42
    Left =510
    Top =615
    Right =18630
    Bottom =9675
    DatasheetForeColor =33554432
    DatasheetGridlinesColor =12632256
    MaxRecords =1000000
    Tag ="@DT1 = FORMS![F_MENU_DATE]![DT1],@DT2= FORMS![F_MENU_DATE]![DT2],@VIS= FORMS![VI"
        "SITORS_KOL]![CUST_NO]"
    RecSrcDt = Begin
        0xa26275f11247e440
    End
    Caption ="عملكرد ويزيتورها به تفكيك مشتري"
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
        Begin ComboBox
            SpecialEffect =2
            Width =1701
            LabelX =-1701
            FontName ="Tahoma"
        End
        Begin Subform
            SpecialEffect =2
            Width =1701
            Height =1701
        End
        Begin Section
            CanGrow = NotDefault
            Height =9694
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
                    Left =2551
                    Top =113
                    Width =1710
                    ColumnWidth =1020
                    ColumnOrder =15
                    Name ="CUST_NO"
                    ControlSource ="CUST_NO"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =737
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
                    Left =2549
                    Top =390
                    ColumnOrder =3
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
                            Left =735
                            Top =390
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
                    Left =2549
                    Top =675
                    ColumnOrder =6
                    TabIndex =2
                    Name ="MABL_K"
                    ControlSource ="MABL_K"
                    Format ="Standard"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =735
                            Top =675
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
                    Left =2549
                    Top =945
                    ColumnWidth =705
                    ColumnOrder =5
                    TabIndex =3
                    Name ="MEGHk"
                    ControlSource ="MEGHk"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =735
                            Top =945
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
                    Left =2549
                    Top =1230
                    ColumnWidth =795
                    ColumnOrder =9
                    TabIndex =4
                    Name ="MEGH_MAR"
                    ControlSource ="MEGH_MAR"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =735
                            Top =1230
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
                    Left =2549
                    Top =1515
                    ColumnWidth =975
                    ColumnOrder =7
                    TabIndex =5
                    Name ="TAKHF"
                    ControlSource ="TAKHF"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =735
                            Top =1515
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
                    Left =2549
                    Top =1800
                    ColumnWidth =1005
                    ColumnOrder =8
                    TabIndex =6
                    Name ="MBAA"
                    ControlSource ="MBAA"
                    Format ="Standard"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =735
                            Top =1800
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
                    Left =2549
                    Top =2085
                    ColumnWidth =1140
                    ColumnOrder =10
                    TabIndex =7
                    Name ="MABMAR"
                    ControlSource ="MABMAR"
                    Format ="Standard"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =735
                            Top =2085
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
                    Left =2549
                    Top =2370
                    ColumnOrder =11
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
                            Left =735
                            Top =2370
                            Width =645
                            Height =240
                            Name ="Label17"
                            Caption ="فروش خالص"
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
                    Left =2664
                    Top =4648
                    Width =2490
                    ColumnWidth =2445
                    ColumnOrder =16
                    TabIndex =9
                    Name ="NAME"
                    ControlSource ="VISNAME"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =850
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
                    Left =2664
                    Top =4988
                    Width =2490
                    ColumnWidth =3150
                    ColumnOrder =1
                    TabIndex =10
                    Name ="CUSTNAME"
                    ControlSource ="CUSTNAME"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =849
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
                    Left =2664
                    Top =5329
                    Width =2490
                    ColumnWidth =2055
                    ColumnOrder =13
                    TabIndex =11
                    Name ="TEL"
                    ControlSource ="TEL"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =850
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
                    Left =2664
                    Top =5669
                    Width =2490
                    ColumnWidth =6525
                    ColumnOrder =14
                    TabIndex =12
                    Name ="ADDRESS"
                    ControlSource ="ADDRESS"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =850
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
                    Left =2664
                    Top =6009
                    Width =2490
                    ColumnWidth =1905
                    ColumnOrder =12
                    TabIndex =13
                    Name ="MOBILE"
                    ControlSource ="MOBILE"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =850
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
                    Left =2549
                    Top =2655
                    ColumnWidth =540
                    ColumnOrder =4
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
                            Left =735
                            Top =2655
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
                    Left =2549
                    Top =2940
                    Width =1710
                    ColumnWidth =1020
                    ColumnOrder =0
                    TabIndex =15
                    Name ="CUSTOMER"
                    ControlSource ="CUSTOMER"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =735
                            Top =2940
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
                    Left =2549
                    Top =3240
                    ColumnWidth =855
                    ColumnOrder =2
                    TabIndex =16
                    Name ="TDF"
                    ControlSource ="TDF"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =735
                            Top =3240
                            Width =915
                            Height =240
                            Name ="Label33"
                            Caption ="تعداد فاكتور"
                        End
                    End
                End
                Begin Subform
                    OverlapFlags =87
                    Top =6519
                    Width =5271
                    TabIndex =17
                    Name ="VISITORS_CUST_FACT"
                    SourceObject ="Form.VISITORS_CUST_FACT"
                    LinkChildFields ="CUSTOMER;CUST_NO"
                    LinkMasterFields ="CUSTOMER;CUST_NO"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =1041
                            Top =6279
                            Width =1740
                            Height =240
                            Name ="VISITORS_CUST_FACT Label"
                            Caption ="VISITORS_CUST_FACT"
                            EventProcPrefix ="VISITORS_CUST_FACT_Label"
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
                    Left =2669
                    Top =4260
                    Width =2490
                    ColumnWidth =3960
                    TabIndex =18
                    Name ="ROUTE_NAME"
                    ControlSource ="ROUTE_NAME"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =855
                            Top =4260
                            Width =1185
                            Height =240
                            Name ="Label37"
                            Caption ="نام مسير"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =178
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =963
                    Top =8617
                    Width =2665
                    Height =300
                    ColumnWidth =1530
                    TabIndex =19
                    Name ="OSTANID"
                    ControlSource ="OSTANID"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT OSCODE, OSNAME FROM TCOD_OSTAN ORDER BY OSNAME"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3786
                            Top =8622
                            Width =705
                            Height =240
                            Name ="Label38"
                            Caption ="استان"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =178
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =963
                    Top =9018
                    Width =2665
                    Height =300
                    TabIndex =20
                    Name ="SHAHRID"
                    ControlSource ="SHAHRID"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT CITYCODE, CITYNAME FROM TCOD_CITY ORDER BY CITYNAME"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3764
                            Top =9012
                            Width =990
                            Height =240
                            Name ="Label40"
                            Caption ="شهرستان"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "VISITORS_CUST_F.cls"
