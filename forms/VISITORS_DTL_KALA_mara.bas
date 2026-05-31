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
    ItemSuffix =38
    Left =270
    Top =600
    Right =6990
    Bottom =8280
    DatasheetForeColor =33554432
    DatasheetGridlinesColor =12632256
    MaxRecords =1000000
    RecSrcDt = Begin
        0xa13f33fb10d7e440
    End
    InputParameters ="@DT1 = FORMS![F_MENU_DATE_VISIT]![DT1],@DT2= FORMS![F_MENU_DATE_VISIT]![DT2],@VI"
        "SITOR= FORMS![F_MENU_DATE_VISIT]![CUST_NO]"
    RecordSource ="VISITOR_DTL_KALA_mara"
    Caption ="عملكرد ويزيتورها"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Tahoma"
    UnknownProp = {258 ,66 ,12 ,4 ,244 } Begin
        0x400044005400310020003d00200046004f0052004d00530021005b0046005f00 ,
        0x4d0045004e0055005f0044004100540045005f00560049005300490054005d00 ,
        0x21005b004400540031005d002c0040004400540032003d00200046004f005200 ,
        0x4d00530021005b0046005f004d0045004e0055005f0044004100540045005f00 ,
        0x560049005300490054005d0021005b004400540032005d002c00400056004900 ,
        0x5300490054004f0052003d00200046004f0052004d00530021005b0046005f00 ,
        0x4d0045004e0055005f0044004100540045005f00560049005300490054005d00 ,
        0x21005b0043005500530054005f004e004f005d00
    End
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
            Height =6930
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
                    ColumnWidth =1065
                    ColumnOrder =0
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
                    ColumnWidth =4725
                    ColumnOrder =3
                    TabIndex =1
                    BackColor =13434828
                    Name ="kala"
                    ControlSource ="kala"
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
                    Left =1934
                    Top =855
                    ColumnOrder =8
                    TabIndex =2
                    Name ="MABL_K"
                    ControlSource ="MABL_K"
                    Format ="Standard"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =855
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
                    ColumnWidth =1020
                    ColumnOrder =6
                    TabIndex =3
                    Name ="MEGHk"
                    ControlSource ="MEGHk"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =1140
                            Width =690
                            Height =240
                            Name ="Label7"
                            Caption ="مقدار فروش"
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
                    ColumnOrder =11
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
                    ColumnOrder =9
                    TabIndex =5
                    Name ="N_MOIN"
                    ControlSource ="N_MOIN"
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
                    ColumnOrder =10
                    TabIndex =6
                    Name ="IMBAA"
                    ControlSource ="IMBAA"
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
                    ColumnOrder =12
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
                    ColumnOrder =7
                    TabIndex =8
                    BackColor =255
                    Name ="GHABEL"
                    ControlSource ="GHABEL"
                    Format ="Standard"
                    ConditionalFormat = Begin
                        0x0100000066000000010000000000000004000000000000000200000001000000 ,
                        0x00000000ff000000000000000000000000000000000000000000000000000000 ,
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
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1814
                    Top =5329
                    Width =2490
                    ColumnWidth =2820
                    ColumnOrder =1
                    TabIndex =9
                    Name ="VISITOR"
                    ControlSource ="VISITOR"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Top =5329
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
                    Left =1814
                    Top =5669
                    Width =2490
                    ColumnWidth =5355
                    ColumnOrder =14
                    TabIndex =10
                    Name ="ADDRESS"
                    ControlSource ="ADDRESS"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Top =5669
                            Width =450
                            Height =240
                            Name ="Label21"
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
                    Left =1814
                    Top =6010
                    Width =2490
                    ColumnWidth =1950
                    ColumnOrder =15
                    TabIndex =11
                    Name ="TEL"
                    ControlSource ="TEL"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Top =6010
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
                    Left =1814
                    Top =6350
                    Width =2490
                    ColumnOrder =16
                    TabIndex =12
                    Name ="TOZIH"
                    ControlSource ="TOZIH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Top =6350
                            Width =540
                            Height =240
                            Name ="Label25"
                            Caption ="توضيح"
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
                    Left =1814
                    Top =6690
                    Width =2490
                    ColumnWidth =1905
                    ColumnOrder =13
                    TabIndex =13
                    Name ="MOBILE"
                    ControlSource ="MOBILE"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Top =6690
                            Width =630
                            Height =240
                            Name ="Label27"
                            Caption ="موبايل"
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
                    Left =906
                    Top =4195
                    ColumnWidth =1080
                    ColumnOrder =2
                    TabIndex =14
                    Name ="CODE"
                    ControlSource ="CODE"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3738
                            Top =4195
                            Width =570
                            Height =240
                            Name ="Label36"
                            Caption ="كد كالا"
                        End
                    End
                End
                Begin ComboBox
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
                    Left =1360
                    Top =3458
                    ColumnWidth =3285
                    TabIndex =15
                    Name ="MENUIT"
                    ControlSource ="MENUIT"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT CODE, NAMES FROM TCODE_MENUITEM"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4057
                            Top =3458
                            Width =705
                            Height =240
                            Name ="Label37"
                            Caption ="گروه"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "VISITORS_DTL_KALA_mara.cls"
