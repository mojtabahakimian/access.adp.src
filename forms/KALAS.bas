Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AllowDeletions = NotDefault
    AllowAdditions = NotDefault
    OrderByOn = NotDefault
    AllowEdits = NotDefault
    ServerFilterByForm = NotDefault
    DefaultView =2
    AllowUpdating =4
    ViewsAllowed =2
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =6688
    RowHeight =300
    ItemSuffix =49
    Left =1050
    Top =1110
    Right =8535
    Bottom =4950
    DatasheetGridlinesColor =12632256
    MaxRecords =100000000
    ShortcutMenuBar ="SERVERFILTER"
    RecSrcDt = Begin
        0x89627c0449eee240
    End
    RecordSource ="KALAS"
    Caption ="ليست گردشها ي كالا"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Tahoma"
    OnLoad ="[Event Procedure]"
    OnError ="[Event Procedure]"
    RecordSourceQualifier ="dbo"
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
            Height =13776
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1134
                    Top =113
                    Width =2268
                    ColumnWidth =1125
                    ColumnOrder =4
                    Name ="NUMBER"
                    ControlSource ="NUMBER"
                    Format ="Fixed"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3969
                            Top =113
                            Width =960
                            Height =240
                            Name ="Label1"
                            Caption ="شماره فاكتور"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1134
                    Top =453
                    Width =2268
                    ColumnWidth =1410
                    ColumnOrder =1
                    TabIndex =1
                    Name ="BARGAH"
                    ControlSource ="BARGAH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3969
                            Top =450
                            Width =735
                            Height =240
                            Name ="Label3"
                            Caption ="نوع فاكتور"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1134
                    Top =793
                    Width =2268
                    ColumnWidth =1380
                    ColumnOrder =11
                    TabIndex =2
                    Name ="ANBNAME"
                    ControlSource ="ANBNAME"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3969
                            Top =793
                            Width =570
                            Height =240
                            Name ="Label5"
                            Caption ="نام انبار"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1134
                    Top =1133
                    Width =2268
                    ColumnWidth =1650
                    ColumnOrder =12
                    TabIndex =3
                    Name ="NUMBER1"
                    ControlSource ="NUMBER1"
                    Format ="Fixed"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3969
                            Top =1133
                            Width =1530
                            Height =240
                            Name ="Label7"
                            Caption ="شماره فاكتور برگشت"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1134
                    Top =1474
                    Width =2268
                    ColumnWidth =1080
                    ColumnOrder =0
                    TabIndex =4
                    Name ="DATE_N"
                    ControlSource ="DATE_N"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =3969
                            Top =1474
                            Width =825
                            Height =240
                            Name ="Label9"
                            Caption ="تاريخ فاكتور"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1134
                    Top =2010
                    Width =2268
                    ColumnWidth =885
                    ColumnOrder =13
                    TabIndex =5
                    Name ="N_S"
                    ControlSource ="N_S"
                    Format ="Fixed"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3969
                            Top =2010
                            Width =885
                            Height =240
                            Name ="Label11"
                            Caption ="شماره سند"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =247
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1077
                    Top =1700
                    Width =2268
                    ColumnWidth =3210
                    ColumnOrder =3
                    TabIndex =6
                    Name ="CUSTNAME"
                    ControlSource ="CUSTNAME"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =3912
                            Top =1700
                            Width =1185
                            Height =240
                            Name ="Label13"
                            Caption ="نام شخص"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1134
                    Top =2577
                    Width =2268
                    ColumnOrder =14
                    TabIndex =7
                    Name ="MOLAH"
                    ControlSource ="MOLAH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3969
                            Top =2577
                            Width =690
                            Height =240
                            Name ="Label15"
                            Caption ="توضيحات"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1247
                    Top =2891
                    Width =2268
                    ColumnWidth =1770
                    ColumnOrder =15
                    TabIndex =8
                    Name ="FNUMCO"
                    ControlSource ="FNUMCO"
                    Format ="Fixed"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4082
                            Top =2891
                            Width =1620
                            Height =240
                            Name ="Label19"
                            Caption ="شماره فاكتور فروشنده"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1134
                    Top =3514
                    Width =2268
                    ColumnWidth =825
                    ColumnOrder =7
                    TabIndex =9
                    Name ="MEGH"
                    ControlSource ="MEGH"
                    Format ="Fixed"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3969
                            Top =3514
                            Width =675
                            Height =240
                            Name ="Label21"
                            Caption ="مقداركالا"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1134
                    Top =3855
                    Width =2268
                    ColumnWidth =1140
                    ColumnOrder =8
                    TabIndex =10
                    Name ="MEGHk"
                    ControlSource ="MEGHk"
                    Format ="Fixed"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3969
                            Top =3855
                            Width =960
                            Height =240
                            Name ="Label23"
                            Caption ="مقدار كل كالا"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1134
                    Top =4195
                    Width =2268
                    ColumnWidth =1245
                    ColumnOrder =16
                    TabIndex =11
                    Name ="MEGH_MAR"
                    ControlSource ="MEGH_MAR"
                    Format ="Fixed"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3969
                            Top =4195
                            Width =1110
                            Height =240
                            Name ="Label25"
                            Caption ="مقدار مرجوعي"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1134
                    Top =4535
                    Width =2268
                    ColumnOrder =9
                    TabIndex =12
                    Name ="MABL"
                    ControlSource ="MABL"
                    Format ="#,###"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3969
                            Top =4535
                            Width =375
                            Height =240
                            Name ="Label27"
                            Caption ="مبلغ"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1134
                    Top =4875
                    Width =2268
                    ColumnWidth =4605
                    ColumnOrder =6
                    TabIndex =13
                    Name ="KALA"
                    ControlSource ="KALA"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3969
                            Top =4875
                            Width =1305
                            Height =240
                            Name ="Label29"
                            Caption ="نام كالا"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1134
                    Top =5215
                    Width =2268
                    ColumnOrder =10
                    TabIndex =14
                    Name ="MABL_K"
                    ControlSource ="MABL_K"
                    Format ="#,###"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3969
                            Top =5215
                            Width =615
                            Height =240
                            Name ="Label31"
                            Caption ="مبلغ كل"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1134
                    Top =6236
                    Width =2268
                    ColumnOrder =17
                    TabIndex =15
                    Name ="VAHEDNAME"
                    ControlSource ="VAHEDNAME"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3969
                            Top =6236
                            Width =630
                            Height =240
                            Name ="Label37"
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
                    Left =1134
                    Top =6576
                    Width =2268
                    ColumnWidth =1695
                    ColumnOrder =18
                    TabIndex =16
                    Name ="GRPNAME"
                    ControlSource ="GRPNAME"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3969
                            Top =6576
                            Width =885
                            Height =240
                            Name ="Label39"
                            Caption ="نام گروه كالا"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1190
                    Top =2267
                    Width =2268
                    ColumnWidth =615
                    ColumnOrder =5
                    TabIndex =17
                    Name ="CODE"
                    ControlSource ="CODE"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4025
                            Top =2267
                            Width =555
                            Height =240
                            Name ="Label40"
                            Caption ="كد كالا"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1303
                    Top =3174
                    Width =2268
                    TabIndex =18
                    Name ="hes"
                    ControlSource ="hes"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4138
                            Top =3178
                            Width =750
                            Height =240
                            Name ="Label43"
                            Caption ="كد شخص"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1134
                    Top =12075
                    Width =2268
                    TabIndex =19
                    Name ="USER_NAME"
                    ControlSource ="USER_NAME"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3969
                            Top =12075
                            Width =1035
                            Height =240
                            Name ="Label44"
                            Caption ="نام كاربر"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1134
                    Top =12415
                    Width =2268
                    TabIndex =20
                    Name ="SHNAME"
                    ControlSource ="SHNAME"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3969
                            Top =12415
                            Width =600
                            Height =240
                            Name ="Label45"
                            Caption ="شيفت"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1134
                    Top =12755
                    Width =2268
                    TabIndex =21
                    Name ="CUSTKNAME"
                    ControlSource ="CUSTKNAME"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3969
                            Top =12755
                            Width =960
                            Height =240
                            Name ="Label46"
                            Caption ="نوع مشتري"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1134
                    Top =13096
                    Width =2268
                    TabIndex =22
                    Name ="DEPNAME"
                    ControlSource ="DEPNAME"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3969
                            Top =13096
                            Width =330
                            Height =240
                            Name ="Label47"
                            Caption ="واحد"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =247
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Top =1417
                    Width =2268
                    TabIndex =23
                    Name ="MANDAH"
                    ControlSource ="MANDAH"
                    Begin
                        Begin Label
                            OverlapFlags =255
                            Left =1518
                            Top =1417
                            Width =750
                            Height =240
                            Name ="Label48"
                            Caption ="ملاحظات"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "KALAS.cls"
