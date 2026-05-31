Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
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
    Width =6688
    RowHeight =300
    ItemSuffix =97
    Left =1725
    Top =720
    Right =9255
    Bottom =4635
    DatasheetGridlinesColor =12632256
    MaxRecords =1000000000
    RecSrcDt = Begin
        0x70bc63089d9fe340
    End
    RecordSource ="KALAS"
    Caption ="ليست گردشها ي كالا"
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
            Height =28586
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =93
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1134
                    Top =113
                    Width =2268
                    ColumnWidth =960
                    ColumnOrder =1
                    Name ="NUMBER"
                    ControlSource ="NUMBER"
                    Format ="Fixed"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =3969
                            Top =113
                            Width =960
                            Height =240
                            Name ="Label1"
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
                    Left =1134
                    Top =453
                    Width =2268
                    ColumnWidth =1185
                    ColumnOrder =2
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
                    Left =1134
                    Top =793
                    Width =2268
                    ColumnWidth =1410
                    ColumnOrder =18
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
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1134
                    Top =1133
                    Width =2268
                    ColumnWidth =1635
                    ColumnOrder =19
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
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1134
                    Top =1474
                    Width =2268
                    ColumnWidth =1020
                    ColumnOrder =0
                    TabIndex =4
                    Name ="DATE_N"
                    ControlSource ="DATE_N"
                    OnDblClick ="[Event Procedure]"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3969
                            Top =1474
                            Width =825
                            Height =240
                            Name ="Label9"
                            Caption ="تاريخ برگه"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =93
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1134
                    Top =2010
                    Width =2268
                    ColumnWidth =990
                    ColumnOrder =20
                    TabIndex =5
                    Name ="N_S"
                    ControlSource ="N_S"
                    Format ="Fixed"
                    Begin
                        Begin Label
                            OverlapFlags =93
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
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2664
                    Top =1927
                    Width =2268
                    ColumnWidth =885
                    ColumnOrder =9
                    TabIndex =6
                    Name ="CUSTNAME"
                    ControlSource ="CUSTNAME"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5499
                            Top =1927
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
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1134
                    Top =2577
                    Width =2268
                    ColumnWidth =1215
                    ColumnOrder =21
                    TabIndex =7
                    Name ="MOLAH"
                    ControlSource ="MOLAH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3975
                            Top =2580
                            Width =1110
                            Height =240
                            Name ="Label15"
                            Caption ="توضيحات فاكتور"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1247
                    Top =2891
                    Width =2268
                    ColumnWidth =1725
                    ColumnOrder =22
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
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1134
                    Top =3514
                    Width =2268
                    ColumnWidth =870
                    ColumnOrder =12
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
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1134
                    Top =3855
                    Width =2268
                    ColumnWidth =1065
                    ColumnOrder =4
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
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1134
                    Top =4195
                    Width =2268
                    ColumnWidth =1215
                    ColumnOrder =23
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
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1134
                    Top =4535
                    Width =2268
                    ColumnWidth =1020
                    ColumnOrder =5
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
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1134
                    Top =4875
                    Width =2268
                    ColumnWidth =3075
                    ColumnOrder =11
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
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1134
                    Top =5215
                    Width =2268
                    ColumnWidth =1260
                    ColumnOrder =6
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
                    ColumnWidth =750
                    ColumnOrder =3
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
                    ColumnWidth =1230
                    ColumnOrder =24
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
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1190
                    Top =2267
                    Width =2268
                    ColumnWidth =615
                    ColumnOrder =10
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
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1303
                    Top =3174
                    Width =2268
                    ColumnWidth =855
                    ColumnOrder =25
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
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1190
                    Top =6973
                    Width =2268
                    ColumnWidth =990
                    ColumnOrder =26
                    TabIndex =19
                    Name ="USER_NAME"
                    ControlSource ="USER_NAME"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4025
                            Top =6973
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
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1190
                    Top =7313
                    Width =2268
                    ColumnWidth =645
                    ColumnOrder =27
                    TabIndex =20
                    Name ="SHNAME"
                    ControlSource ="SHNAME"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4025
                            Top =7313
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
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1190
                    Top =7653
                    Width =2268
                    ColumnWidth =1005
                    ColumnOrder =28
                    TabIndex =21
                    Name ="CUSTKNAME"
                    ControlSource ="CUSTKNAME"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4025
                            Top =7653
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
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1190
                    Top =7994
                    Width =2268
                    ColumnWidth =495
                    ColumnOrder =13
                    TabIndex =22
                    Name ="DEPNAME"
                    ControlSource ="DEPNAME"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4025
                            Top =7994
                            Width =330
                            Height =240
                            Name ="Label47"
                            Caption ="واحد"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =215
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2838
                    Top =8505
                    Width =852
                    ColumnWidth =1200
                    ColumnOrder =29
                    TabIndex =23
                    Name ="MANDAH"
                    ControlSource ="MANDAH"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            ReadingOrder =2
                            NumeralShapes =2
                            Left =2085
                            Top =8505
                            Width =1095
                            Height =240
                            Name ="Label48"
                            Caption ="ملاحظات سطر"
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
                    Left =1757
                    Top =9127
                    ColumnWidth =495
                    ColumnOrder =30
                    TabIndex =24
                    Name ="MAS"
                    ControlSource ="MAS"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4709
                            Top =9127
                            Width =450
                            Height =240
                            Name ="Label49"
                            Caption ="مدت"
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
                    Left =1700
                    Top =9474
                    ColumnWidth =990
                    ColumnOrder =31
                    TabIndex =25
                    Name ="N_RASID"
                    ControlSource ="N_RASID"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4653
                            Top =9467
                            Width =885
                            Height =240
                            Name ="Label51"
                            Caption ="محل مصرف"
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
                    Left =1700
                    Top =9841
                    ColumnWidth =1350
                    ColumnOrder =17
                    TabIndex =26
                    Name ="N_FANI"
                    ControlSource ="N_FANI"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4653
                            Top =9834
                            Width =885
                            Height =240
                            Name ="Label53"
                            Caption ="شماره فني"
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
                    Left =1700
                    Top =5669
                    ColumnWidth =795
                    ColumnOrder =32
                    TabIndex =27
                    Name ="SHARAYET"
                    ControlSource ="SHARAYET"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4187
                            Top =5669
                            Width =915
                            Height =240
                            Name ="Label54"
                            Caption ="ملاحظات"
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
                    Left =1643
                    Top =11111
                    ColumnWidth =1005
                    ColumnOrder =33
                    TabIndex =28
                    Name ="IMBAA"
                    ControlSource ="IMBAA"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4485
                            Top =11115
                            Width =900
                            Height =240
                            Name ="Label55"
                            Caption ="ارزش افزوده"
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
                    Left =1643
                    Top =11451
                    ColumnWidth =1515
                    ColumnOrder =34
                    TabIndex =29
                    Name ="HMBAA"
                    ControlSource ="HMBAA"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4380
                            Top =11445
                            Width =1410
                            Height =240
                            Name ="Label56"
                            Caption ="حساب ارزش افزوده"
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
                    Left =1643
                    Top =11792
                    ColumnWidth =690
                    ColumnOrder =35
                    TabIndex =30
                    Name ="TAMIR"
                    ControlSource ="TAMIR"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4430
                            Top =11792
                            Width =615
                            Height =240
                            Name ="Label57"
                            Caption ="بارگيري"
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
                    Left =1643
                    Top =12132
                    ColumnWidth =1605
                    ColumnOrder =36
                    TabIndex =31
                    Name ="TICMBAA"
                    ControlSource ="TICMBAA"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4230
                            Top =12135
                            Width =1500
                            Height =240
                            Name ="Label58"
                            Caption ="محاسبه ارزش افزوده"
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
                    Left =1643
                    Top =12472
                    ColumnWidth =795
                    ColumnOrder =37
                    TabIndex =32
                    Name ="OKF"
                    ControlSource ="OKF"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4605
                            Top =12465
                            Width =690
                            Height =240
                            Name ="Label59"
                            Caption ="تاييد برگه"
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
                    Left =1380
                    Top =13950
                    ColumnWidth =615
                    ColumnOrder =38
                    TabIndex =33
                    Name ="TOZIH"
                    ControlSource ="TOZIH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4182
                            Top =13950
                            Width =1545
                            Height =240
                            Name ="Label60"
                            Caption ="توضيح"
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
                    Left =1380
                    Top =14290
                    ColumnWidth =1125
                    ColumnOrder =39
                    TabIndex =34
                    Name ="B_SEF"
                    ControlSource ="B_SEF"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4195
                            Top =14292
                            Width =1965
                            Height =240
                            Name ="Label61"
                            Caption ="بهينه سفارش"
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
                    Left =1380
                    Top =14630
                    ColumnWidth =1125
                    ColumnOrder =40
                    TabIndex =35
                    Name ="N_SEF"
                    ControlSource ="N_SEF"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4182
                            Top =14630
                            Width =1545
                            Height =240
                            Name ="Label62"
                            Caption ="نقطه سفارش"
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
                    Left =1380
                    Top =14971
                    ColumnWidth =1230
                    ColumnOrder =41
                    TabIndex =36
                    Name ="MIN_M"
                    ControlSource ="MIN_M"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4152
                            Top =14971
                            Width =1575
                            Height =240
                            Name ="Label63"
                            Caption ="حداقل موجودي"
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
                    Left =1380
                    Top =15311
                    ColumnWidth =1290
                    ColumnOrder =42
                    TabIndex =37
                    Name ="MAX_M"
                    ControlSource ="MAX_M"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4122
                            Top =15311
                            Width =1605
                            Height =240
                            Name ="Label64"
                            Caption ="حد اكثر موجودي"
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
                    Left =1380
                    Top =15651
                    ColumnWidth =480
                    ColumnOrder =43
                    TabIndex =38
                    Name ="RADAH"
                    ControlSource ="RADAH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4122
                            Top =15651
                            Width =1605
                            Height =240
                            Name ="Label65"
                            Caption ="گروه"
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
                    Left =1380
                    Top =15991
                    ColumnWidth =405
                    ColumnOrder =44
                    TabIndex =39
                    Name ="KINDK"
                    ControlSource ="KINDK"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4197
                            Top =15991
                            Width =1530
                            Height =240
                            Name ="Label66"
                            Caption ="نوع"
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
                    Left =1380
                    Top =16331
                    ColumnWidth =720
                    ColumnOrder =45
                    TabIndex =40
                    Name ="MABL_F"
                    ControlSource ="MABL_F"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4077
                            Top =16331
                            Width =1650
                            Height =240
                            Name ="Label67"
                            Caption ="في"
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
                    Left =1380
                    Top =16671
                    ColumnWidth =1230
                    ColumnOrder =46
                    TabIndex =41
                    Name ="DEPART"
                    ControlSource ="DEPART"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4062
                            Top =16671
                            Width =1665
                            Height =240
                            Name ="Label68"
                            Caption ="واحد ارائه كننده"
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
                    Left =1380
                    Top =17012
                    ColumnWidth =1200
                    ColumnOrder =47
                    TabIndex =42
                    Name ="CMBAA"
                    ControlSource ="CMBAA"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4122
                            Top =17012
                            Width =1605
                            Height =240
                            Name ="Label69"
                            Caption ="مشمول ماليات"
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
                    Left =1380
                    Top =17352
                    ColumnWidth =780
                    ColumnOrder =48
                    TabIndex =43
                    Name ="VAZN"
                    ControlSource ="VAZN"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4257
                            Top =17352
                            Width =1470
                            Height =240
                            Name ="Label70"
                            Caption ="وزن واحد"
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
                    Left =1643
                    Top =17801
                    ColumnWidth =525
                    ColumnOrder =49
                    TabIndex =44
                    Name ="N_TAF"
                    ControlSource ="N_TAF"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4430
                            Top =17801
                            Width =1560
                            Height =240
                            Name ="Label71"
                            Caption ="price"
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
                    Left =1643
                    Top =18141
                    ColumnWidth =900
                    ColumnOrder =50
                    TabIndex =45
                    Name ="TOTALARZ"
                    ControlSource ="TOTALARZ"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4130
                            Top =18141
                            Width =915
                            Height =240
                            Name ="Label72"
                            Caption ="total price"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1473
                    Top =13096
                    ColumnWidth =1245
                    ColumnOrder =51
                    TabIndex =46
                    Name ="ANBARF"
                    ControlSource ="ANBARF"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4020
                            Top =13095
                            Width =1140
                            Height =240
                            Name ="Label73"
                            Caption ="شماره صادراتي"
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
                    Left =2891
                    Top =18765
                    ColumnWidth =1035
                    ColumnOrder =14
                    TabIndex =47
                    Name ="N_KOL"
                    ControlSource ="N_KOL"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5685
                            Top =18765
                            Width =930
                            Height =240
                            Name ="Label74"
                            Caption ="درصد تخفيف"
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
                    Left =2891
                    Top =19105
                    ColumnWidth =630
                    ColumnOrder =16
                    TabIndex =48
                    Name ="N_MOIN"
                    ControlSource ="N_MOIN"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5558
                            Top =19105
                            Width =735
                            Height =240
                            Name ="Label75"
                            Caption ="تخفيف"
                        End
                    End
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =3684
                    Top =19502
                    ColumnWidth =960
                    ColumnOrder =52
                    TabIndex =49
                    Name ="MM"
                    ControlSource ="MM"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT MON_ID, MON FROM MON"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6313
                            Top =19502
                            Width =375
                            Height =240
                            Name ="Label76"
                            Caption ="ماه"
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
                    Left =2494
                    Top =23017
                    ColumnWidth =1590
                    ColumnOrder =7
                    TabIndex =50
                    Name ="AVRAGE"
                    ControlSource ="AVRAGE"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5160
                            Top =23010
                            Width =870
                            Height =240
                            Name ="Label77"
                            Caption ="في ميانگين"
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
                    Left =2550
                    Top =23300
                    ColumnWidth =1590
                    ColumnOrder =8
                    TabIndex =51
                    Name ="mabrial"
                    ControlSource ="mabrial"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5310
                            Top =23295
                            Width =780
                            Height =240
                            Name ="Label78"
                            Caption ="مبلغ ريالي"
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
                    Left =1700
                    Top =21430
                    Width =2661
                    Height =300
                    ColumnWidth =975
                    ColumnOrder =55
                    TabIndex =52
                    Name ="ECODE"
                    ControlSource ="ECODE"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4461
                            Top =21429
                            Width =870
                            Height =240
                            Name ="Label16"
                            Caption ="كد اقتصادي"
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
                    Left =1700
                    Top =21770
                    Width =2661
                    Height =300
                    ColumnWidth =840
                    ColumnOrder =56
                    TabIndex =53
                    Name ="PCODE"
                    ControlSource ="PCODE"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4461
                            Top =21774
                            Width =735
                            Height =240
                            Name ="Label17"
                            Caption ="كد پستي"
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
                    Left =1700
                    Top =22110
                    Width =2661
                    Height =300
                    ColumnWidth =600
                    ColumnOrder =53
                    TabIndex =54
                    Name ="IYALAT"
                    ControlSource ="IYALAT"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4461
                            Top =22110
                            Width =660
                            Height =240
                            Name ="Label18"
                            Caption ="استان"
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
                    Left =1700
                    Top =22451
                    Width =2661
                    Height =300
                    ColumnWidth =885
                    ColumnOrder =54
                    TabIndex =55
                    Name ="CITY"
                    ControlSource ="CITY"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4461
                            Top =22449
                            Width =930
                            Height =240
                            Name ="Label83"
                            Caption ="شهرستان"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =1
                    NumeralShapes =2
                    IMESentenceMode =3
                    Top =25575
                    ColumnWidth =855
                    ColumnOrder =57
                    TabIndex =56
                    Name ="CDDATE"
                    ControlSource ="CDDATE"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2660
                            Top =25582
                            Width =750
                            Height =240
                            Name ="Label84"
                            Caption ="تاريخ ايجاد"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =1
                    NumeralShapes =2
                    IMESentenceMode =3
                    Top =25915
                    ColumnWidth =1005
                    ColumnOrder =58
                    TabIndex =57
                    Name ="CDTIME"
                    ControlSource ="CDTIME"
                    InputMask ="##:##:##"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =2690
                            Top =25912
                            Width =900
                            Height =240
                            Name ="Label85"
                            Caption ="ساعت ايجاد"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =1
                    NumeralShapes =2
                    IMESentenceMode =3
                    Top =26255
                    ColumnWidth =825
                    ColumnOrder =59
                    TabIndex =58
                    Name ="OKDATE"
                    ControlSource ="OKDATE"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2667
                            Top =26255
                            Width =735
                            Height =240
                            Name ="Label86"
                            Caption ="تاريخ چاپ"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =1
                    NumeralShapes =2
                    IMESentenceMode =3
                    Top =26595
                    ColumnWidth =975
                    ColumnOrder =60
                    TabIndex =59
                    Name ="OKTIME"
                    ControlSource ="OKTIME"
                    InputMask ="##:##:##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2690
                            Top =26602
                            Width =870
                            Height =240
                            Name ="Label87"
                            Caption ="ساعت چاپ"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ScrollBarAlign =1
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =3571
                    Top =25625
                    ColumnWidth =270
                    ColumnOrder =61
                    TabIndex =60
                    Name ="col1"
                    ControlSource ="col1"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6268
                            Top =25625
                            Width =420
                            Height =240
                            Name ="Label88"
                            Caption ="1"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =215
                    TextAlign =3
                    ScrollBarAlign =1
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =3571
                    Top =25965
                    ColumnWidth =270
                    ColumnOrder =62
                    TabIndex =61
                    Name ="col2"
                    ControlSource ="col2"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6268
                            Top =25965
                            Width =420
                            Height =240
                            Name ="Label89"
                            Caption ="2"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ScrollBarAlign =1
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =3571
                    Top =26305
                    ColumnWidth =270
                    ColumnOrder =63
                    TabIndex =62
                    Name ="col3"
                    ControlSource ="col3"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6268
                            Top =26305
                            Width =420
                            Height =240
                            Name ="Label90"
                            Caption ="3"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ScrollBarAlign =1
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =3571
                    Top =26645
                    ColumnWidth =270
                    ColumnOrder =64
                    TabIndex =63
                    Name ="col4"
                    ControlSource ="col4"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6268
                            Top =26645
                            Width =420
                            Height =240
                            Name ="Label91"
                            Caption ="4"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ScrollBarAlign =1
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =3571
                    Top =26985
                    ColumnWidth =270
                    ColumnOrder =65
                    TabIndex =64
                    Name ="col5"
                    ControlSource ="col5"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6268
                            Top =26985
                            Width =420
                            Height =240
                            Name ="Label92"
                            Caption ="5"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ScrollBarAlign =1
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =3571
                    Top =27325
                    ColumnWidth =270
                    ColumnOrder =66
                    TabIndex =65
                    Name ="col6"
                    ControlSource ="col6"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6268
                            Top =27325
                            Width =420
                            Height =240
                            Name ="Label93"
                            Caption ="6"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ScrollBarAlign =1
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =3571
                    Top =27666
                    ColumnWidth =270
                    ColumnOrder =67
                    TabIndex =66
                    Name ="col7"
                    ControlSource ="col7"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6268
                            Top =27666
                            Width =420
                            Height =240
                            Name ="Label94"
                            Caption ="7"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ScrollBarAlign =1
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =3571
                    Top =28006
                    ColumnWidth =270
                    ColumnOrder =68
                    TabIndex =67
                    Name ="col8"
                    ControlSource ="col8"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6268
                            Top =28006
                            Width =420
                            Height =240
                            Name ="Label95"
                            Caption ="8"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ScrollBarAlign =1
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =3571
                    Top =28346
                    ColumnWidth =270
                    ColumnOrder =69
                    TabIndex =68
                    Name ="col9"
                    ControlSource ="col9"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6268
                            Top =28346
                            Width =420
                            Height =240
                            Name ="Label96"
                            Caption ="9"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =247
                    TextAlign =3
                    IMESentenceMode =3
                    ColumnWidth =1455
                    ColumnOrder =15
                    TabIndex =69
                    Name ="TKHN"
                    ControlSource ="TKHN"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =2880
                            Width =1350
                            Height =240
                            Name ="Label123"
                            Caption ="درصد تخفيف نقدي"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "KALAS_MAIN.cls"
