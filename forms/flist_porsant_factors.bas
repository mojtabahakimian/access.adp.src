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
    Width =5100
    RowHeight =345
    ItemSuffix =39
    Left =825
    Top =540
    Right =8295
    Bottom =4335
    DatasheetForeColor =33554432
    DatasheetGridlinesColor =12632256
    MaxRecords =1000000
    RecSrcDt = Begin
        0x62aad58631f3e440
    End
    RecordSource ="list_porsant_factors"
    Caption ="ليست پورسانت"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Tahoma"
    AllowFormView =0
    RecordSourceQualifier ="dbo"
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
        End
        Begin CommandButton
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
            FELineBreak = NotDefault
            SpecialEffect =2
            OldBorderStyle =0
            Width =1701
            LabelX =-1701
            FontName ="Tahoma"
            AsianLineBreak =255
        End
        Begin ListBox
            SpecialEffect =2
            Width =1701
            Height =1417
            LabelX =-1701
            FontName ="Tahoma"
        End
        Begin ComboBox
            SpecialEffect =2
            Width =1701
            LabelX =-1701
            FontName ="Tahoma"
        End
        Begin Section
            Height =10331
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
                    Left =1312
                    Top =113
                    ColumnWidth =1005
                    ColumnOrder =7
                    Name ="NUMBER"
                    ControlSource ="NUMBER"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3867
                            Top =113
                            Width =960
                            Height =240
                            Name ="Label1"
                            Caption ="شماره حواله"
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
                    Left =523
                    Top =453
                    Width =2490
                    ColumnWidth =5685
                    ColumnOrder =9
                    TabIndex =1
                    Name ="NAME"
                    ControlSource ="NAME"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3642
                            Top =453
                            Width =1185
                            Height =240
                            Name ="Label3"
                            Caption ="نام مشتري"
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
                    Left =1312
                    Top =793
                    ColumnWidth =1815
                    ColumnOrder =11
                    TabIndex =2
                    Name ="SumOfMABL_K"
                    ControlSource ="SumOfMABL_K"
                    Format ="Standard"
                    ConditionalFormat = Begin
                        0x0100000068000000010000000000000006000000000000000300000001000000 ,
                        0x00000000ffff9900000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x2d00310000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3702
                            Top =793
                            Width =1125
                            Height =240
                            Name ="Label5"
                            Caption ="مبلغ فاكتور"
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
                    Left =523
                    Top =1133
                    Width =2490
                    ColumnWidth =1035
                    ColumnOrder =1
                    TabIndex =3
                    Name ="DATE_N"
                    ControlSource ="DATE_N"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4002
                            Top =1133
                            Width =825
                            Height =240
                            Name ="Label7"
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
                    Left =523
                    Top =1474
                    Width =2490
                    ColumnWidth =1290
                    ColumnOrder =2
                    TabIndex =4
                    Name ="CUST_NO"
                    ControlSource ="CUST_NO"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4047
                            Top =1474
                            Width =780
                            Height =240
                            Name ="Label9"
                            Caption ="كد ويزيتور"
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
                    Left =1312
                    Top =1814
                    ColumnWidth =870
                    ColumnOrder =4
                    TabIndex =5
                    Name ="DARSAD"
                    ControlSource ="DARSAD"
                    ConditionalFormat = Begin
                        0x0100000068000000010000000000000006000000000000000300000001000000 ,
                        0x00000000ccffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x2d00310000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4137
                            Top =1814
                            Width =690
                            Height =240
                            Name ="Label11"
                            Caption ="درصد"
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
                    Left =1312
                    Top =2154
                    ColumnOrder =5
                    TabIndex =6
                    Name ="PURSANT"
                    ControlSource ="PURSANT"
                    Format ="Standard"
                    ConditionalFormat = Begin
                        0x0100000068000000010000000000000006000000000000000300000001000000 ,
                        0x00000000ccffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x2d00310000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4062
                            Top =2154
                            Width =765
                            Height =240
                            Name ="Label13"
                            Caption ="پورسانت"
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
                    Left =523
                    Top =2494
                    Width =2490
                    ColumnOrder =13
                    TabIndex =7
                    Name ="TOZIH"
                    ControlSource ="TOZIH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4287
                            Top =2494
                            Width =540
                            Height =240
                            Name ="Label15"
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
                    Left =523
                    Top =3514
                    Width =2490
                    ColumnWidth =3240
                    ColumnOrder =3
                    TabIndex =8
                    Name ="Expr1"
                    ControlSource ="Expr1"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3642
                            Top =3514
                            Width =1185
                            Height =240
                            Name ="Label21"
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
                    Left =523
                    Top =3855
                    Width =2490
                    ColumnWidth =2235
                    ColumnOrder =15
                    TabIndex =9
                    Name ="TEL"
                    ControlSource ="TEL"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4437
                            Top =3855
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
                    Left =523
                    Top =4195
                    Width =2490
                    ColumnOrder =16
                    TabIndex =10
                    Name ="MOBILE"
                    ControlSource ="MOBILE"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4197
                            Top =4195
                            Width =630
                            Height =240
                            Name ="Label25"
                            Caption ="موبايل"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    ReadingOrder =1
                    Left =1497
                    Top =2860
                    ColumnOrder =17
                    TabIndex =11
                    Name ="STAT"
                    ControlSource ="STAT"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextFontCharSet =178
                            Left =835
                            Top =2834
                            Width =705
                            Height =240
                            Name ="Label26"
                            Caption ="مبلغ ثابت"
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
                    Left =566
                    Top =3200
                    ColumnWidth =4035
                    ColumnOrder =14
                    TabIndex =12
                    Name ="PORID"
                    ControlSource ="PORID"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT VISITORS_PORSANT.PORID, CAST(VISITORS_PORSANT.PORID AS nvarchar) + N' - '"
                        " + CAST(VISITORS_PORSANT.VDATE AS nvarchar) + N' - ' + ISNULL(CUSTKIND.CUSTKNAME"
                        ", N'بدون گروه (همه)') + N' - ' + ISNULL(VISITORS_PORSANT.COMMENT, N' ') + N' - '"
                        " + CUST_HESAB.NAME AS Expr1 FROM VISITORS_PORSANT INNER JOIN CUST_HESAB ON VISIT"
                        "ORS_PORSANT.HES = CUST_HESAB.hes LEFT OUTER JOIN CUSTKIND ON VISITORS_PORSANT.CU"
                        "ST_COD = CUSTKIND.CUST_COD"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =178
                            Left =3355
                            Top =3194
                            Width =1650
                            Height =240
                            Name ="Label16"
                            Caption ="الگوي پرداخت پورسانت"
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
                    Left =523
                    Top =4555
                    Width =2490
                    ColumnWidth =915
                    ColumnOrder =10
                    TabIndex =13
                    Name ="AGHLAM"
                    ControlSource ="AGHLAM"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4200
                            Top =4560
                            Width =810
                            Height =240
                            Name ="Label28"
                            Caption ="تعداد اقلام"
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
                    Left =1303
                    Top =6746
                    ColumnWidth =2040
                    ColumnOrder =21
                    TabIndex =17
                    Name ="USER_NAME"
                    ControlSource ="USER_NAME"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3670
                            Top =6746
                            Width =1035
                            Height =240
                            Name ="Label32"
                            Caption ="كاربر"
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
                    Left =1303
                    Top =5725
                    ColumnWidth =1530
                    ColumnOrder =18
                    TabIndex =14
                    Name ="DEPATMAN"
                    ControlSource ="DEPATMAN"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT DEPATMAN, DEPNAME FROM DEPART"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4105
                            Top =5725
                            Width =600
                            Height =240
                            Name ="Label29"
                            Caption ="واحد"
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
                    Left =1303
                    Top =6066
                    ColumnOrder =19
                    TabIndex =15
                    Name ="SHIFT"
                    ControlSource ="SHIFT"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT SHIFT_ID, SHNAME FROM SHIFT"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4105
                            Top =6066
                            Width =600
                            Height =240
                            Name ="Label30"
                            Caption ="شيفت"
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
                    Left =1303
                    Top =6406
                    ColumnWidth =1770
                    ColumnOrder =20
                    TabIndex =16
                    Name ="CUST_KIND"
                    ControlSource ="CUST_KIND"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT CUST_COD, CUSTKNAME FROM CUSTKIND"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3745
                            Top =6406
                            Width =960
                            Height =240
                            Name ="Label31"
                            Caption ="نوع مشتري"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1757
                    Top =7086
                    ColumnOrder =8
                    TabIndex =18
                    Name ="hes"
                    ControlSource ="hes"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3570
                            Top =7080
                            Width =1530
                            Height =600
                            Name ="Label33"
                            Caption ="کد مشتري"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1757
                    Top =7806
                    ColumnWidth =4800
                    ColumnOrder =22
                    TabIndex =19
                    Name ="ROUTE_NAME"
                    ControlSource ="ROUTE_NAME"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =3570
                            Top =7800
                            Width =1530
                            Height =600
                            Name ="Label35"
                            Caption ="مسير ويزيت"
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
                    Left =963
                    Top =5159
                    ColumnWidth =1065
                    ColumnOrder =6
                    TabIndex =20
                    Name ="Text36"
                    ControlSource ="NUMBER"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3518
                            Top =5159
                            Width =960
                            Height =240
                            Name ="Label37"
                            Caption ="شماره فاكتور"
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
                    Left =1133
                    Top =8163
                    ColumnWidth =1755
                    ColumnOrder =12
                    TabIndex =21
                    Name ="Text33"
                    ControlSource ="Expr2"
                    Format ="Standard"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =3523
                            Top =8163
                            Width =1125
                            Height =240
                            Name ="Label34"
                            Caption ="مبلغ خالص"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1190
                    Top =9581
                    ColumnWidth =390
                    ColumnOrder =0
                    TabIndex =22
                    Name ="mm"
                    ControlSource ="mm"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4217
                            Top =9581
                            Width =375
                            Height =240
                            Name ="Label38"
                            Caption ="ماه"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "flist_porsant_factors.cls"
