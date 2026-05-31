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
    Width =4950
    RowHeight =345
    ItemSuffix =25
    Left =1725
    Top =150
    Right =14880
    Bottom =7515
    DatasheetForeColor =33554432
    DatasheetGridlinesColor =12632256
    MaxRecords =10000000
    RecSrcDt = Begin
        0x5096bd6e12bee440
    End
    RecordSource ="SELECT customers_whodid_purchase.CUST_NO, CUST_HESAB.* FROM CUST_HESAB LEFT OUTE"
        "R JOIN customers_whodid_purchase(1, 999999) customers_whodid_purchase ON CUST_HE"
        "SAB.hes = customers_whodid_purchase.CUST_NO WHERE (CUST_HESAB.hes LIKE N'115-%')"
        " AND (customers_whodid_purchase.CUST_NO IS NULL)"
    Caption ="ليست مشترياني كه در اين محدوده از شما خريد نكرده اند"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Tahoma"
    AllowFormView =0
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
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
        Begin Section
            Height =6973
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
                    Name ="hes"
                    ControlSource ="hes"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =120
                            Width =930
                            Height =240
                            Name ="Label1"
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
                    Left =1927
                    Top =453
                    Width =2490
                    ColumnWidth =3195
                    TabIndex =1
                    Name ="NAME"
                    ControlSource ="NAME"
                    OnDblClick ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =450
                            Width =1470
                            Height =240
                            Name ="Label3"
                            Caption ="نام حساب مشتري"
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
                    Top =850
                    ColumnWidth =4950
                    TabIndex =2
                    Name ="ADDRESS"
                    ControlSource ="ADDRESS"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3685
                            Top =850
                            Width =510
                            Height =240
                            Name ="Label14"
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
                    Left =793
                    Top =1190
                    TabIndex =3
                    Name ="CITY"
                    ControlSource ="CITY"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3715
                            Top =1190
                            Width =480
                            Height =240
                            Name ="Label15"
                            Caption ="شهر"
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
                    Top =1530
                    TabIndex =4
                    Name ="CODE_E"
                    ControlSource ="CODE_E"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3310
                            Top =1530
                            Width =885
                            Height =240
                            Name ="Label16"
                            Caption ="ساير"
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
                    Top =1871
                    TabIndex =5
                    Name ="ECODE"
                    ControlSource ="ECODE"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =3550
                            Top =1876
                            Width =720
                            Height =405
                            Name ="Label17"
                            Caption ="كد اقتصادي"
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
                    Top =2211
                    TabIndex =6
                    Name ="IYALAT"
                    ControlSource ="IYALAT"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =3535
                            Top =2211
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
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =793
                    Top =2551
                    TabIndex =7
                    Name ="MCODEM"
                    ControlSource ="MCODEM"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3400
                            Top =2551
                            Width =795
                            Height =240
                            Name ="Label19"
                            Caption ="كد ملي"
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
                    Top =2891
                    TabIndex =8
                    Name ="MOBILE"
                    ControlSource ="MOBILE"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3505
                            Top =2891
                            Width =690
                            Height =240
                            Name ="Label20"
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
                    Left =793
                    Top =3231
                    TabIndex =9
                    Name ="PCODE"
                    ControlSource ="PCODE"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3550
                            Top =3231
                            Width =645
                            Height =240
                            Name ="Label21"
                            Caption ="كد پستي"
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
                    Top =3571
                    ColumnWidth =2010
                    TabIndex =10
                    Name ="TEL"
                    ControlSource ="TEL"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3745
                            Top =3571
                            Width =450
                            Height =240
                            Name ="Label22"
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
                    Left =793
                    Top =3911
                    ColumnWidth =2625
                    TabIndex =11
                    Name ="TOZIH"
                    ControlSource ="TOZIH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3595
                            Top =3911
                            Width =600
                            Height =240
                            Name ="Label23"
                            Caption ="توضيحات"
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
                    Left =1473
                    Top =4932
                    ColumnWidth =5400
                    TabIndex =12
                    Name ="ROUTE_NAME"
                    ControlSource ="ROUTE_NAME"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3720
                            Top =4932
                            Width =1155
                            Height =240
                            Name ="Label24"
                            Caption ="مسير ويزيتور"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "customers_didnot_purchase.cls"
