Version =20
VersionRequired =20
Begin Form
    KeyPreview = NotDefault
    OrderByOn = NotDefault
    DefaultView =2
    AllowUpdating =4
    ViewsAllowed =2
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =11083
    RowHeight =285
    ItemSuffix =30
    Left =1590
    Top =540
    Right =13650
    Bottom =4260
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xd9ab67c1f122e440
    End
    RecordSource ="TDETA_HES3"
    Caption ="تعريف سرفصل تفصيلي سطح 3"
    OnCurrent ="[Event Procedure]"
    OnDelete ="[Event Procedure]"
    BeforeUpdate ="[Event Procedure]"
    AfterUpdate ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Tahoma"
    OnKeyPress ="[Event Procedure]"
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
            Height =7200
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =93
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3438
                    Top =113
                    ColumnWidth =825
                    ColumnOrder =0
                    Name ="N_KOL"
                    ControlSource ="N_KOL"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6230
                            Top =113
                            Width =780
                            Height =240
                            Name ="Label1"
                            Caption ="حساب كل"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3438
                    Top =453
                    ColumnWidth =990
                    ColumnOrder =1
                    TabIndex =1
                    Name ="NUMBER"
                    ControlSource ="NUMBER"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5975
                            Top =453
                            Width =1035
                            Height =240
                            Name ="Label3"
                            Caption ="حساب معين"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3438
                    Top =793
                    ColumnWidth =1140
                    ColumnOrder =2
                    TabIndex =2
                    Name ="TNUMBER"
                    ControlSource ="TNUMBER"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5970
                            Top =795
                            Width =1140
                            Height =240
                            Name ="Label5"
                            Caption ="حساب تفصيلي"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =226
                    Top =2324
                    Width =4965
                    ColumnWidth =4590
                    ColumnOrder =5
                    TabIndex =3
                    Name ="NAME"
                    ControlSource ="NAME"
                    OnDblClick ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5877
                            Top =2324
                            Width =1185
                            Height =240
                            Name ="Label7"
                            Caption ="*نام حساب"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1741
                    Top =2665
                    Width =3450
                    ColumnOrder =6
                    TabIndex =4
                    Name ="TOZIH"
                    ControlSource ="TOZIH"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =6387
                            Top =2665
                            Width =675
                            Height =240
                            Name ="Label9"
                            Caption ="توضيحات"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =247
                    Top =2976
                    Width =4965
                    ColumnOrder =7
                    TabIndex =5
                    Name ="ADDRESS"
                    ControlSource ="ADDRESS"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =6633
                            Top =2976
                            Width =450
                            Height =240
                            Name ="Label13"
                            Caption ="آدرس"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =247
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3559
                    Top =2636
                    Width =1710
                    ColumnOrder =8
                    TabIndex =6
                    Name ="TEL"
                    ControlSource ="TEL"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =6750
                            Top =2636
                            Width =390
                            Height =240
                            Name ="Label15"
                            Caption ="تلفن"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =247
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3174
                    Top =3174
                    ColumnWidth =1590
                    ColumnOrder =9
                    TabIndex =7
                    Name ="hes"
                    ControlSource ="=[N_KOL] & \"-\" & [NUMBER] & \"-\" & [TNUMBER] & \"-\" & [TNUMBER2] & \"-\" & ["
                        "TNUMBER3]"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =5966
                            Top =3174
                            Width =780
                            Height =240
                            Name ="Label19"
                            Caption ="حساب"
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
                    Left =3174
                    Top =4082
                    Width =2661
                    Height =300
                    ColumnOrder =10
                    TabIndex =8
                    Name ="ECODE"
                    ControlSource ="ECODE"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5935
                            Top =4081
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
                    Left =3174
                    Top =4422
                    Width =2661
                    Height =300
                    ColumnOrder =11
                    TabIndex =9
                    Name ="PCODE"
                    ControlSource ="PCODE"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5935
                            Top =4426
                            Width =735
                            Height =240
                            Name ="Label20"
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
                    Left =3174
                    Top =4762
                    Width =2661
                    Height =300
                    ColumnOrder =12
                    TabIndex =10
                    Name ="IYALAT"
                    ControlSource ="IYALAT"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5935
                            Top =4762
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
                    Left =3174
                    Top =5103
                    Width =2661
                    Height =300
                    ColumnOrder =13
                    TabIndex =11
                    Name ="CITY"
                    ControlSource ="CITY"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5935
                            Top =5101
                            Width =675
                            Height =240
                            Name ="Label21"
                            Caption ="شهرستا"
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
                    Left =3174
                    Top =5443
                    Width =2661
                    Height =300
                    ColumnOrder =14
                    TabIndex =12
                    Name ="MCODEM"
                    ControlSource ="MCODEM"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5935
                            Top =5443
                            Width =795
                            Height =240
                            Name ="Label22"
                            Caption ="كد ملي"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3438
                    Top =1155
                    ColumnOrder =3
                    TabIndex =13
                    Name ="TNUMBER2"
                    ControlSource ="TNUMBER2"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5970
                            Top =1155
                            Width =1230
                            Height =240
                            Name ="Label24"
                            Caption ="حساب تفصيلي2"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Top =3472
                    Width =8811
                    Height =300
                    ColumnWidth =1875
                    ColumnOrder =15
                    TabIndex =14
                    Name ="Text23"
                    ControlSource ="MOBILE"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            TextAlign =1
                            ReadingOrder =2
                            NumeralShapes =2
                            Left =8908
                            Top =3458
                            Width =2175
                            Height =240
                            Name ="Label25"
                            Caption ="موبايل جهت ارسال پيامك"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3438
                    Top =1515
                    ColumnOrder =4
                    TabIndex =15
                    Name ="TNUMBER3"
                    ControlSource ="TNUMBER3"
                    OnDblClick ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5970
                            Top =1515
                            Width =1230
                            Height =240
                            Name ="Label27"
                            Caption ="حساب تفصيلي3"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Top =1814
                    Width =8781
                    Height =317
                    ColumnWidth =7215
                    TabIndex =16
                    BackColor =13434828
                    ForeColor =8388608
                    Name ="ROUTE_NAME"
                    ControlSource ="ROUTE_NAME"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT Visit_route.ROUTE_NAME, Visit_route.ROUTE_NAME + N' - ' + CUST_HESAB.NAME"
                        " + N' - ' + CUST_HESAB.hes AS Expr1 FROM Visit_route INNER JOIN CUST_HESAB ON Vi"
                        "sit_route.HES = CUST_HESAB.hes WHERE (Visit_route.RACTIVE = 1)"
                    ColumnWidths ="0"
                    BeforeUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =8859
                            Top =1814
                            Width =1815
                            Height =330
                            ForeColor =8388608
                            Name ="Label26"
                            Caption ="مسير ويزيت"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =3140
                    Top =5952
                    Width =2661
                    Height =300
                    TabIndex =17
                    Name ="OSTANID"
                    ControlSource ="OSTANID"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT OSCODE, OSNAME FROM TCOD_OSTAN ORDER BY OSNAME"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5908
                            Top =5957
                            Width =705
                            Height =240
                            Name ="Label38"
                            Caption ="کد استان"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =3118
                    Top =6353
                    Width =2661
                    Height =300
                    TabIndex =18
                    Name ="SHAHRID"
                    ControlSource ="SHAHRID"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT CITYCODE, CITYNAME FROM TCOD_CITY ORDER BY CITYNAME"
                    ColumnWidths ="0"
                    OnEnter ="[Event Procedure]"
                    OnExit ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5886
                            Top =6347
                            Width =990
                            Height =240
                            Name ="Label40"
                            Caption ="کد شهرستان"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =93
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =1
                    NumeralShapes =2
                    IMESentenceMode =3
                    TabIndex =19
                    Name ="USERCO"
                    ControlSource ="USERCO"
                    DefaultValue ="=[Forms]![baseknow]![USERCOD]"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =2652
                            Width =750
                            Height =240
                            Name ="Label23"
                            Caption ="کد کاربر"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =1
                    NumeralShapes =2
                    IMESentenceMode =3
                    Top =340
                    TabIndex =20
                    Name ="USER_NAME"
                    ControlSource ="USER_NAME"
                    DefaultValue ="UCurrentUser()"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2367
                            Top =340
                            Width =1035
                            Height =240
                            Name ="Label28"
                            Caption ="نام کاربر"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =223
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =1
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1701
                    TabIndex =21
                    Name ="morem"
                    DefaultValue ="\"اطلاعات بيشتر\""
                    OnClick ="[Event Procedure]"
                    ConditionalFormat = Begin
                        0x010000006e000000010000000100000002000000000000000600000001000000 ,
                        0x00000000ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x2200c425c425c425220000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =255
                            Width =1035
                            Height =240
                            Name ="Label29"
                            Caption =" "
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OldBorderStyle =0
                    OverlapFlags =247
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Width =2635
                    Height =300
                    TabIndex =22
                    ForeColor =8388608
                    Name ="tob"
                    ControlSource ="tob"
                    RowSourceType ="Value List"
                    RowSource ="1;\"حقيقي\";2;\"حقوقي\""
                    ColumnWidths ="0"
                    DefaultValue ="2"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextFontCharSet =0
                            Left =2761
                            Width =1020
                            Height =300
                            ForeColor =8388608
                            Name ="Label43"
                            Caption ="شخصيت"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "TDETA_HES_SHEET3.cls"
