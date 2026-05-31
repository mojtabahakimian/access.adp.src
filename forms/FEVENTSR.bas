Version =20
VersionRequired =20
Begin Form
    DefaultView =2
    AllowUpdating =4
    ViewsAllowed =2
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =5625
    RowHeight =330
    ItemSuffix =21
    Left =4380
    Top =2010
    Right =9465
    Bottom =3450
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xaba5ba65fdb9e540
    End
    RecordSource ="SELECT CRMEVENTS.COMPANY_NAME, CRMEVENTS.INFO_TIME, CRMEVENTS.INFO_DATE, CRMEVEN"
        "TS.SALER, CRMEVENTS.BUYER, CRMEVENTS.COMMENT, CRMEVENTS.NEXT_DATE, CRMEVENTS.NEX"
        "T_TIME, CRMEVENTS.STATUS, CRMEVENTS.pic, COPMANES.PERNUM, COPMANES.USER_NAME, CO"
        "PMANES.date_sabt, COPMANES.dt FROM COPMANES INNER JOIN CRMEVENTS ON COPMANES.id "
        "= CRMEVENTS.idc"
    Caption ="گزارش جزييات"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Tahoma"
    AllowFormView =0
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
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
        Begin ComboBox
            SpecialEffect =2
            Width =1701
            LabelX =-1701
            FontName ="Tahoma"
        End
        Begin Section
            Height =6870
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
                    Left =2608
                    Top =113
                    Width =2490
                    ColumnWidth =2370
                    ColumnOrder =1
                    Name ="COMPANY_NAME"
                    ControlSource ="COMPANY_NAME"
                    OnDblClick ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =794
                            Top =113
                            Width =780
                            Height =240
                            Name ="Label1"
                            Caption ="نام شركت"
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
                    Left =2608
                    Top =453
                    ColumnWidth =945
                    ColumnOrder =0
                    TabIndex =1
                    Name ="INFO_DATE"
                    ControlSource ="INFO_DATE"
                    DefaultValue ="=FARSIDATE(Date())"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =794
                            Top =453
                            Width =930
                            Height =240
                            Name ="Label3"
                            Caption ="تاريخ"
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
                    Left =2608
                    Top =793
                    ColumnWidth =645
                    ColumnOrder =12
                    TabIndex =2
                    Name ="INFO_TIME"
                    ControlSource ="INFO_TIME"
                    DefaultValue ="=Hour(Now())*100+Minute(Now())"
                    InputMask ="##:##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =794
                            Top =793
                            Width =900
                            Height =240
                            Name ="Label5"
                            Caption ="ساعت"
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
                    Left =2608
                    Top =1133
                    Width =2490
                    ColumnWidth =1365
                    ColumnOrder =4
                    TabIndex =3
                    Name ="SALER"
                    ControlSource ="SALER"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =794
                            Top =1133
                            Width =540
                            Height =240
                            Name ="Label7"
                            Caption ="تماس گيرنده"
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
                    Left =2608
                    Top =1474
                    Width =2490
                    ColumnWidth =1815
                    ColumnOrder =3
                    TabIndex =4
                    Name ="BUYER"
                    ControlSource ="BUYER"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =794
                            Top =1474
                            Width =555
                            Height =240
                            Name ="Label9"
                            Caption ="طرف مذاكره"
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
                    Left =2608
                    Top =1814
                    Width =2490
                    ColumnWidth =7200
                    ColumnOrder =5
                    TabIndex =5
                    Name ="COMMENT"
                    ControlSource ="COMMENT"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =794
                            Top =1814
                            Width =825
                            Height =240
                            Name ="Label11"
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
                    Left =2608
                    Top =2154
                    ColumnWidth =1350
                    ColumnOrder =9
                    TabIndex =6
                    Name ="NEXT_DATE"
                    ControlSource ="NEXT_DATE"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =794
                            Top =2154
                            Width =930
                            Height =240
                            Name ="Label13"
                            Caption ="تاريخ تماس بعدي"
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
                    Left =2608
                    Top =2494
                    ColumnOrder =7
                    TabIndex =7
                    Name ="NEXT_TIME"
                    ControlSource ="NEXT_TIME"
                    InputMask ="##:##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =794
                            Top =2494
                            Width =900
                            Height =240
                            Name ="Label15"
                            Caption ="ساعت تماس بعدي"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =2900
                    Top =2955
                    ColumnOrder =8
                    TabIndex =8
                    ConditionalFormat = Begin
                        0x010000006a000000010000000000000000000000000000000200000001000000 ,
                        0x00000000ccffcc00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x31000000310030000000
                    End
                    Name ="STATUS"
                    ControlSource ="STATUS"
                    RowSourceType ="Value List"
                    RowSource ="1;\"شروع ارتباط\";2;\"درحال مذاكره\";3;\"قرار ملاقات\";4;\"ملاقات شده \";5;\"در "
                        "شرف قرار داد\";6;\"راكد\";7;\"قرارداد بسته شد\";8;\"روي ميز\""
                    ColumnWidths ="0"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1086
                            Top =2955
                            Width =585
                            Height =240
                            Name ="Label27"
                            Caption ="وضعيت"
                        End
                    End
                End
                Begin BoundObjectFrame
                    OverlapFlags =85
                    OLETypeAllowed =0
                    Left =2211
                    Top =3798
                    Width =1701
                    Height =225
                    ColumnOrder =6
                    TabIndex =9
                    Name ="pic"
                    ControlSource ="pic"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4665
                            Top =3795
                            Width =570
                            Height =240
                            Name ="Label16"
                            Caption ="ضميمه"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =2267
                    Top =4478
                    ColumnOrder =10
                    TabIndex =10
                    Name ="PERNUM"
                    ControlSource ="PERNUM"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4623
                            Top =4478
                            Width =990
                            Height =240
                            Name ="Label17"
                            Caption ="تعداد پرسنل:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =2324
                    Top =5272
                    ColumnOrder =2
                    TabIndex =11
                    Name ="USER_NAME"
                    ControlSource ="USER_NAME"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4578
                            Top =5272
                            Width =1035
                            Height =240
                            Name ="Label18"
                            Caption ="کاربر"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1813
                    Top =5782
                    ColumnWidth =2085
                    ColumnOrder =13
                    TabIndex =12
                    Name ="date_sabt"
                    ControlSource ="date_sabt"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4345
                            Top =5782
                            Width =870
                            Height =240
                            Name ="Label19"
                            Caption ="تاريخ ثبت "
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1870
                    Top =6462
                    ColumnOrder =11
                    TabIndex =13
                    Name ="dt"
                    ControlSource ="dt"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4980
                            Top =6465
                            Width =645
                            Height =405
                            Name ="Label20"
                            Caption ="تاريخ فارسي"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "FEVENTSR.cls"
