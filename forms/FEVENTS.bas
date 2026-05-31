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
    Width =6633
    RowHeight =405
    ItemSuffix =24
    Left =6300
    Top =750
    Right =9645
    Bottom =2190
    DatasheetBackColor =16777164
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x87056ec6fbb9e540
    End
    RecordSource ="CRMEVENTS"
    Caption ="روي داد"
    BeforeUpdate ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Tahoma"
    OnError ="[Event Procedure]"
    AllowFormView =0
    RecordSourceQualifier ="dbo"
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
        End
        Begin CheckBox
            SpecialEffect =2
            LabelX =230
            LabelY =-30
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
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =3613
                    Top =113
                    Width =2490
                    ColumnWidth =2415
                    ColumnOrder =1
                    Name ="COMPANY_NAME"
                    ControlSource ="COMPANY_NAME"
                    OnDblClick ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1799
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
                    Left =3613
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
                            Left =1799
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
                    Left =3613
                    Top =793
                    ColumnWidth =645
                    ColumnOrder =2
                    TabIndex =2
                    Name ="INFO_TIME"
                    ControlSource ="INFO_TIME"
                    DefaultValue ="=Hour(Now())*100+Minute(Now())"
                    InputMask ="##:##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1799
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
                    Left =3613
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
                            Left =1799
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
                    Left =3613
                    Top =2154
                    ColumnWidth =990
                    ColumnOrder =6
                    TabIndex =6
                    Name ="NEXT_DATE"
                    ControlSource ="NEXT_DATE"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1799
                            Top =2154
                            Width =1245
                            Height =240
                            Name ="Label13"
                            Caption =" تماس بعدي"
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
                    Left =3613
                    Top =2494
                    ColumnWidth =1005
                    ColumnOrder =7
                    TabIndex =7
                    Name ="NEXT_TIME"
                    ControlSource ="NEXT_TIME"
                    InputMask ="##:##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1800
                            Top =2490
                            Width =1020
                            Height =240
                            Name ="Label15"
                            Caption ="ساعت تماس "
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
                    Left =3905
                    Top =2955
                    ColumnWidth =1935
                    ColumnOrder =9
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
                            Left =2091
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
                    Left =3216
                    Top =3798
                    Width =1701
                    Height =225
                    ColumnOrder =10
                    TabIndex =9
                    Name ="pic"
                    ControlSource ="pic"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5670
                            Top =3795
                            Width =570
                            Height =240
                            Name ="Label16"
                            Caption ="ضميمه"
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
                    Left =3613
                    Top =1133
                    Width =2490
                    ColumnWidth =1365
                    ColumnOrder =3
                    TabIndex =3
                    Name ="SALER"
                    ControlSource ="SALER"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT SALER FROM CRMEVENTS GROUP BY SALER ORDER BY SALER"
                    DefaultValue ="=UCURRENTUSER()"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =854
                            Top =1133
                            Width =1485
                            Height =240
                            Name ="Label7"
                            Caption ="تماس گيرنده"
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
                    Left =3613
                    Top =1474
                    Width =2490
                    ColumnWidth =1815
                    ColumnOrder =4
                    TabIndex =4
                    Name ="BUYER"
                    ControlSource ="BUYER"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT BUYER FROM CRMEVENTS GROUP BY BUYER ORDER BY BUYER"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =854
                            Top =1474
                            Width =1500
                            Height =240
                            Name ="Label9"
                            Caption ="طرف مذاكره"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1984
                    Top =4818
                    ColumnWidth =3555
                    ColumnOrder =11
                    TabIndex =10
                    Name ="PAYAM"
                    ControlSource ="PAYAM"
                    OnDblClick ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4741
                            Top =4818
                            Width =645
                            Height =240
                            Name ="Label21"
                            Caption ="پيام"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    ReadingOrder =1
                    Left =2858
                    Top =6009
                    ColumnWidth =615
                    ColumnOrder =8
                    TabIndex =11
                    Name ="miting"
                    ControlSource ="miting"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =2198
                            Top =5979
                            Width =690
                            Height =240
                            Name ="Label23"
                            Caption ="جلسه"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1247
                    Top =5555
                    Width =2196
                    ColumnWidth =690
                    TabIndex =12
                    Name ="userid"
                    ControlSource ="userid"
                    DefaultValue ="=[Forms]![BASEKNOW]![USERCOD]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4109
                            Top =5555
                            Width =1035
                            Height =240
                            Name ="Label40"
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
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =453
                    Top =3463
                    Width =2196
                    ColumnWidth =2400
                    TabIndex =13
                    Name ="CDATETI"
                    ControlSource ="CDATETI"
                    DefaultValue ="=Now()"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3318
                            Top =3458
                            Width =1380
                            Height =240
                            Name ="Label22"
                            Caption ="تاريخ و ساعت ايجاد"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "FEVENTS.cls"
