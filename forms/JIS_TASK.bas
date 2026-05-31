Version =20
VersionRequired =20
Begin Form
    AutoCenter = NotDefault
    AllowDeletions = NotDefault
    KeyPreview = NotDefault
    DefaultView =2
    AllowUpdating =4
    ViewsAllowed =2
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =6703
    RowHeight =315
    ItemSuffix =26
    Left =3150
    Top =30
    Right =17685
    Bottom =5610
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xddd8147f333ce440
    End
    ServerFilter ="TASK LIKE N'%صافا%'"
    RecordSource ="TASKS"
    Caption ="وظايف"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Tahoma"
    OnKeyPress ="[Event Procedure]"
    OnError ="[Event Procedure]"
    AllowFormView =0
    RecordSourceQualifier ="dbo"
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
        Begin Subform
            SpecialEffect =2
            Width =1701
            Height =1701
        End
        Begin Section
            CanGrow = NotDefault
            Height =7034
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
                    Left =3528
                    Top =113
                    Width =471
                    ColumnWidth =630
                    ColumnOrder =0
                    Name ="IDNUM"
                    ControlSource ="IDNUM"
                    OnDblClick ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1714
                            Top =113
                            Width =570
                            Height =240
                            Name ="Label1"
                            Caption ="شماره"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    ColumnHidden = NotDefault
                    RowSourceTypeInt =1
                    SpecialEffect =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =3528
                    Top =453
                    Width =1260
                    ColumnWidth =0
                    ColumnOrder =1
                    TabIndex =1
                    ConditionalFormat = Begin
                        0x010000006c000000020000000000000002000000000000000200000001000000 ,
                        0x00000000ccffcc00000000000200000003000000050000000100000000000000 ,
                        0xccffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x310000000000320000000000
                    End
                    Name ="GR"
                    ControlSource ="GR"
                    RowSourceType ="Value List"
                    RowSource ="1;\"مديريت\";2;\"پشتيباني\";3;\"فروش\""
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1721
                            Top =450
                            Width =375
                            Height =240
                            Name ="Label3"
                            Caption ="گروه"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =3528
                    Top =793
                    Width =1260
                    ColumnWidth =1650
                    ColumnOrder =2
                    TabIndex =2
                    ConditionalFormat = Begin
                        0x0100000072000000030000000000000002000000000000000200000001000000 ,
                        0x00000000ccffcc00000000000200000003000000050000000100000000000000 ,
                        0xffff9900000000000200000006000000080000000100000000000000ccffff00 ,
                        0x310000000000320000000000330000000000
                    End
                    Name ="PERSONEL"
                    ControlSource ="PERSONEL"
                    RowSourceType ="Value List"
                    RowSource ="1;\"حكيميان\";2;\"آقاي فلاح\";4;\"چم حيدر\";12;\"خانم ملكي\";15;\"خانم اسماعيلي\""
                        ";124;\"خانم دستا\";125;\"خانم محمودي\";127;\"حكيميان -2\";131;\"انكوباتور\""
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1714
                            Top =793
                            Width =840
                            Height =240
                            Name ="Label5"
                            Caption ="مجري"
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
                    Left =3528
                    Top =1133
                    Width =1260
                    ColumnWidth =12540
                    ColumnOrder =4
                    TabIndex =3
                    Name ="TASK"
                    ControlSource ="TASK"
                    AfterUpdate ="[Event Procedure]"
                    OnDblClick ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1721
                            Top =1140
                            Width =495
                            Height =240
                            Name ="Label7"
                            Caption ="وظيفه"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =3528
                    Top =1474
                    Width =1260
                    ColumnWidth =990
                    ColumnOrder =5
                    TabIndex =4
                    ConditionalFormat = Begin
                        0x0100000066000000010000000000000002000000000000000200000001000000 ,
                        0x00000000ff000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x310000000000
                    End
                    Name ="PERIORITY"
                    ControlSource ="PERIORITY"
                    RowSourceType ="Value List"
                    RowSource ="1;\"فوري\";2;\"معمولي\""
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1714
                            Top =1474
                            Width =885
                            Height =240
                            Name ="Label9"
                            Caption ="اولويت"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =3528
                    Top =1814
                    Width =1260
                    ColumnWidth =1230
                    ColumnOrder =6
                    TabIndex =5
                    ConditionalFormat = Begin
                        0x0100000066000000010000000000000002000000000000000200000001000000 ,
                        0x00000000ffcc9900000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x320000000000
                    End
                    Name ="STATUS"
                    ControlSource ="STATUS"
                    RowSourceType ="Value List"
                    RowSource ="1;\"انجام نشده\";2;\"انجام شده\";3;\"لغو شده\""
                    ColumnWidths ="0"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1714
                            Top =1814
                            Width =645
                            Height =240
                            Name ="Label11"
                            Caption ="وضعيت"
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
                    Left =3528
                    Top =2154
                    Width =471
                    ColumnWidth =945
                    ColumnOrder =7
                    TabIndex =6
                    Name ="STDATE"
                    ControlSource ="STDATE"
                    DefaultValue ="=FARSIDATE(Date())"
                    InputMask ="##/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1721
                            Top =2160
                            Width =840
                            Height =240
                            Name ="Label13"
                            Caption ="تاريخ  ارجاع"
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
                    Left =3528
                    Top =2494
                    Width =471
                    ColumnWidth =885
                    ColumnOrder =8
                    TabIndex =7
                    Name ="STTIME"
                    ControlSource ="STTIME"
                    DefaultValue ="=Hour(Now())*100+Minute(Now())"
                    InputMask ="##:##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1721
                            Top =2490
                            Width =780
                            Height =240
                            Name ="Label15"
                            Caption ="زمان ارجاع"
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
                    Left =3528
                    Top =2834
                    Width =471
                    ColumnWidth =885
                    ColumnOrder =9
                    TabIndex =8
                    Name ="ENDATE"
                    ControlSource ="ENDATE"
                    InputMask ="##/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1721
                            Top =2835
                            Width =780
                            Height =240
                            Name ="Label17"
                            Caption ="تاريخ انجام"
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
                    Left =3528
                    Top =3174
                    Width =471
                    ColumnWidth =870
                    ColumnOrder =10
                    TabIndex =9
                    Name ="ENTIME"
                    ControlSource ="ENTIME"
                    InputMask ="##:##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1721
                            Top =3180
                            Width =765
                            Height =240
                            Name ="Label19"
                            Caption ="زمان انجام"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =3528
                    Top =3514
                    Width =1260
                    ColumnWidth =1290
                    ColumnOrder =11
                    TabIndex =10
                    Name ="USERNAME"
                    ControlSource ="USERNAME"
                    DefaultValue ="UCurrentUser()"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1714
                            Top =3514
                            Width =885
                            Height =240
                            Name ="Label21"
                            Caption ="كاربر"
                        End
                    End
                End
                Begin Subform
                    OverlapFlags =87
                    Left =1588
                    Top =4138
                    Width =5115
                    Height =2445
                    TabIndex =11
                    Name ="EVENTS"
                    SourceObject ="Form.EVENTS"
                    LinkChildFields ="IDNUM"
                    LinkMasterFields ="IDNUM"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =5502
                            Top =3898
                            Width =630
                            Height =240
                            Name ="EVENTS Label"
                            Caption ="EVENTS"
                            EventProcPrefix ="EVENTS_Label"
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
                    Left =1588
                    Top =3795
                    Width =2331
                    ColumnWidth =3555
                    ColumnOrder =3
                    TabIndex =12
                    Name ="COMP_COD"
                    ControlSource ="COMP_COD"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="COMPANYES"
                    ColumnWidths ="0"
                    BeforeUpdate ="[Event Procedure]"
                    OnExit ="[Event Procedure]"
                    OnDblClick ="[Event Procedure]"
                    OnNotInList ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4000
                            Top =3795
                            Width =990
                            Height =240
                            Name ="Label24"
                            Caption ="تماس گيرنده"
                        End
                    End
                End
                Begin BoundObjectFrame
                    SizeMode =1
                    OverlapFlags =85
                    OLETypeAllowed =0
                    Left =1650
                    Top =6689
                    Width =2886
                    Height =345
                    ColumnWidth =2085
                    TabIndex =13
                    Name ="pic"
                    ControlSource ="pic"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5910
                            Top =6690
                            Width =570
                            Height =240
                            Name ="Label25"
                            Caption ="ضميمه"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "JIS_TASK.cls"
