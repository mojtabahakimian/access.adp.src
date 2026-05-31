Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    KeyPreview = NotDefault
    AllowEdits = NotDefault
    DefaultView =0
    AllowUpdating =4
    ScrollBars =0
    ViewsAllowed =1
    BorderStyle =3
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =15137
    RowHeight =315
    ItemSuffix =33
    Left =60
    Top =1455
    Right =15180
    Bottom =11250
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x753911b59164e440
    End
    RecordSource ="SELECT TASKS.* FROM TASKS WHERE (IDNUM = 1) ORDER BY IDNUM"
    Caption ="وظايف"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Tahoma"
    OnError ="[Event Procedure]"
    AllowDatasheetView =0
    Begin
        Begin Label
            BackStyle =0
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
            Height =9807
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =13706
                    Top =56
                    Width =741
                    Height =330
                    ColumnWidth =630
                    ColumnOrder =0
                    Name ="IDNUM"
                    ControlSource ="IDNUM"
                    OnDblClick ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =14502
                            Top =56
                            Width =525
                            Height =330
                            Name ="Label1"
                            Caption ="شماره"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    Visible = NotDefault
                    ColumnHidden = NotDefault
                    RowSourceTypeInt =1
                    SpecialEffect =0
                    OverlapFlags =93
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =8449
                    Top =2317
                    Width =1260
                    ColumnWidth =0
                    ColumnOrder =1
                    TabIndex =10
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
                            OverlapFlags =93
                            Left =10212
                            Top =2314
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
                    Left =8002
                    Top =56
                    Width =4965
                    Height =330
                    ColumnWidth =1530
                    ColumnOrder =2
                    TabIndex =1
                    ConditionalFormat = Begin
                        0x0100000072000000030000000000000002000000000000000200000001000000 ,
                        0x00000000ccffcc00000000000200000003000000050000000100000000000000 ,
                        0xffff9900000000000200000006000000080000000100000000000000ccffff00 ,
                        0x310000000000320000000000330000000000
                    End
                    Name ="PERSONEL"
                    ControlSource ="PERSONEL"
                    RowSourceType ="Value List"
                    RowSource ="1;\"حكيميان\";2;\"آقاي فلاح\";4;\"چم حيدر\";12;\"ملكي\";125;\"محمودي\";127;\"حكي"
                        "ميان -2\";131;\"انكوباتور\";137;\"تكنيك ها و مثالها\";140;\"كارآموز\";141;\"آبيا"
                        "ر\";142;\"جعفري\""
                    ColumnWidths ="0"
                    DefaultValue ="1"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =13028
                            Top =56
                            Width =495
                            Height =330
                            Name ="Label5"
                            Caption ="گيرنده"
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
                    Left =62
                    Top =566
                    Width =14160
                    Height =1020
                    ColumnWidth =5280
                    ColumnOrder =4
                    TabIndex =3
                    BackColor =10079487
                    Name ="TASK"
                    ControlSource ="TASK"
                    AfterUpdate ="[Event Procedure]"
                    OnDblClick ="[Event Procedure]"
                    OnChange ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =14336
                            Top =737
                            Width =735
                            Height =690
                            Name ="Label7"
                            Caption ="دستور/ وظيفه/ درخواست"
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
                    Left =13095
                    Top =1934
                    Width =1260
                    Height =330
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
                    DefaultValue ="2"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =14456
                            Top =1934
                            Width =510
                            Height =330
                            Name ="Label9"
                            Caption ="اولويت"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    TabStop = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =11055
                    Top =1934
                    Width =1260
                    Height =330
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
                    DefaultValue ="1"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =12454
                            Top =1934
                            Width =585
                            Height =330
                            Name ="Label11"
                            Caption ="وضعيت"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =8775
                    Top =1927
                    Width =1206
                    Height =330
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
                            Left =10088
                            Top =1933
                            Width =840
                            Height =330
                            Name ="Label13"
                            Caption ="تاريخ  ارجاع"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =6729
                    Top =1931
                    Width =1041
                    Height =330
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
                            Left =7877
                            Top =1927
                            Width =780
                            Height =330
                            Name ="Label15"
                            Caption ="زمان ارجاع"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =93
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =11956
                    Top =2834
                    Width =471
                    ColumnWidth =885
                    ColumnOrder =9
                    TabIndex =8
                    Name ="ENDATE"
                    InputMask ="##/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =13719
                            Top =2835
                            Width =780
                            Height =240
                            Name ="Label17"
                            Caption ="تاريخ انجام"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =93
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =11956
                    Top =3174
                    Width =471
                    ColumnWidth =870
                    ColumnOrder =10
                    TabIndex =9
                    Name ="ENTIME"
                    InputMask ="##:##"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =13719
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
                    TabStop = NotDefault
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =4526
                    Top =1927
                    Width =1635
                    Height =330
                    ColumnWidth =1290
                    ColumnOrder =11
                    TabIndex =11
                    Name ="USERNAME"
                    ControlSource ="USERNAME"
                    DefaultValue ="UCurrentUser()"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6282
                            Top =1927
                            Width =375
                            Height =330
                            Name ="Label21"
                            Caption ="كاربر"
                        End
                    End
                End
                Begin Subform
                    OverlapFlags =247
                    Left =105
                    Top =2490
                    Width =14925
                    Height =6585
                    TabIndex =13
                    Name ="taskspevent"
                    SourceObject ="Form.taskspevent"
                    LinkChildFields ="IDNUM"
                    LinkMasterFields ="IDNUM"
                    OnEnter ="[Event Procedure]"
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
                    Left =105
                    Top =60
                    Width =6696
                    Height =330
                    ColumnWidth =4950
                    ColumnOrder =3
                    TabIndex =2
                    Name ="COMP_COD"
                    ControlSource ="COMP_COD"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="COMPANYES"
                    ColumnWidths ="0"
                    OnExit ="[Event Procedure]"
                    OnDblClick ="[Event Procedure]"
                    OnNotInList ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6882
                            Top =60
                            Width =990
                            Height =330
                            Name ="Label24"
                            Caption ="تماس گيرنده"
                        End
                    End
                End
                Begin BoundObjectFrame
                    TabStop = NotDefault
                    SizeMode =1
                    OverlapFlags =85
                    OLETypeAllowed =0
                    Left =46
                    Top =1870
                    Width =1251
                    Height =345
                    ColumnWidth =2085
                    TabIndex =12
                    Name ="pic"
                    ControlSource ="pic"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1418
                            Top =1923
                            Width =915
                            Height =240
                            Name ="Label25"
                            Caption ="الصاق تصوير"
                        End
                    End
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    Left =2437
                    Top =1927
                    Width =1935
                    Height =340
                    TabIndex =14
                    Name ="Command114"
                    Caption ="ثبت و مشاهده ضمائم"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    Cancel = NotDefault
                    OverlapFlags =85
                    Left =2777
                    Top =9297
                    Width =2220
                    Height =340
                    TabIndex =15
                    Name ="Command28"
                    Caption ="ذخيره و خروج  F2"
                    OnClick ="[Event Procedure]"
                End
            End
        End
    End
End
CodeBehindForm
' See "tasksp_EDIT.cls"
