Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    KeyPreview = NotDefault
    DefaultView =0
    AllowUpdating =4
    ScrollBars =0
    ViewsAllowed =1
    BorderStyle =3
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =8625
    RowHeight =315
    ItemSuffix =34
    Left =3315
    Top =1275
    Right =11940
    Bottom =8115
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xaf7c96670f5ce440
    End
    Caption ="ارسال پيام"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Tahoma"
    OnLoad ="[Event Procedure]"
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
            Height =6860
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
                    Left =7170
                    Top =56
                    Width =741
                    Height =330
                    ColumnWidth =630
                    ColumnOrder =0
                    Name ="IDNUM"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =7966
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
                    Locked = NotDefault
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
                    Left =4425
                    Top =528
                    Width =3285
                    Height =330
                    TabIndex =1
                    ConditionalFormat = Begin
                        0x0100000072000000030000000000000002000000000000000200000001000000 ,
                        0x00000000ccffcc00000000000200000003000000050000000100000000000000 ,
                        0xffff9900000000000200000006000000080000000100000000000000ccffff00 ,
                        0x310000000000320000000000330000000000
                    End
                    Name ="PERSONEL"
                    RowSourceType ="Value List"
                    RowSource ="1;\"حكيميان\";2;\"آقاي فلاح\";4;\"چم حيدر\";12;\"ملكي\";125;\"محمودي\";127;\"حكي"
                        "ميان -2\";131;\"انكوباتور\";137;\"تكنيك ها و مثالها\";140;\"كارآموز\";141;\"آبيا"
                        "ر\";142;\"جعفري\""
                    ColumnWidths ="0"
                    DefaultValue ="1"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =7749
                            Top =525
                            Width =735
                            Height =345
                            Name ="Label5"
                            Caption ="فرستنده:"
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    OverlapFlags =87
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =87
                    Top =1360
                    Width =8400
                    Height =3510
                    ColumnWidth =5280
                    ColumnOrder =4
                    TabIndex =3
                    BackColor =13209
                    ForeColor =16777215
                    Name ="PAYAM"
                    AfterUpdate ="[Event Procedure]"
                    OnDblClick ="[Event Procedure]"
                    OnChange ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =7650
                            Top =1020
                            Width =915
                            Height =345
                            Name ="Label7"
                            Caption ="متن پيامك"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    Locked = NotDefault
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
                    Left =6356
                    Top =5052
                    Width =1635
                    Height =330
                    ColumnWidth =1230
                    ColumnOrder =6
                    TabIndex =4
                    ConditionalFormat = Begin
                        0x0100000066000000010000000000000002000000000000000200000001000000 ,
                        0x00000000ffcc9900000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x320000000000
                    End
                    Name ="STATUS"
                    RowSourceType ="Value List"
                    RowSource ="1;\"مشاهده نشده\";2;\"مشاهده شده\""
                    ColumnWidths ="0"
                    DefaultValue ="2"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =8025
                            Top =5052
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
                    OverlapFlags =93
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =4080
                    Top =5040
                    Width =1206
                    Height =330
                    ColumnWidth =945
                    ColumnOrder =7
                    TabIndex =5
                    Name ="STDATE"
                    DefaultValue ="=FARSIDATE(Date())"
                    InputMask ="##/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5393
                            Top =5046
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
                    Left =2160
                    Top =5044
                    Width =1041
                    Height =330
                    ColumnWidth =885
                    ColumnOrder =8
                    TabIndex =6
                    Name ="STTIME"
                    DefaultValue ="=Hour(Now())*100+Minute(Now())"
                    InputMask ="##:##"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =3308
                            Top =5040
                            Width =780
                            Height =330
                            Name ="Label15"
                            Caption ="زمان ارجاع"
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
                    Left =97
                    Top =5045
                    Width =1635
                    Height =330
                    ColumnWidth =1290
                    ColumnOrder =11
                    TabIndex =7
                    Name ="USERNAME"
                    DefaultValue ="UCurrentUser()"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1748
                            Top =5045
                            Width =390
                            Height =330
                            Name ="Label21"
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
                    Left =120
                    Top =60
                    Width =6336
                    Height =330
                    ColumnWidth =4950
                    ColumnOrder =3
                    TabIndex =2
                    Name ="COMP_COD"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="COMPANYES"
                    ColumnWidths ="0"
                    OnExit ="[Event Procedure]"
                    OnDblClick ="[Event Procedure]"
                    OnNotInList ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6537
                            Top =60
                            Width =555
                            Height =330
                            Name ="Label24"
                            Caption ="گيرنده:"
                        End
                    End
                End
                Begin CommandButton
                    Cancel = NotDefault
                    OverlapFlags =85
                    Left =2154
                    Top =6129
                    Width =2220
                    Height =340
                    TabIndex =8
                    Name ="Command28"
                    Caption ="انصراف"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =4380
                    Top =6129
                    Width =2220
                    Height =340
                    TabIndex =9
                    Name ="Command33"
                    Caption ="ارسال پيامك   f3"
                    OnClick ="[Event Procedure]"
                End
            End
        End
    End
End
CodeBehindForm
' See "mesagepanel_CUSTOMER.cls"
