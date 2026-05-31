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
    ItemSuffix =41
    Left =3315
    Top =1185
    Right =11940
    Bottom =8025
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xaf7c96670f5ce440
    End
    Caption ="يادآوري"
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
        Begin Subform
            SpecialEffect =2
            Width =1701
            Height =1701
        End
        Begin UnboundObjectFrame
            SpecialEffect =2
            OldBorderStyle =1
            Width =4536
            Height =2835
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
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =2771
                    Top =56
                    Width =3660
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
                    RowSourceType ="Value List"
                    RowSource ="1;\"حكيميان\";2;\"آقاي فلاح\";4;\"چم حيدر\";125;\"محمودي\";127;\"حكيميان -2\";13"
                        "1;\"انكوباتور\";137;\"تكنيك ها و مثالها\";140;\"كارآموز\";141;\"آبيار\";144;\"فر"
                        "وش\""
                    ColumnWidths ="0"
                    DefaultValue ="1"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6492
                            Top =56
                            Width =495
                            Height =330
                            Name ="Label5"
                            Caption ="براي"
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =90
                    Top =1633
                    Width =8400
                    Height =3585
                    ColumnWidth =5280
                    ColumnOrder =4
                    TabIndex =3
                    BackColor =16764057
                    Name ="PAYAM"
                    AfterUpdate ="[Event Procedure]"
                    OnDblClick ="[Event Procedure]"
                    OnChange ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =7265
                            Top =1350
                            Width =1185
                            Height =240
                            Name ="Label7"
                            Caption ="متن پيام يادآوري"
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
                    Left =6259
                    Top =957
                    Width =1635
                    Height =330
                    ColumnWidth =1230
                    ColumnOrder =6
                    TabIndex =6
                    ConditionalFormat = Begin
                        0x0100000066000000010000000000000002000000000000000200000001000000 ,
                        0x00000000ffcc9900000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x320000000000
                    End
                    Name ="STATUS"
                    RowSourceType ="Value List"
                    RowSource ="1;\"درجريان\";2;\"تمام شده\""
                    ColumnWidths ="0"
                    DefaultValue ="1"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =7928
                            Top =957
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
                    Left =3983
                    Top =945
                    Width =1206
                    Height =330
                    ColumnWidth =945
                    ColumnOrder =7
                    TabIndex =7
                    Name ="CTDATE"
                    DefaultValue ="=FARSIDATE(Date())"
                    InputMask ="##/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5296
                            Top =951
                            Width =840
                            Height =330
                            Name ="Label13"
                            Caption ="تاريخ  ايجاد"
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
                    Left =2063
                    Top =949
                    Width =1041
                    Height =330
                    ColumnWidth =885
                    ColumnOrder =8
                    TabIndex =8
                    Name ="CTTIME"
                    Format ="Long Time"
                    DefaultValue ="=Now()"
                    InputMask ="##:##:##"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =3211
                            Top =945
                            Width =780
                            Height =330
                            Name ="Label15"
                            Caption ="زمان ايجاد"
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
                    Top =950
                    Width =1635
                    Height =330
                    ColumnWidth =1290
                    ColumnOrder =11
                    TabIndex =9
                    Name ="USERNAME"
                    DefaultValue ="UCurrentUser()"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1651
                            Top =950
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
                    Left =99
                    Top =566
                    Width =7431
                    Height =330
                    ColumnWidth =4950
                    ColumnOrder =3
                    TabIndex =2
                    Name ="COMP_COD"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT hes, NAME FROM CUST_HESAB"
                    ColumnWidths ="0"
                    OnExit ="[Event Procedure]"
                    OnDblClick ="[Event Procedure]"
                    DefaultValue ="110"
                    OnNotInList ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =7611
                            Top =566
                            Width =990
                            Height =330
                            Name ="Label24"
                            Caption ="درخصوص"
                        End
                    End
                End
                Begin CommandButton
                    Cancel = NotDefault
                    OverlapFlags =85
                    Left =283
                    Top =6066
                    Width =2220
                    Height =340
                    TabIndex =11
                    Name ="Command28"
                    Caption ="انصراف"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =2505
                    Top =6060
                    Width =2220
                    Height =340
                    TabIndex =10
                    Name ="Command29"
                    Caption ="ذخيره"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =93
                    Left =99
                    Top =56
                    Width =2565
                    Height =340
                    TabIndex =12
                    Name ="Command30"
                    Caption ="گيرندگان بيشتر"
                    OnClick ="[Event Procedure]"
                End
                Begin TextBox
                    Visible = NotDefault
                    TabStop = NotDefault
                    OldBorderStyle =1
                    OverlapFlags =87
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =38
                    Top =56
                    Width =75
                    Height =330
                    TabIndex =13
                    Name ="USS"
                End
                Begin CheckBox
                    OverlapFlags =85
                    ReadingOrder =1
                    Left =2574
                    Top =5441
                    TabIndex =14
                    Name ="SMSOK"
                    DefaultValue ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =90
                            Top =5385
                            Width =2460
                            Height =240
                            Name ="Label35"
                            Caption ="براي يادآوري پيامك هم ارسال شود"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =6195
                    Top =5385
                    Width =1206
                    Height =330
                    TabIndex =4
                    Name ="STDATE"
                    DefaultValue ="=FARSIDATE(Date())"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =7508
                            Top =5442
                            Width =990
                            Height =240
                            Name ="Label37"
                            Caption ="تاريخ يادآوري:"
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
                    Left =2895
                    Top =5385
                    Width =2196
                    Height =330
                    TabIndex =5
                    Name ="STTIME"
                    DefaultValue ="=Format(Now(),\"hh:mm:ss\")"
                    InputMask ="00:00:00;0;_"
                    Begin
                        Begin Label
                            OverlapFlags =87
                            Left =5205
                            Top =5445
                            Width =975
                            Height =240
                            Name ="Label39"
                            Caption ="زمان يادآوري:"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "REMAINDERPANEL.cls"
