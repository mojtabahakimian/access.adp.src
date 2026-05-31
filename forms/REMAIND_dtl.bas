Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    AllowDeletions = NotDefault
    AllowAdditions = NotDefault
    KeyPreview = NotDefault
    AllowEdits = NotDefault
    AllowUpdating =4
    ScrollBars =2
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =14895
    RowHeight =345
    ItemSuffix =36
    Left =15
    Top =990
    Right =14865
    Bottom =6660
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x0dfd13548f64e440
    End
    RecordSource ="SELECT * FROM REMAINDER WHERE (PERSONEL = 1) ORDER BY IDNUM"
    Caption ="بايگاني يادآوري"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Tahoma"
    OnError ="[Event Procedure]"
    RecordSourceQualifier ="dbo"
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
        Begin Section
            CanGrow = NotDefault
            Height =3798
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
                    Left =13392
                    Top =105
                    Width =741
                    Height =330
                    ColumnWidth =630
                    ColumnOrder =0
                    Name ="IDNUM"
                    ControlSource ="IDNUM"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =14188
                            Top =105
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
                    Left =9068
                    Top =105
                    Width =3585
                    Height =330
                    ColumnWidth =1530
                    ColumnOrder =6
                    TabIndex =1
                    Name ="PERSONEL"
                    ControlSource ="PERSONEL"
                    RowSourceType ="Value List"
                    RowSource ="1;\"حكيميان\";2;\"آقاي فلاح\";4;\"چم حيدر\";125;\"محمودي\";127;\"حكيميان -2\";13"
                        "1;\"انكوباتور\";137;\"تكنيك ها و مثالها\";140;\"كارآموز\";141;\"آبيار\";144;\"فر"
                        "وش\""
                    ColumnWidths ="0"
                    DefaultValue ="1"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =12714
                            Top =105
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
                    TextAlign =2
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =167
                    Top =1303
                    Width =14625
                    Height =1305
                    ColumnWidth =6360
                    ColumnOrder =3
                    TabIndex =3
                    BackColor =16764057
                    Name ="PAYAM"
                    ControlSource ="PAYAM"
                    AfterUpdate ="[Event Procedure]"
                    OnDblClick ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =13545
                            Top =960
                            Width =1200
                            Height =240
                            Name ="Label7"
                            Caption ="متن  يادآوري"
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
                    Top =120
                    Width =1245
                    Height =330
                    ColumnWidth =1230
                    ColumnOrder =5
                    TabIndex =4
                    Name ="STATUS"
                    ControlSource ="STATUS"
                    RowSourceType ="Value List"
                    RowSource ="1;\"درجريان\";2;\"تمام شده\";3;\"لغو شده\""
                    ColumnWidths ="0"
                    DefaultValue ="1"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =1279
                            Top =120
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
                    OldBorderStyle =1
                    OverlapFlags =93
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =6633
                    Top =105
                    Width =1590
                    Height =330
                    ColumnWidth =1290
                    ColumnOrder =7
                    TabIndex =5
                    Name ="USERNAME"
                    ControlSource ="USERNAME"
                    DefaultValue ="UCurrentUser()"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =8284
                            Top =105
                            Width =735
                            Height =240
                            Name ="Label21"
                            Caption ="فرستنده"
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
                    Left =6321
                    Top =615
                    Width =7431
                    Height =330
                    ColumnWidth =4950
                    ColumnOrder =4
                    TabIndex =2
                    Name ="COMP_COD"
                    ControlSource ="COMP_COD"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT hes, NAME FROM CUST_HESAB"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =13833
                            Top =615
                            Width =990
                            Height =330
                            Name ="Label24"
                            Caption ="در خصوص"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =6292
                    Top =3344
                    Width =2220
                    Height =340
                    TabIndex =6
                    Name ="Command33"
                    Caption ="لغو"
                    OnClick ="[Event Procedure]"
                End
                Begin CheckBox
                    OverlapFlags =85
                    ReadingOrder =1
                    Left =2767
                    Top =2777
                    TabIndex =7
                    Name ="SMSOK"
                    ControlSource ="SMSOK"
                    DefaultValue ="0"
                    Begin
                        Begin Label
                            OverlapFlags =87
                            Left =283
                            Top =2721
                            Width =2460
                            Height =240
                            Name ="Label35"
                            Caption ="براي يادآوري پيامك هم ارسال شود"
                        End
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    OverlapFlags =93
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =6388
                    Top =2721
                    Width =1206
                    Height =330
                    TabIndex =8
                    Name ="Text34"
                    ControlSource ="STDATE"
                    DefaultValue ="=FARSIDATE(Date())"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =7701
                            Top =2778
                            Width =990
                            Height =240
                            Name ="Label37"
                            Caption ="تاريخ يادآوري:"
                        End
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =3088
                    Top =2721
                    Width =2196
                    Height =330
                    TabIndex =9
                    Name ="Text35"
                    ControlSource ="STTIME"
                    Format ="Long Time"
                    DefaultValue ="=Format(Now(),\"hh:mm:ss AMPM\")"
                    InputMask ="00:00:00\\ >LL;0;_"
                    Begin
                        Begin Label
                            OverlapFlags =87
                            Left =5398
                            Top =2781
                            Width =975
                            Height =240
                            Name ="Label39"
                            Caption ="زمان يادآوري:"
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
                    Left =4470
                    Top =120
                    Width =1206
                    Height =330
                    TabIndex =10
                    Name ="CTDATE"
                    DefaultValue ="=FARSIDATE(Date())"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =5783
                            Top =126
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
                    OverlapFlags =87
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1860
                    Top =124
                    Width =1776
                    Height =330
                    TabIndex =11
                    Name ="CTTIME"
                    Format ="Long Time"
                    DefaultValue ="=Now()"
                    InputMask ="##:##:##"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =3743
                            Top =120
                            Width =780
                            Height =330
                            Name ="Label15"
                            Caption ="زمان ايجاد"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "REMAIND_dtl.cls"
