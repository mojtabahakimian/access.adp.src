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
    Width =14853
    RowHeight =1215
    ItemSuffix =36
    Left =15
    Top =990
    Right =14865
    Bottom =6660
    DatasheetGridlinesColor =12632256
    MaxRecords =1000000
    RecSrcDt = Begin
        0x3cab2dac0c64e440
    End
    RecordSource ="select * from mesagep where personel = 1 or username  ='حكيميان'  ORDER BY IDNUM"
    Caption ="بايگاني پيامها"
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
            Height =5442
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
                    ColumnOrder =3
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
                            Caption ="گيرنده"
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
                    Left =122
                    Top =1303
                    Width =14670
                    Height =3345
                    ColumnWidth =13305
                    ColumnOrder =5
                    TabIndex =3
                    BackColor =1644726
                    ForeColor =16777215
                    Name ="PAYAM"
                    ControlSource ="PAYAM"
                    AfterUpdate ="[Event Procedure]"
                    OnDblClick ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =14204
                            Top =963
                            Width =645
                            Height =240
                            Name ="Label7"
                            Caption ="متن پيام"
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
                    Left =60
                    Top =105
                    Width =1575
                    Height =330
                    ColumnWidth =1230
                    ColumnOrder =7
                    TabIndex =4
                    Name ="STATUS"
                    ControlSource ="STATUS"
                    RowSourceType ="Value List"
                    RowSource ="1;\"مشاهده نشده\";2;\"مشاهده شده\""
                    ColumnWidths ="0"
                    DefaultValue ="1"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1669
                            Top =105
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
                    Left =4207
                    Top =105
                    Width =1206
                    Height =330
                    ColumnWidth =945
                    ColumnOrder =1
                    TabIndex =5
                    Name ="STDATE"
                    ControlSource ="STDATE"
                    DefaultValue ="=FARSIDATE(Date())"
                    InputMask ="##/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5520
                            Top =111
                            Width =1005
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
                    Left =2287
                    Top =109
                    Width =1041
                    Height =330
                    ColumnWidth =885
                    ColumnOrder =2
                    TabIndex =6
                    Name ="STTIME"
                    ControlSource ="STTIME"
                    DefaultValue ="=Hour(Now())*100+Minute(Now())"
                    InputMask ="##:##"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =3435
                            Top =105
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
                    Left =6633
                    Top =105
                    Width =1590
                    Height =330
                    ColumnWidth =1290
                    ColumnOrder =4
                    TabIndex =7
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
                    ColumnOrder =6
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
                            Caption ="تماس گيرنده"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =4308
                    Top =4876
                    Width =2220
                    Height =340
                    TabIndex =8
                    Name ="Command33"
                    Caption ="پاسخ"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =6530
                    Top =4875
                    Width =2220
                    Height =340
                    TabIndex =9
                    Name ="Command37"
                    Caption ="ارجاع به ديگري (Forward)"
                    OnClick ="[Event Procedure]"
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =570
                    Top =598
                    Width =2556
                    Height =330
                    TabIndex =10
                    Name ="CRT"
                    ControlSource ="CRT"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3222
                            Top =598
                            Width =750
                            Height =240
                            Name ="Label34"
                            Caption ="تاريخ ايجاد"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4019
                    Top =598
                    Height =330
                    TabIndex =11
                    Name ="UID"
                    ControlSource ="UID"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5770
                            Top =570
                            Width =405
                            Height =240
                            Name ="Label35"
                            Caption ="کاربر"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "mesage_dtl.cls"
