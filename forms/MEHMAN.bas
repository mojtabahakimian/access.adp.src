Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    OrderByOn = NotDefault
    DefaultView =0
    AllowUpdating =4
    ScrollBars =0
    ViewsAllowed =1
    TabularCharSet =178
    TabularFamily =42
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =64
    GridY =64
    Width =11905
    DatasheetFontHeight =10
    ItemSuffix =67
    Left =405
    Top =780
    Right =11400
    Bottom =4965
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xb6a3d6e758a8e240
    End
    ServerFilter ="IDH IS NULL"
    RecordSource ="MEHMAN"
    Caption ="كارت ورودي مهمان"
    OnCurrent ="[Event Procedure]"
    BeforeUpdate ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Arial (Arabic)"
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
        Begin Subform
            SpecialEffect =2
            Width =1701
            Height =1701
        End
        Begin ToggleButton
            TextFontCharSet =178
            Width =283
            Height =283
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Tahoma"
        End
        Begin Section
            CanGrow = NotDefault
            Height =6803
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =8994
                    Top =75
                    Width =1236
                    Height =315
                    ForeColor =10040115
                    Name ="ID"
                    ControlSource ="ID"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =10302
                            Top =82
                            Width =525
                            Height =300
                            Name ="Label1"
                            Caption ="شماره"
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
                    Left =7747
                    Top =449
                    Width =2490
                    Height =315
                    TabIndex =1
                    BackColor =13434828
                    ForeColor =10040115
                    Name ="MNAME"
                    ControlSource ="MNAME"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =10309
                            Top =449
                            Width =300
                            Height =300
                            Name ="Label3"
                            Caption ="نام"
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
                    Left =7747
                    Top =819
                    Width =2490
                    Height =315
                    TabIndex =2
                    BackColor =13434828
                    ForeColor =10040115
                    Name ="MFAMILY"
                    ControlSource ="MFAMILY"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =10309
                            Top =819
                            Width =1020
                            Height =300
                            Name ="Label5"
                            Caption ="نام خانوادگي"
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
                    Left =7747
                    Top =1215
                    Width =2490
                    Height =315
                    TabIndex =3
                    ForeColor =10040115
                    Name ="FATHER"
                    ControlSource ="FATHER"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =10309
                            Top =1215
                            Width =570
                            Height =300
                            Name ="Label7"
                            Caption ="نام پدر"
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
                    Left =7747
                    Top =1595
                    Width =2490
                    Height =315
                    TabIndex =4
                    ForeColor =10040115
                    Name ="KNOW"
                    ControlSource ="KNOW"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =10309
                            Top =1595
                            Width =1440
                            Height =300
                            Name ="Label9"
                            Caption ="شماره شناسنامه"
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
                    Left =7747
                    Top =1960
                    Width =2490
                    Height =315
                    TabIndex =5
                    ForeColor =10040115
                    Name ="MSODUR"
                    ControlSource ="MSODUR"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =10309
                            Top =1960
                            Width =855
                            Height =300
                            Name ="Label11"
                            Caption ="محل صدور"
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
                    Left =7747
                    Top =2335
                    Width =2490
                    Height =315
                    TabIndex =6
                    ForeColor =10040115
                    Name ="BDATE"
                    ControlSource ="BDATE"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =10309
                            Top =2335
                            Width =810
                            Height =300
                            Name ="Label13"
                            Caption ="تاريخ تولد"
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
                    Left =3510
                    Top =4215
                    Width =6735
                    Height =315
                    TabIndex =11
                    ForeColor =10040115
                    Name ="ADDRESS"
                    ControlSource ="ADDRESS"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =10302
                            Top =4215
                            Width =510
                            Height =300
                            Name ="Label27"
                            Caption ="آدرس"
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
                    Left =3510
                    Top =1216
                    Width =1658
                    Height =330
                    TabIndex =15
                    ForeColor =10040115
                    Name ="VDATE"
                    ControlSource ="VDATE"
                    BeforeUpdate ="[Event Procedure]"
                    DefaultValue ="=FARSIDATE(Date())"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =6048
                            Top =1217
                            Width =1455
                            Height =300
                            Name ="Label29"
                            Caption ="تاريخ وساعت ورود"
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
                    Left =3510
                    Top =1596
                    Width =1658
                    Height =330
                    TabIndex =16
                    ForeColor =10040115
                    Name ="KDATE"
                    ControlSource ="KDATE"
                    BeforeUpdate ="[Event Procedure]"
                    AfterUpdate ="[Event Procedure]"
                    OnDblClick ="[Event Procedure]"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =6048
                            Top =1597
                            Width =1575
                            Height =300
                            Name ="Label31"
                            Caption ="تاريخ  و ساعت خروج"
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
                    Left =4260
                    Top =1961
                    Width =1688
                    Height =315
                    TabIndex =17
                    ForeColor =10040115
                    Name ="HAMRAH"
                    ControlSource ="HAMRAH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =6055
                            Top =1962
                            Width =1170
                            Height =300
                            Name ="Label33"
                            Caption ="تعداد همراهان"
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
                    Left =4260
                    Top =2719
                    Width =1688
                    Height =315
                    TabIndex =19
                    ForeColor =10040115
                    Name ="EZAFAH"
                    ControlSource ="EZAFAH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =6041
                            Top =2719
                            Width =855
                            Height =300
                            Name ="Label37"
                            Caption ="تخت اضافه"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3510
                    Top =3084
                    Width =2438
                    Height =315
                    TabIndex =20
                    ForeColor =10040115
                    Name ="MABLPISH"
                    ControlSource ="MABLPISH"
                    Format ="#,###"
                    OnKeyPress ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =6041
                            Top =3060
                            Width =945
                            Height =300
                            Name ="Label39"
                            Caption ="پيش پرداخت"
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
                    Left =3510
                    Top =3840
                    Width =2438
                    Height =315
                    TabIndex =22
                    ForeColor =10040115
                    Name ="MOLAH"
                    ControlSource ="MOLAH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =6041
                            Top =3810
                            Width =735
                            Height =300
                            Name ="Label41"
                            Caption ="ملاحظات"
                        End
                    End
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =7747
                    Top =2719
                    Width =2490
                    Height =315
                    TabIndex =7
                    ForeColor =10040115
                    Name ="TAHSIL"
                    ControlSource ="TAHSIL"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TCODE_MADRAK.CODE, TCODE_MADRAK.MADRAK_NAME FROM TCODE_MADRAK;"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =10302
                            Top =2719
                            Width =690
                            Height =300
                            Name ="Label15"
                            Caption ="تحصيلات"
                        End
                    End
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =7747
                    Top =3084
                    Width =2490
                    Height =315
                    TabIndex =8
                    ForeColor =10040115
                    Name ="JOB"
                    ControlSource ="JOB"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT MEHMAN.JOB FROM MEHMAN GROUP BY MEHMAN.JOB ORDER BY MEHMAN.JOB;"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =10302
                            Top =3084
                            Width =570
                            Height =300
                            Name ="Label17"
                            Caption ="شغل"
                        End
                    End
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =7747
                    Top =3465
                    Width =2490
                    Height =315
                    TabIndex =9
                    ForeColor =10040115
                    Name ="GHASD"
                    ControlSource ="GHASD"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT MEHMAN.GHASD FROM MEHMAN GROUP BY MEHMAN.GHASD ORDER BY MEHMAN.GHASD;"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =10302
                            Top =3465
                            Width =1005
                            Height =300
                            Name ="Label19"
                            Caption ="قصد مسافرت"
                        End
                    End
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3510
                    Top =90
                    Width =2438
                    Height =315
                    TabIndex =12
                    ForeColor =10040115
                    Name ="SCITY"
                    ControlSource ="SCITY"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT MEHMAN.SCITY FROM MEHMAN GROUP BY MEHMAN.SCITY ORDER BY MEHMAN.SCITY;"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =6041
                            Top =90
                            Width =720
                            Height =300
                            Name ="Label21"
                            Caption ="شهر مبدا"
                        End
                    End
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3510
                    Top =449
                    Width =2438
                    Height =315
                    TabIndex =13
                    ForeColor =10040115
                    Name ="DCITY"
                    ControlSource ="DCITY"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT MEHMAN.DCITY FROM MEHMAN GROUP BY MEHMAN.DCITY ORDER BY MEHMAN.DCITY;"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =6048
                            Top =449
                            Width =945
                            Height =300
                            Name ="Label23"
                            Caption ="شهر مقصد"
                        End
                    End
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3510
                    Top =819
                    Width =2438
                    Height =315
                    TabIndex =14
                    ForeColor =10040115
                    Name ="ENGIN"
                    ControlSource ="ENGIN"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT MEHMAN.ENGIN FROM MEHMAN GROUP BY MEHMAN.ENGIN ORDER BY MEHMAN.ENGIN;"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =6048
                            Top =819
                            Width =975
                            Height =300
                            Name ="Label25"
                            Caption ="وسيله سفر"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4260
                    Top =2336
                    Width =1688
                    Height =315
                    TabIndex =18
                    BackColor =13434828
                    ForeColor =10040115
                    Name ="ROOMN"
                    ControlSource ="ROOMN"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT ROOM_NUM, RFLAG FROM ROOM WHERE (RFLAG = 2) OR (RFLAG = 4)"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =6048
                            Top =2336
                            Width =900
                            Height =300
                            Name ="Label35"
                            Caption ="شماره اتاق"
                        End
                    End
                End
                Begin OptionGroup
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    Left =3450
                    Top =1950
                    Width =783
                    Height =1005
                    TabIndex =23
                    Name ="Frame51"
                    ControlSource ="SEX"
                    DefaultValue ="0"
                    Begin
                        Begin ToggleButton
                            OverlapFlags =87
                            AccessKey =227
                            TextFontCharSet =0
                            Left =3528
                            Top =2007
                            Width =675
                            Height =465
                            FontWeight =700
                            OptionValue =0
                            ForeColor =10040115
                            Name ="Toggle54"
                            Caption ="&مرد"
                        End
                        Begin ToggleButton
                            OverlapFlags =87
                            AccessKey =210
                            TextFontCharSet =0
                            Left =3525
                            Top =2475
                            Width =675
                            Height =465
                            FontWeight =700
                            ForeColor =10040115
                            Name ="Toggle55"
                            Caption ="&زن"
                        End
                    End
                End
                Begin Subform
                    OverlapFlags =85
                    Left =67
                    Top =4584
                    Width =11670
                    Height =2175
                    TabIndex =24
                    Name ="MEHMAN_sub"
                    SourceObject ="Form.MEHMAN_sub"
                    LinkChildFields ="IDH"
                    LinkMasterFields ="ID"
                    OnEnter ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            TextAlign =2
                            Left =28
                            Top =4251
                            Width =765
                            Height =300
                            Name ="MEHMAN_sub Label"
                            Caption ="همراهان"
                            EventProcPrefix ="MEHMAN_sub_Label"
                        End
                    End
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =7747
                    Top =3840
                    Width =2490
                    Height =315
                    TabIndex =10
                    ForeColor =10040115
                    Name ="SAZMANM"
                    ControlSource ="SAZMANM"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT MEHMAN.SAZMANM FROM MEHMAN GROUP BY MEHMAN.SAZMANM ORDER BY MEHMAN.SAZMAN"
                        "M;"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =10302
                            Top =3840
                            Width =1065
                            Height =300
                            Name ="Label59"
                            Caption ="سازمان/اداره"
                        End
                    End
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3510
                    Top =3466
                    Width =2438
                    Height =315
                    TabIndex =21
                    ForeColor =10040115
                    Name ="TABEI"
                    ControlSource ="TABEI"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT MEHMAN.TABEI FROM MEHMAN GROUP BY MEHMAN.TABEI ORDER BY MEHMAN.TABEI;"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =6047
                            Top =3435
                            Width =540
                            Height =300
                            Name ="Label47"
                            Caption ="تابعيت"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1410
                    Top =90
                    Width =846
                    Height =315
                    TabIndex =25
                    Name ="TAFZ"
                    ControlSource ="TAFZ"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =2343
                            Top =94
                            Width =1050
                            Height =300
                            Name ="Label60"
                            Caption ="كد حسابداري:"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =285
                    Top =465
                    Width =1971
                    Height =315
                    TabIndex =26
                    Name ="USER_NAME"
                    ControlSource ="USER_NAME"
                    DefaultValue ="uCurrentUser()"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =2352
                            Top =465
                            Width =570
                            Height =300
                            Name ="Label62"
                            Caption ="كاربر:"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5188
                    Top =1217
                    Width =773
                    Height =315
                    TabIndex =27
                    ForeColor =10040115
                    Name ="VTIME"
                    ControlSource ="VTIME"
                    DefaultValue ="Hour(Now())*100+Minute(Now())"
                    InputMask ="##:##"
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5190
                    Top =1597
                    Width =758
                    Height =315
                    TabIndex =28
                    ForeColor =10040115
                    Name ="KTIME"
                    ControlSource ="KTIME"
                    OnDblClick ="[Event Procedure]"
                    InputMask ="##:##"
                End
            End
        End
    End
End
CodeBehindForm
' See "MEHMAN.cls"
