Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    DefaultView =0
    AllowUpdating =4
    ViewsAllowed =1
    TabularCharSet =178
    TabularFamily =42
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =15080
    DatasheetFontHeight =10
    ItemSuffix =30
    Left =1860
    Top =780
    Right =20760
    Bottom =10065
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x1ca18c2ecb81e240
    End
    RecordSource ="WORKHEAD"
    Caption ="ثبت كاركرد پرسنل"
    OnCurrent ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial (Arabic)"
    OnLoad ="[Event Procedure]"
    OnError ="[Event Procedure]"
    AllowDatasheetView =0
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
        Begin Section
            CanGrow = NotDefault
            Height =9075
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =13098
                    Top =75
                    Width =1071
                    Height =315
                    Name ="WDATE"
                    ControlSource ="WDATE"
                    BeforeUpdate ="[Event Procedure]"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            TextAlign =1
                            Left =14286
                            Top =75
                            Width =615
                            Height =300
                            Name ="Label1"
                            Caption ="تاريخ"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =9797
                    Top =435
                    Width =4380
                    Height =315
                    TabIndex =1
                    Name ="COMMENT"
                    ControlSource ="COMMENT"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            TextAlign =1
                            Left =14264
                            Top =435
                            Width =750
                            Height =300
                            Name ="Label3"
                            Caption ="توضيحات"
                        End
                    End
                End
                Begin Subform
                    OverlapFlags =85
                    Left =63
                    Top =855
                    Width =14850
                    Height =8220
                    TabIndex =2
                    Name ="WORKING_SUB2"
                    SourceObject ="Form.WORKING_SUB2"
                    LinkChildFields ="WDATE"
                    LinkMasterFields ="WDATE"
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    DecimalPlaces =0
                    SpecialEffect =0
                    OverlapFlags =95
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4325
                    Top =435
                    Width =1041
                    Height =300
                    TabIndex =3
                    BackColor =32768
                    ForeColor =65535
                    Name ="N_S"
                    ControlSource ="N_S"
                    OnDblClick ="[Event Procedure]"
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =93
                            TextFontCharSet =0
                            Left =5375
                            Top =435
                            Width =1485
                            Height =300
                            BackColor =32768
                            ForeColor =65535
                            Name ="Label67"
                            Caption ="ثبت دراسناد شماره:"
                        End
                    End
                End
                Begin CommandButton
                    Enabled = NotDefault
                    OverlapFlags =85
                    AccessKey =213
                    TextFontCharSet =0
                    Left =527
                    Top =75
                    Width =945
                    Height =690
                    FontWeight =700
                    TabIndex =6
                    Name ="Command22"
                    Caption ="&صدور سند پرداخت"
                    OnClick ="[Event Procedure]"
                End
                Begin TextBox
                    Visible = NotDefault
                    TabStop = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =93
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2991
                    Width =411
                    Height =345
                    FontWeight =700
                    TabIndex =7
                    Name ="LETSANAD"
                    DefaultValue ="False"
                    FontName ="Arabic Transparent"
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =11055
                    Top =60
                    Width =426
                    Height =315
                    TabIndex =8
                    Name ="FRID"
                    ControlSource ="FRID"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            TextAlign =1
                            Left =11604
                            Top =64
                            Width =1455
                            Height =240
                            Name ="Label24"
                            Caption ="تعداد روزهاي تعطيل"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =93
                    TextFontCharSet =0
                    Left =6330
                    Top =30
                    Width =1230
                    Height =405
                    FontWeight =700
                    TabIndex =9
                    Name ="Command25"
                    Caption ="  فايل ساعت"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontCharSet =0
                    Left =1487
                    Top =75
                    Width =945
                    Height =690
                    FontWeight =700
                    TabIndex =10
                    ForeColor =16711680
                    Name ="Command26"
                    Caption ="پردازش و محاسبه"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    OverlapFlags =247
                    TextFontCharSet =0
                    Left =2492
                    Top =30
                    Width =1290
                    Height =405
                    FontWeight =700
                    TabIndex =11
                    Name ="Command27"
                    Caption ="ديسك  تجارت"
                    OnClick ="[Event Procedure]"
                End
                Begin CheckBox
                    OverlapFlags =93
                    ReadingOrder =1
                    Left =9600
                    Top =120
                    Width =396
                    Height =189
                    TabIndex =12
                    Name ="TENDAR"
                    ControlSource ="TENDAR"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =10051
                            Top =90
                            Width =900
                            Height =240
                            Name ="Label316"
                            Caption ="ده درصدي"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontCharSet =0
                    Left =3782
                    Top =30
                    Width =1290
                    Height =405
                    FontWeight =700
                    TabIndex =13
                    Name ="Command28"
                    Caption ="ديسك  ملت"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    OverlapFlags =223
                    TextFontCharSet =0
                    Left =5070
                    Top =30
                    Width =1290
                    Height =405
                    FontWeight =700
                    TabIndex =14
                    Name ="Command29"
                    Caption ="ديسك  مهر"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =93
                    TextFontCharSet =0
                    Left =8277
                    Top =283
                    Width =1170
                    Height =465
                    FontWeight =700
                    TabIndex =15
                    Name ="ESLAH"
                    Caption ="اصلاح "
                    OnClick ="[Event Procedure]"
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =6008
                    Top =116
                    TabIndex =16
                    Name ="CRT"
                    ControlSource ="CRT"
                    DefaultValue ="Now()"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =8213
                            Top =113
                            Width =1515
                            Height =240
                            Name ="Label35"
                            Caption ="تاريخ ايجاد"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =5952
                    Top =396
                    TabIndex =17
                    Name ="UID"
                    ControlSource ="UID"
                    DefaultValue ="[Forms]![BASEKNOW]![USERCOD]"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =8157
                            Top =398
                            Width =1365
                            Height =240
                            Name ="Label36"
                            Caption ="کد کاربر ايجاد کننده"
                        End
                    End
                End
                Begin ComboBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =95
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    DecimalPlaces =0
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =3410
                    Top =435
                    Width =906
                    Height =300
                    TabIndex =4
                    BackColor =32768
                    ForeColor =65535
                    Name ="N_S2"
                    ControlSource ="N_S2"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT BAYEG, N_S FROM DEED_HED"
                    ColumnWidths ="0"
                    OnDblClick ="[Event Procedure]"
                End
                Begin ComboBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =87
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    DecimalPlaces =0
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =2495
                    Top =435
                    Width =906
                    Height =300
                    TabIndex =5
                    BackColor =32768
                    ForeColor =65535
                    Name ="N_S3"
                    ControlSource ="N_S3"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT BAYEG, N_S FROM DEED_HED"
                    ColumnWidths ="0"
                    OnDblClick ="[Event Procedure]"
                End
            End
        End
    End
End
CodeBehindForm
' See "WORKHEAD.cls"
