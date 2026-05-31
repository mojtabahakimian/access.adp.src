Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    AllowEdits = NotDefault
    DefaultView =0
    AllowUpdating =4
    ScrollBars =0
    TabularCharSet =178
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =12930
    RowHeight =300
    ItemSuffix =31
    Left =2085
    Top =975
    Right =13020
    Bottom =6555
    DatasheetForeColor =33554432
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x36efeae052dbe440
    End
    RecordSource ="TR_PMORAKH"
    Caption ="ثبت و ويرايش مرخصي"
    OnCurrent ="[Event Procedure]"
    AfterUpdate ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Tahoma"
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
        Begin FormHeader
            Height =630
            BackColor =8421376
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    TextFontFamily =10
                    Left =5026
                    Top =30
                    Width =1770
                    Height =600
                    FontSize =16
                    ForeColor =10092543
                    Name ="Label8"
                    Caption ="ثبت مرخصي"
                    FontName ="Titr"
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    TextFontFamily =10
                    Left =5041
                    Top =15
                    Width =1770
                    Height =600
                    FontSize =16
                    Name ="Label9"
                    Caption ="ثبت مرخصي"
                    FontName ="Titr"
                End
            End
        End
        Begin Section
            SpecialEffect =1
            Height =6472
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5161
                    Top =1245
                    Width =2106
                    Height =315
                    ColumnWidth =1485
                    TabIndex =1
                    Name ="MODATE"
                    ControlSource ="MODATE"
                    BeforeUpdate ="[Event Procedure]"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="=FARSIDATE(Date())"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =7368
                            Top =1245
                            Width =1185
                            Height =300
                            Name ="Label3"
                            Caption ="تاريخ درخواست"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5178
                    Top =4500
                    Width =2106
                    Height =315
                    ColumnWidth =1365
                    TabIndex =12
                    BackColor =13434828
                    Name ="MORAKHDAY"
                    ControlSource ="MORAKHDAY"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =7383
                            Top =4500
                            Width =1290
                            Height =300
                            Name ="Label5"
                            Caption ="مرخصي به دقيقه"
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
                    ColumnCount =2
                    Left =3121
                    Top =855
                    Width =4131
                    Height =315
                    ColumnWidth =3990
                    Name ="CODE"
                    ControlSource ="CODE"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT CODE, PNAME + N' ' + PFAMILY + N' ' + RTRIM(CAST(CODE AS NVARCHAR)) AS PE"
                        "R FROM PERSONEL ORDER BY PNAME + N' ' + PFAMILY + N' ' + RTRIM(CAST(CODE AS NVAR"
                        "CHAR))"
                    ColumnWidths ="0"
                    AfterUpdate ="[Event Procedure]"
                    OnNotInList ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =7308
                            Top =855
                            Width =525
                            Height =300
                            Name ="Label1"
                            Caption ="پرسنل"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5178
                    Top =4880
                    Width =2106
                    Height =315
                    ColumnWidth =1800
                    TabIndex =13
                    BackColor =16777164
                    Name ="MAND"
                    ControlSource ="MAND"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =7383
                            Top =4875
                            Width =1725
                            Height =300
                            Name ="Label7"
                            Caption ="مانده مرخصي به دقيقه"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =6602
                    Top =3708
                    Width =666
                    Height =315
                    ColumnWidth =1545
                    TabIndex =8
                    BackColor =13434828
                    Name ="minu"
                    Format ="Fixed"
                    ValidationRule ="Not Is Null And <60"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="0"
                    ControlTipText ="دقيقه"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =7368
                            Top =3705
                            Width =1590
                            Height =300
                            Name ="Label11"
                            Caption ="مرخصي درخواستي"
                        End
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =6602
                    Top =4088
                    Width =666
                    Height =315
                    ColumnWidth =1335
                    TabIndex =11
                    BackColor =16777164
                    Name ="MMINU"
                    Format ="Fixed"
                    DefaultValue ="0"
                    ControlTipText ="دقيقه"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =7367
                            Top =4088
                            Width =1080
                            Height =300
                            Name ="Label13"
                            Caption ="مانده مرخصي"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5879
                    Top =3708
                    Width =666
                    Height =315
                    TabIndex =7
                    BackColor =13434828
                    Name ="HOU"
                    Format ="Fixed"
                    ValidationRule ="Not Is Null And <8"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="0"
                    ControlTipText ="ساعت"
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    TabStop = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5879
                    Top =4088
                    Width =666
                    Height =315
                    TabIndex =10
                    BackColor =16777164
                    Name ="MHOU"
                    Format ="Fixed"
                    DefaultValue ="0"
                    ControlTipText ="ساعت"
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5163
                    Top =3705
                    Width =666
                    Height =315
                    TabIndex =6
                    BackColor =13434828
                    Name ="DA"
                    Format ="Fixed"
                    ValidationRule ="Not Is Null"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="1"
                    ControlTipText ="روز"
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    TabStop = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5163
                    Top =4085
                    Width =666
                    Height =315
                    TabIndex =9
                    BackColor =16777164
                    Name ="MDA"
                    Format ="Fixed"
                    DefaultValue ="0"
                    ControlTipText ="روز"
                End
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =2
                    Left =5298
                    Top =3346
                    Width =390
                    Height =300
                    Name ="Label20"
                    Caption ="روز"
                End
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =2
                    Left =5941
                    Top =3345
                    Width =510
                    Height =300
                    Name ="Label21"
                    Caption ="ساعت"
                End
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =2
                    Left =6661
                    Top =3345
                    Width =510
                    Height =300
                    Name ="Label22"
                    Caption ="دقيقه"
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1456
                    Top =2115
                    Width =5811
                    Height =315
                    TabIndex =3
                    Name ="MOLAH"
                    ControlSource ="MOLAH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =7368
                            Top =2115
                            Width =1185
                            Height =300
                            Name ="Label26"
                            Caption ="توضيحات"
                        End
                    End
                End
                Begin ComboBox
                    RowSourceTypeInt =1
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =5161
                    Top =1680
                    Width =2106
                    Height =315
                    TabIndex =2
                    Name ="KINDM"
                    ControlSource ="KINDM"
                    RowSourceType ="Value List"
                    RowSource ="0;\"استحقاقي\";1;\"استراحت پزشكي\";2;\"ساعتي\";3;\"بدون حقوق\""
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =7368
                            Top =1680
                            Width =1185
                            Height =300
                            Name ="Label24"
                            Caption ="نوع مرخصي"
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
                    Left =5146
                    Top =2490
                    Width =2106
                    Height =315
                    TabIndex =4
                    Name ="MOSTDATE"
                    ControlSource ="MOSTDATE"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="=FARSIDATE(Date())"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =7353
                            Top =2490
                            Width =1185
                            Height =300
                            Name ="Label28"
                            Caption ="تاريخ شروع"
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
                    Left =5146
                    Top =2880
                    Width =2106
                    Height =315
                    TabIndex =5
                    Name ="MOENDATE"
                    ControlSource ="MOENDATE"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="=FARSIDATE(Date())"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =7353
                            Top =2880
                            Width =1185
                            Height =300
                            Name ="Label30"
                            Caption ="تاريخ پايان"
                        End
                    End
                End
                Begin CheckBox
                    Locked = NotDefault
                    OverlapFlags =85
                    ReadingOrder =1
                    Left =9841
                    Top =1029
                    Width =340
                    Height =218
                    TabIndex =14
                    Name ="OKF"
                    ControlSource ="OKF"
                    DefaultValue ="False"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =10231
                            Top =975
                            Width =975
                            Height =240
                            ForeColor =255
                            Name ="Label146"
                            Caption ="تائيد "
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =1
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =7050
                    Top =5505
                    Width =1431
                    Height =345
                    FontSize =10
                    FontWeight =700
                    TabIndex =15
                    BackColor =13408767
                    Name ="UP_DATE"
                    ControlSource ="UP_DATE"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =8582
                            Top =5502
                            Width =900
                            Height =240
                            Name ="Label148"
                            Caption ="تاريخ اصلاح:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =1
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =946
                    Top =6127
                    Width =2841
                    Height =345
                    FontSize =10
                    FontWeight =700
                    TabIndex =16
                    BackColor =13408767
                    Name ="UP_TIME"
                    ControlSource ="UP_TIME"
                    Format ="General Date"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3903
                            Top =6129
                            Width =1050
                            Height =240
                            Name ="Label149"
                            Caption ="ساعت اصلاح:"
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
                    Left =991
                    Top =5547
                    Height =345
                    TabIndex =17
                    BackColor =13408767
                    Name ="UP_USER_NAME"
                    ControlSource ="UP_USER_NAME"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2743
                            Top =5547
                            Width =1290
                            Height =240
                            Name ="Label150"
                            Caption ="كاربر اصلاح كننده:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =1
                    TextFontFamily =18
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =9644
                    Top =5552
                    Width =1431
                    Height =345
                    FontSize =10
                    FontWeight =700
                    TabIndex =18
                    BackColor =13408767
                    Name ="PC_NAME"
                    ControlSource ="PC_NAME"
                    FontName ="Times New Roman"
                    InputMask ="##/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =11175
                            Top =5547
                            Width =1755
                            Height =240
                            Name ="Label152"
                            Caption ="نام كامپيوتر اصلاح كننده:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =1
                    TextFontFamily =18
                    ReadingOrder =2
                    KeyboardLanguage =11
                    ScrollBarAlign =2
                    IMESentenceMode =3
                    Left =4501
                    Top =5497
                    Width =1929
                    Height =345
                    FontSize =12
                    FontWeight =700
                    TabIndex =19
                    BackColor =13408767
                    Name ="IPADD"
                    ControlSource ="IPADD"
                    FontName ="Times New Roman"
                    InputMask ="##/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4152
                            Top =5549
                            Width =285
                            Height =240
                            Name ="Label155"
                            Caption ="IP:"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =6519
                    Top =458
                    Width =756
                    Height =315
                    TabIndex =20
                    BackColor =16777164
                    Name ="IDNUM"
                    ControlSource ="IDNUM"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =7374
                            Top =453
                            Width =855
                            Height =240
                            Name ="Label32"
                            Caption ="شماره برگه"
                        End
                    End
                End
            End
        End
        Begin FormFooter
            Height =0
            BackColor =-2147483633
            Name ="FormFooter"
        End
    End
End
CodeBehindForm
' See "TR_PMORAKH.cls"
