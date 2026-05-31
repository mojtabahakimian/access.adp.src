Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    DividingLines = NotDefault
    OrderByOn = NotDefault
    DefaultView =0
    AllowUpdating =4
    ScrollBars =0
    TabularCharSet =178
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =12450
    RowHeight =330
    ItemSuffix =52
    Left =2115
    Top =1515
    Right =14280
    Bottom =5385
    HelpContextId =43
    DatasheetForeColor =33554432
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x275089017028e640
    End
    RecordSource ="SELECT CAST(CODE AS int) AS nc, CODE, NAME AS NAM, N_FANI, TOZIH, VAHED, RADAH, "
        "KINDK, MABL_F, DEPART, CMBAA, mu, sstid, vra FROM STUF_DEF WHERE (KINDK = 2) ORD"
        "ER BY NAME"
    Caption ="خدمات"
    OnDelete ="[Event Procedure]"
    BeforeUpdate ="[Event Procedure]"
    AfterUpdate ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    HelpFile ="dena.hlp"
    DatasheetFontName ="Tahoma"
    OnLoad ="[Event Procedure]"
    OnError ="[Event Procedure]"
    Begin
        Begin Label
            BackStyle =0
            TextFontCharSet =178
            FontName ="Tahoma"
        End
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
            Width =850
            Height =850
        End
        Begin Line
            Width =1701
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
        Begin UnboundObjectFrame
            SpecialEffect =2
            OldBorderStyle =1
            Width =4536
            Height =2835
        End
        Begin CustomControl
            SpecialEffect =2
            Width =4536
            Height =2835
        End
        Begin FormHeader
            CanGrow = NotDefault
            SpecialEffect =1
            Height =585
            BackColor =8421376
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =2
                    Left =6700
                    Width =1860
                    Height =510
                    FontSize =16
                    FontWeight =700
                    ForeColor =13434828
                    Name ="Label47"
                    Caption ="تعريف خدمات"
                    FontName ="Titr"
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            SpecialEffect =2
            Height =4705
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
                    Left =9276
                    Top =124
                    Width =1605
                    Height =315
                    ColumnOrder =0
                    Name ="CODE"
                    ControlSource ="CODE"
                    AfterUpdate ="[Event Procedure]"
                    InputMask ="############"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =11044
                            Top =131
                            Width =600
                            Height =300
                            Name ="Label1"
                            Caption ="كد :"
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
                    Left =8497
                    Top =844
                    Width =2385
                    Height =315
                    ColumnOrder =7
                    TabIndex =2
                    Name ="N_FANI"
                    ControlSource ="N_FANI"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =11045
                            Top =844
                            Width =915
                            Height =300
                            Name ="Label5"
                            Caption ="شماره فني:"
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
                    Left =7252
                    Top =1189
                    Width =3630
                    Height =315
                    ColumnOrder =6
                    TabIndex =3
                    Name ="TOZIH"
                    ControlSource ="TOZIH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =11045
                            Top =1189
                            Width =615
                            Height =300
                            Name ="Label7"
                            Caption ="توضيح:"
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
                    Left =9037
                    Top =2254
                    Width =1851
                    Height =315
                    ColumnOrder =4
                    TabIndex =6
                    Name ="MABL_F"
                    ControlSource ="MABL_F"
                    Format ="#,##0.00\" ريال\";-#,##0.00\" ريال\""
                    OnKeyPress ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =11050
                            Top =2250
                            Width =945
                            Height =300
                            Name ="Label23"
                            Caption ="قيمت واحد :"
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
                    Left =9014
                    Top =1549
                    Width =1866
                    Height =315
                    ColumnOrder =3
                    TabIndex =4
                    Name ="VAHED"
                    ControlSource ="VAHED"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TCOD_VAHEDS.CODE, TCOD_VAHEDS.NAMES FROM TCOD_VAHEDS ORDER BY TCOD_VAHEDS"
                        ".NAMES;"
                    ColumnWidths ="0"
                    OnDblClick ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =11043
                            Top =1549
                            Width =465
                            Height =300
                            Name ="Label9"
                            Caption ="واحد:"
                        End
                    End
                End
                Begin ComboBox
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =9018
                    Top =1889
                    Width =1866
                    Height =315
                    ColumnOrder =5
                    TabIndex =5
                    Name ="RADAH"
                    ControlSource ="RADAH"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TCOD_STUFGROUP.CODE, TCOD_STUFGROUP.NAMES FROM TCOD_STUFGROUP WHERE (((TC"
                        "OD_STUFGROUP.CODE)=0)) ORDER BY TCOD_STUFGROUP.NAMES;"
                    ColumnWidths ="0"
                    DefaultValue ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =11047
                            Top =1889
                            Width =795
                            Height =300
                            Name ="Label19"
                            Caption ="گروه:"
                        End
                    End
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    Left =7536
                    Top =64
                    Width =1695
                    Height =360
                    TabIndex =11
                    Name ="Command39"
                    Caption ="مرتب سازي براساس كد"
                    OnClick ="[Event Procedure]"
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =7784
                    Top =1562
                    Width =1131
                    Height =255
                    ColumnWidth =1417
                    ColumnOrder =8
                    TabIndex =9
                    Name ="KINDK"
                    ControlSource ="KINDK"
                    DefaultValue ="2"
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =7222
                    Top =484
                    Width =3660
                    Height =315
                    ColumnOrder =1
                    TabIndex =1
                    Name ="NAM"
                    ControlSource ="NAM"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =11045
                            Top =484
                            Width =660
                            Height =300
                            Name ="Label3"
                            Caption ="نام :"
                        End
                    End
                End
                Begin Subform
                    OverlapFlags =87
                    Left =226
                    Top =300
                    Width =6525
                    Height =2610
                    TabIndex =10
                    Name ="TAKHPERS_SUB"
                    SourceObject ="Form.TAKHPERS_SUB"
                    LinkChildFields ="TAKH_COD"
                    LinkMasterFields ="CODE"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextFontCharSet =0
                            Left =1753
                            Width =1185
                            Height =300
                            Name ="TAKHPERS_SUB Label"
                            Caption ="تخفيفات مصوب"
                            EventProcPrefix ="TAKHPERS_SUB_Label"
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
                    Left =7222
                    Top =2625
                    Width =3666
                    Height =315
                    ColumnOrder =2
                    TabIndex =7
                    Name ="DEPART"
                    ControlSource ="DEPART"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT DEPART.DEPATMAN, DEPART.DEPNAME FROM DEPART ORDER BY DEPART.DEPNAME;"
                    ColumnWidths ="0"
                    ValidationRule ="Is Not Null"
                    DefaultValue ="1"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =11047
                            Top =2625
                            Width =1290
                            Height =300
                            Name ="Label51"
                            Caption ="واحد ارائه كننده:"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    ReadingOrder =1
                    Left =10662
                    Top =3049
                    Width =223
                    Height =230
                    TabIndex =8
                    Name ="CMBAA"
                    ControlSource ="CMBAA"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =10935
                            Top =3011
                            Width =1485
                            Height =240
                            ForeColor =255
                            Name ="Label48"
                            Caption ="مشمول ماليات ب.ا.ا"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =7440
                    Top =3402
                    Width =2805
                    Height =315
                    TabIndex =12
                    Name ="sstid"
                    ControlSource ="sstid"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =10395
                            Top =3401
                            Width =2055
                            Height =240
                            Name ="Label62"
                            Caption ="شناسه کالا در سامانه موديان"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =7426
                    Top =3855
                    Width =2805
                    Height =315
                    TabIndex =13
                    Name ="mu"
                    ControlSource ="mu"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT IDD, NAME_MO FROM TCOD_VAHED_EXTENDED ORDER BY NAME_MO"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =10383
                            Top =3855
                            Width =1905
                            Height =240
                            Name ="Label64"
                            Caption ="واحد کالا در سامانه موديان:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =9524
                    Top =4251
                    Width =720
                    Height =315
                    TabIndex =14
                    BackColor =13408767
                    Name ="vra"
                    ControlSource ="vra"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =10394
                            Top =4251
                            Width =1365
                            Height =240
                            Name ="Label66"
                            Caption ="درصد ارزش افزوده:"
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
' See "KHAD_DEF.cls"
