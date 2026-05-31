Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    OrderByOn = NotDefault
    AllowEdits = NotDefault
    DefaultView =0
    AllowUpdating =4
    ScrollBars =0
    ViewsAllowed =1
    TabularCharSet =178
    TabularFamily =42
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =14583
    DatasheetFontHeight =10
    ItemSuffix =48
    Left =330
    Right =14415
    Bottom =7875
    HelpContextId =43
    DatasheetGridlinesColor =12632256
    MaxRecords =10000000
    RecSrcDt = Begin
        0xe258176f2ef7e340
    End
    RecordSource ="TR_STUF_DEF"
    Caption ="سابقه تعريف كالاي جديد"
    HelpFile ="dena.hlp"
    DatasheetFontName ="Arial (Arabic)"
    AllowDatasheetView =0
    RecordSourceQualifier ="dbo"
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
        Begin Section
            CanGrow = NotDefault
            Height =10475
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin Subform
                    OverlapFlags =85
                    Left =278
                    Top =3504
                    Width =14055
                    Height =2325
                    Name ="STUF_FSK_sub"
                    SourceObject ="Form.TR_STUF_FSK_sub"
                    LinkChildFields ="CODE;UP_DATE;UP_TIME"
                    LinkMasterFields ="CODE;UP_DATE;UP_TIME"
                End
                Begin TextBox
                    TabStop = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5329
                    Top =5885
                    Width =1626
                    Height =315
                    TabIndex =1
                    Name ="MABLK"
                    ControlSource ="=[STUF_FSK_sub].[Form]![SMABL_A]"
                    Format ="#,##0.00\" ريال\";-#,##0.00\" ريال\""
                End
                Begin TextBox
                    TabStop = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =6994
                    Top =5885
                    Width =1236
                    Height =315
                    TabIndex =2
                    Name ="Text28"
                    ControlSource ="=IIf([MOGUDIKOL]<>0,[MABLK]/[MOGUDIKOL],0)"
                    Format ="#,##0.00\" ريال\";-#,##0.00\" ريال\""
                End
                Begin TextBox
                    TabStop = NotDefault
                    OverlapFlags =87
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =8239
                    Top =5885
                    Width =891
                    Height =315
                    TabIndex =3
                    Name ="MOGUDIKOL"
                    ControlSource ="=[STUF_FSK_sub].[Form]![SMOGODI_A]"
                End
                Begin Subform
                    OverlapFlags =215
                    Left =10089
                    Top =6251
                    Width =4138
                    Height =2499
                    TabIndex =4
                    Name ="Child32"
                    SourceObject ="Form.TR_MODULE_D_SUB"
                    LinkChildFields ="CODE;UP_DATE;UP_TIME"
                    LinkMasterFields ="CODE;UP_DATE;UP_TIME"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextFontCharSet =0
                            Left =11683
                            Top =5967
                            Width =990
                            Height =300
                            Name ="Label33"
                            Caption ="ساير واحد ها"
                        End
                    End
                End
                Begin Rectangle
                    OverlapFlags =223
                    Left =10034
                    Top =5952
                    Width =4218
                    Height =2949
                    Name ="Box35"
                End
                Begin Label
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =2
                    Left =4764
                    Top =3174
                    Width =4080
                    Height =240
                    Name ="Label40"
                    Caption ="موجودي ابتداي دوره و انبارهاي  حاوي كالا"
                End
                Begin Subform
                    OverlapFlags =87
                    Left =375
                    Top =6382
                    Width =4215
                    Height =2325
                    TabIndex =5
                    Name ="TAKHPERS_SUB"
                    SourceObject ="Form.TR_TAKHPERS_SUB2"
                    LinkChildFields ="TAKH_COD;UP_DATE;UP_TIME"
                    LinkMasterFields ="CODE;UP_DATE;UP_TIME"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextFontCharSet =0
                            Left =1821
                            Top =6060
                            Width =1185
                            Height =300
                            Name ="TAKHPERS_SUB Label"
                            Caption ="تخفيفات مصوب"
                            EventProcPrefix ="TAKHPERS_SUB_Label"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =1
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =420
                    Top =9018
                    Width =1431
                    Height =345
                    FontSize =10
                    FontWeight =700
                    TabIndex =6
                    BackColor =13408767
                    Name ="UP_DATE"
                    ControlSource ="UP_DATE"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1952
                            Top =9015
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
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =437
                    Top =9465
                    Width =2841
                    Height =345
                    FontSize =10
                    FontWeight =700
                    TabIndex =7
                    BackColor =13408767
                    Name ="UP_TIME"
                    ControlSource ="UP_TIME"
                    Format ="General Date"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3394
                            Top =9467
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
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =420
                    Top =9905
                    Height =345
                    TabIndex =8
                    BackColor =13408767
                    Name ="UP_USER_NAME"
                    ControlSource ="UP_USER_NAME"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2172
                            Top =9905
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
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =3014
                    Top =9065
                    Width =1431
                    Height =345
                    FontSize =10
                    FontWeight =700
                    TabIndex =9
                    BackColor =13408767
                    Name ="PC_NAME"
                    ControlSource ="PC_NAME"
                    FontName ="Times New Roman"
                    InputMask ="##/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4545
                            Top =9060
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
                    ScrollBarAlign =2
                    IMESentenceMode =3
                    Left =3930
                    Top =9855
                    Width =1929
                    Height =345
                    FontSize =12
                    FontWeight =700
                    TabIndex =10
                    BackColor =13408767
                    Name ="IPADD"
                    ControlSource ="IPADD"
                    FontName ="Times New Roman"
                    InputMask ="##/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3581
                            Top =9907
                            Width =285
                            Height =240
                            Name ="Label155"
                            Caption ="IP:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =11880
                    Top =360
                    Width =1605
                    Height =315
                    TabIndex =11
                    Name ="CODE"
                    ControlSource ="CODE"
                    InputMask ="############"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextFontCharSet =0
                            Left =13648
                            Top =367
                            Width =600
                            Height =300
                            Name ="Label1"
                            Caption ="كد كالا:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =10691
                    Top =1245
                    Width =2805
                    Height =315
                    TabIndex =12
                    Name ="N_FANI"
                    ControlSource ="N_FANI"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextFontCharSet =0
                            Left =13650
                            Top =1245
                            Width =930
                            Height =285
                            Name ="Label5"
                            Caption ="شماره فني:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =8352
                    Top =1245
                    Width =1470
                    Height =315
                    TabIndex =13
                    Name ="TOZIH"
                    ControlSource ="TOZIH"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextFontCharSet =0
                            Left =9985
                            Top =1245
                            Width =615
                            Height =300
                            Name ="Label7"
                            Caption ="توضيح:"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5460
                    Top =1485
                    Width =1356
                    Height =315
                    TabIndex =14
                    Name ="B_SEF"
                    ControlSource ="B_SEF"
                    Format ="#,##0.00\" ريال\";-#,##0.00\" ريال\""
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextFontCharSet =0
                            Left =7009
                            Top =1485
                            Width =1215
                            Height =300
                            Name ="Label11"
                            Caption ="في خرده فروش:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5460
                    Top =795
                    Width =1356
                    Height =315
                    TabIndex =15
                    Name ="N_SEF"
                    ControlSource ="N_SEF"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextFontCharSet =0
                            Left =6979
                            Top =795
                            Width =1080
                            Height =300
                            Name ="Label13"
                            Caption ="نقطه سفارش:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3061
                    Top =1755
                    Width =862
                    Height =315
                    TabIndex =16
                    Name ="MIN_M"
                    ControlSource ="MIN_M"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextFontCharSet =0
                            Left =4083
                            Top =1755
                            Width =1185
                            Height =300
                            Name ="Label15"
                            Caption ="حداقل موجودي:"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5460
                    Top =1875
                    Width =1356
                    Height =315
                    TabIndex =17
                    Name ="MAX_M"
                    ControlSource ="MAX_M"
                    Format ="#,##0.00\" ريال\";-#,##0.00\" ريال\""
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextFontCharSet =0
                            Left =6979
                            Top =1875
                            Width =1425
                            Height =240
                            Name ="Label17"
                            Caption ="قيمت مصرف كننده:"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5460
                    Top =1140
                    Width =1356
                    Height =315
                    TabIndex =18
                    Name ="MABL_F"
                    ControlSource ="MABL_F"
                    Format ="#,##0.00\" ريال\";-#,##0.00\" ريال\""
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextFontCharSet =0
                            Left =7024
                            Top =1140
                            Width =1260
                            Height =300
                            Name ="Label23"
                            Caption ="في عمده فروش:"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =10691
                    Top =1755
                    Width =2805
                    Height =315
                    TabIndex =19
                    Name ="VAHED"
                    ControlSource ="VAHED"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TCOD_VAHEDS.CODE, TCOD_VAHEDS.NAMES FROM TCOD_VAHEDS ORDER BY TCOD_VAHEDS"
                        ".NAMES;"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextFontCharSet =0
                            Left =13648
                            Top =1755
                            Width =465
                            Height =300
                            Name ="Label9"
                            Caption ="واحد:"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =8779
                    Top =2322
                    Width =4701
                    Height =315
                    TabIndex =20
                    Name ="RADAH"
                    ControlSource ="RADAH"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TCOD_STUFGROUP.CODE, TCOD_STUFGROUP.NAMES FROM TCOD_STUFGROUP WHERE (((TC"
                        "OD_STUFGROUP.CODE)<>0)) ORDER BY TCOD_STUFGROUP.NAMES;"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextFontCharSet =0
                            Left =13648
                            Top =2322
                            Width =795
                            Height =300
                            Name ="Label19"
                            Caption ="گروه كالا:"
                        End
                    End
                End
                Begin Rectangle
                    OverlapFlags =255
                    Top =285
                    Width =14583
                    Height =2949
                    Name ="Box34"
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =247
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =8905
                    Top =1755
                    Width =81
                    Height =255
                    TabIndex =21
                    Name ="KINDK"
                    ControlSource ="KINDK"
                    DefaultValue ="1"
                End
                Begin TextBox
                    OverlapFlags =247
                    TextFontCharSet =0
                    TextAlign =3
                    ScrollBarAlign =2
                    NumeralShapes =3
                    Left =8250
                    Top =811
                    Width =5250
                    Height =315
                    TabIndex =22
                    BackColor =13434828
                    Name ="NAME"
                    ControlSource ="NAME"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextFontCharSet =0
                            Left =13648
                            Top =811
                            Width =660
                            Height =300
                            Name ="Label3"
                            Caption ="نام كالا:"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =247
                    ReadingOrder =1
                    Left =8130
                    Top =390
                    Width =283
                    Height =245
                    TabIndex =23
                    Name ="CMBAA"
                    ControlSource ="CMBAA"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =8403
                            Top =367
                            Width =1485
                            Height =240
                            ForeColor =255
                            Name ="Label48"
                            Caption ="مشمول ماليات ب.ا.ا"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =247
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3061
                    Top =1245
                    Width =862
                    Height =315
                    TabIndex =24
                    Name ="VAZN"
                    ControlSource ="VAZN"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextFontCharSet =0
                            Left =4068
                            Top =1245
                            Width =735
                            Height =240
                            Name ="Label49"
                            Caption ="وزن واحد:"
                        End
                    End
                End
                Begin CheckBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =247
                    ReadingOrder =1
                    Left =9930
                    Top =399
                    Width =340
                    Height =218
                    TabIndex =25
                    Name ="OKF"
                    ControlSource ="OKF"
                    DefaultValue ="False"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextFontCharSet =0
                            Left =10320
                            Top =345
                            Width =405
                            Height =270
                            ForeColor =255
                            Name ="Label146"
                            Caption ="تائيد "
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =247
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =120
                    Top =823
                    Width =3786
                    Height =315
                    TabIndex =26
                    Name ="MENUIT"
                    ControlSource ="MENUIT"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="TCODE_MENUITEM"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextFontCharSet =0
                            Left =4099
                            Top =795
                            Width =1470
                            Height =330
                            Name ="Label52"
                            Caption ="زيرمجموعه منوي:"
                        End
                    End
                End
                Begin OptionGroup
                    BackStyle =1
                    OverlapFlags =255
                    Left =120
                    Top =1245
                    Width =2721
                    Height =901
                    TabIndex =27
                    BackColor =10079487
                    Name ="Frame58"
                    Begin
                        Begin Label
                            OverlapFlags =255
                            TextAlign =2
                            Left =2136
                            Top =1522
                            Width =705
                            Height =240
                            BackColor =10079487
                            ForeColor =128
                            Name ="Label54"
                            Caption ="جايزه"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =247
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =240
                    Top =1738
                    Width =846
                    Height =315
                    TabIndex =28
                    ForeColor =16711680
                    Name ="MEGHJAY"
                    ControlSource ="MEGHJAY"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextFontCharSet =0
                            Left =1158
                            Top =1738
                            Width =1275
                            Height =300
                            ForeColor =16711680
                            Name ="Label55"
                            Caption ="جايزه:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =247
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =240
                    Top =1380
                    Width =846
                    Height =315
                    TabIndex =29
                    ForeColor =16711680
                    Name ="MEGHTA"
                    ControlSource ="MEGHTA"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextFontCharSet =0
                            Left =1153
                            Top =1380
                            Width =945
                            Height =240
                            ForeColor =16711680
                            Name ="Label57"
                            Caption ="به ازاي هر:"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =247
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =120
                    Top =2338
                    Width =6906
                    Height =315
                    TabIndex =30
                    BackColor =10092543
                    Name ="PGID"
                    ControlSource ="PGID"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT PGID, PGNAME FROM PRICE_GRP"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextFontCharSet =0
                            Left =7219
                            Top =2310
                            Width =1470
                            Height =330
                            Name ="Label58"
                            Caption ="گروه قيمت گذاري:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =247
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =8743
                    Top =2778
                    Width =2805
                    Height =315
                    TabIndex =31
                    Name ="sstid"
                    ControlSource ="sstid"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextFontCharSet =0
                            Left =11698
                            Top =2777
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
                    OverlapFlags =247
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =3718
                    Top =2778
                    Width =2805
                    Height =315
                    TabIndex =32
                    Name ="mu"
                    ControlSource ="mu"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT IDD, NAME_MO FROM TCOD_VAHED_EXTENDED ORDER BY NAME_MO"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextFontCharSet =0
                            Left =6675
                            Top =2778
                            Width =1905
                            Height =240
                            Name ="Label64"
                            Caption ="واحد کالا در سامانه موديان:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =247
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1303
                    Top =2778
                    Width =720
                    Height =315
                    TabIndex =33
                    BackColor =13408767
                    Name ="vra"
                    ControlSource ="vra"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextFontCharSet =0
                            Left =2173
                            Top =2778
                            Width =1365
                            Height =240
                            Name ="Label66"
                            Caption ="درصد ارزش افزوده:"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "TR_STUF_DEF.cls"
