Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    DividingLines = NotDefault
    KeyPreview = NotDefault
    DefaultView =0
    AllowUpdating =4
    ScrollBars =2
    TabularCharSet =178
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =64
    GridY =64
    Width =14925
    RowHeight =285
    ItemSuffix =70
    Left =4170
    Top =630
    Right =18960
    Bottom =9810
    HelpContextId =43
    DatasheetForeColor =33554432
    DatasheetGridlinesColor =12632256
    MaxRecords =1000000
    RecSrcDt = Begin
        0x21d24861f2f0e540
    End
    RecordSource ="Q_STUF_DEF"
    Caption ="تعريف كالاي جديد"
    OnCurrent ="[Event Procedure]"
    OnDelete ="[Event Procedure]"
    BeforeUpdate ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    HelpFile ="dena.hlp"
    DatasheetFontName ="Tahoma"
    OnKeyPress ="[Event Procedure]"
    OnLoad ="[Event Procedure]"
    OnError ="[Event Procedure]"
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
        Begin Image
            BackStyle =0
            OldBorderStyle =0
            PictureAlignment =2
            Width =1701
            Height =1701
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
            Height =12153
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =11925
                    Top =135
                    Width =1605
                    Height =315
                    ColumnWidth =900
                    Name ="CODE"
                    ControlSource ="CODE"
                    BeforeUpdate ="[Event Procedure]"
                    AfterUpdate ="[Event Procedure]"
                    InputMask ="############"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextFontCharSet =0
                            Left =13693
                            Top =142
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
                    Left =10736
                    Top =1020
                    Width =2805
                    Height =315
                    ColumnWidth =2850
                    TabIndex =2
                    Name ="N_FANI"
                    ControlSource ="N_FANI"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextFontCharSet =0
                            Left =13695
                            Top =1020
                            Width =930
                            Height =285
                            Name ="Label5"
                            Caption ="شماره فني:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =95
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =8397
                    Top =1020
                    Width =1635
                    Height =315
                    TabIndex =14
                    Name ="TOZIH"
                    ControlSource ="TOZIH"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextFontCharSet =0
                            Left =10030
                            Top =1020
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
                    Left =5505
                    Top =1260
                    Width =1356
                    Height =315
                    ColumnWidth =1320
                    TabIndex =9
                    Name ="B_SEF"
                    ControlSource ="B_SEF"
                    Format ="#,##0.00\" ريال\";-#,##0.00\" ريال\""
                    OnKeyPress ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextFontCharSet =0
                            Left =7054
                            Top =1260
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
                    Left =5505
                    Top =570
                    Width =1356
                    Height =315
                    ColumnWidth =1185
                    TabIndex =7
                    Name ="N_SEF"
                    ControlSource ="N_SEF"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextFontCharSet =0
                            Left =7024
                            Top =570
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
                    Left =3106
                    Top =1530
                    Width =862
                    Height =315
                    TabIndex =13
                    Name ="MIN_M"
                    ControlSource ="MIN_M"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextFontCharSet =0
                            Left =4128
                            Top =1530
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
                    Left =5505
                    Top =1650
                    Width =1356
                    Height =315
                    ColumnWidth =1530
                    TabIndex =10
                    Name ="MAX_M"
                    ControlSource ="MAX_M"
                    Format ="#,##0.00\" ريال\";-#,##0.00\" ريال\""
                    OnKeyPress ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextFontCharSet =0
                            Left =7024
                            Top =1650
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
                    Left =5505
                    Top =915
                    Width =1356
                    Height =315
                    ColumnWidth =1365
                    TabIndex =8
                    Name ="MABL_F"
                    ControlSource ="MABL_F"
                    Format ="#,##0.00\" ريال\";-#,##0.00\" ريال\""
                    OnKeyPress ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextFontCharSet =0
                            Left =7069
                            Top =915
                            Width =1260
                            Height =300
                            Name ="Label23"
                            Caption ="في عمده فروش:"
                        End
                    End
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =10751
                    Top =1815
                    Width =2805
                    Height =315
                    ColumnWidth =1005
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
                            OverlapFlags =93
                            TextFontCharSet =0
                            Left =13708
                            Top =1815
                            Width =465
                            Height =300
                            Name ="Label9"
                            Caption ="واحد:"
                        End
                    End
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =8835
                    Top =2205
                    Width =4731
                    Height =315
                    ColumnWidth =1485
                    TabIndex =5
                    Name ="RADAH"
                    ControlSource ="RADAH"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TCOD_STUFGROUP.CODE, TCOD_STUFGROUP.NAMES FROM TCOD_STUFGROUP WHERE (((TC"
                        "OD_STUFGROUP.CODE)<>0)) ORDER BY TCOD_STUFGROUP.NAMES;"
                    ColumnWidths ="0"
                    OnDblClick ="[Event Procedure]"
                    OnNotInList ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextFontCharSet =0
                            Left =13704
                            Top =2205
                            Width =795
                            Height =300
                            Name ="Label19"
                            Caption ="گروه كالا:"
                        End
                    End
                End
                Begin Subform
                    OverlapFlags =85
                    Left =60
                    Top =3328
                    Width =14565
                    Height =2430
                    TabIndex =18
                    Name ="STUF_FSK_sub"
                    SourceObject ="Form.STUF_FSK_sub"
                    LinkChildFields ="CODE"
                    LinkMasterFields ="CODE"
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    TabStop = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =6180
                    Top =5818
                    Width =1626
                    Height =315
                    TabIndex =19
                    Name ="MABLK"
                    ControlSource ="=[STUF_FSK_sub].[Form]![SMABL_A]"
                    Format ="#,##0.00\" ريال\";-#,##0.00\" ريال\""
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    TabStop = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =7845
                    Top =5818
                    Width =1236
                    Height =315
                    TabIndex =20
                    Name ="Text28"
                    ControlSource ="=IIf([MOGUDIKOL]<>0,[MABLK]/[MOGUDIKOL],0)"
                    Format ="#,##0.00\" ريال\";-#,##0.00\" ريال\""
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =87
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =9090
                    Top =5818
                    Width =891
                    Height =315
                    TabIndex =21
                    Name ="MOGUDIKOL"
                    ControlSource ="=[STUF_FSK_sub].[Form]![SMOGODI_A]"
                End
                Begin Subform
                    OverlapFlags =85
                    Left =11685
                    Top =6494
                    Width =2971
                    Height =2670
                    TabIndex =22
                    Name ="MODULE_D_SUB"
                    SourceObject ="Form.MODULE_D_SUB"
                    LinkChildFields ="CODE"
                    LinkMasterFields ="CODE"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =13020
                            Top =6165
                            Width =990
                            Height =300
                            Name ="Label33"
                            Caption ="ساير واحد ها"
                        End
                    End
                End
                Begin Rectangle
                    OverlapFlags =255
                    Left =45
                    Top =60
                    Width =14598
                    Height =2964
                    Name ="Box34"
                End
                Begin CommandButton
                    OverlapFlags =247
                    TextFontCharSet =0
                    Left =2025
                    Top =120
                    Width =1845
                    Height =360
                    TabIndex =25
                    Name ="Command39"
                    Caption ="مرتب سازي براساس كد"
                    OnClick ="[Event Procedure]"
                End
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =0
                    Left =5572
                    Top =3045
                    Width =2925
                    Height =240
                    Name ="Label40"
                    Caption ="موجودي ابتداي دوره و انبارهاي  حاوي كالا"
                End
                Begin Subform
                    OverlapFlags =85
                    Left =15
                    Top =6493
                    Width =7395
                    Height =2670
                    TabIndex =23
                    Name ="TAKHPERS_SUB"
                    SourceObject ="Form.TAKHPERS_SUB2"
                    LinkChildFields ="TAKH_COD"
                    LinkMasterFields ="CODE"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =2890
                            Top =6107
                            Width =1185
                            Height =300
                            Name ="TAKHPERS_SUB Label"
                            Caption ="قيمت  مصوب"
                            OnDblClick ="[Event Procedure]"
                            EventProcPrefix ="TAKHPERS_SUB_Label"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    ColumnHidden = NotDefault
                    OverlapFlags =247
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =8950
                    Top =1530
                    Width =81
                    Height =255
                    TabIndex =24
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
                    Left =8295
                    Top =586
                    Width =5250
                    Height =315
                    ColumnWidth =5685
                    TabIndex =1
                    BackColor =13434828
                    Name ="NAM"
                    ControlSource ="NAM"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextFontCharSet =0
                            Left =13693
                            Top =586
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
                    Left =8175
                    Top =165
                    Width =283
                    Height =245
                    ColumnWidth =1590
                    TabIndex =6
                    Name ="CMBAA"
                    ControlSource ="CMBAA"
                    Begin
                        Begin Label
                            OverlapFlags =255
                            Left =8448
                            Top =142
                            Width =1485
                            Height =240
                            ForeColor =255
                            Name ="Label48"
                            Caption ="مشمول ماليات ب.ا.ا"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =247
                    Left =120
                    Top =120
                    Width =1871
                    Height =360
                    TabIndex =26
                    Name ="Command50"
                    Caption ="كدينگ شماره فني(باركد)"
                    OnClick ="[Event Procedure]"
                End
                Begin TextBox
                    OverlapFlags =247
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3106
                    Top =1020
                    Width =862
                    Height =315
                    TabIndex =12
                    Name ="VAZN"
                    ControlSource ="VAZN"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextFontCharSet =0
                            Left =4113
                            Top =1020
                            Width =735
                            Height =240
                            Name ="Label49"
                            Caption ="وزن واحد:"
                        End
                    End
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =247
                    TextFontCharSet =0
                    Left =3900
                    Top =120
                    Width =1845
                    Height =360
                    TabIndex =27
                    Name ="ESLAH"
                    Caption ="اصلاح "
                    OnClick ="[Event Procedure]"
                End
                Begin CheckBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =247
                    ReadingOrder =1
                    Left =9975
                    Top =174
                    Width =340
                    Height =218
                    ColumnWidth =510
                    TabIndex =28
                    Name ="OKF"
                    ControlSource ="OKF"
                    DefaultValue ="False"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextFontCharSet =0
                            Left =10365
                            Top =120
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
                    Left =165
                    Top =598
                    Width =3786
                    Height =315
                    TabIndex =11
                    Name ="MENUIT"
                    ControlSource ="MENUIT"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="TCODE_MENUITEM"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextFontCharSet =0
                            Left =4144
                            Top =570
                            Width =1470
                            Height =330
                            Name ="Label52"
                            Caption ="زيرمجموعه منوي:"
                        End
                    End
                End
                Begin Image
                    SizeMode =3
                    PictureType =1
                    Left =7515
                    Top =6510
                    Width =4106
                    Height =2639
                    Name ="PIC"
                    ObjectPalette = Begin
                        0x000300010000000020000000400000006000000080000000a0000000c0000000 ,
                        0xe00000000020000020200000402000006020000080200000a0200000c0200000 ,
                        0xe02000000040000020400000404000006040000080400000a0400000c0400000 ,
                        0xe04000000060000020600000406000006060000080600000a0600000c0600000 ,
                        0xe06000000080000020800000408000006080000080800000a0800000c0800000 ,
                        0xe080000000a0000020a0000040a0000060a0000080a00000a0a00000c0a00000 ,
                        0xe0a0000000c0000020c0000040c0000060c0000080c00000a0c00000c0c00000 ,
                        0xe0c0000000e0000020e0000040e0000060e0000080e00000a0e00000c0e00000 ,
                        0xe0e000000000400020004000400040006000400080004000a0004000c0004000 ,
                        0xe00040000020400020204000402040006020400080204000a0204000c0204000 ,
                        0xe02040000040400020404000404040006040400080404000a0404000c0404000 ,
                        0xe04040000060400020604000406040006060400080604000a0604000c0604000 ,
                        0xe06040000080400020804000408040006080400080804000a0804000c0804000 ,
                        0xe080400000a0400020a0400040a0400060a0400080a04000a0a04000c0a04000 ,
                        0xe0a0400000c0400020c0400040c0400060c0400080c04000a0c04000c0c04000 ,
                        0xe0c0400000e0400020e0400040e0400060e0400080e04000a0e04000c0e04000 ,
                        0xe0e040000000800020008000400080006000800080008000a0008000c0008000 ,
                        0xe00080000020800020208000402080006020800080208000a0208000c0208000 ,
                        0xe02080000040800020408000404080006040800080408000a0408000c0408000 ,
                        0xe04080000060800020608000406080006060800080608000a0608000c0608000 ,
                        0xe06080000080800020808000408080006080800080808000a0808000c0808000 ,
                        0xe080800000a0800020a0800040a0800060a0800080a08000a0a08000c0a08000 ,
                        0xe0a0800000c0800020c0800040c0800060c0800080c08000a0c08000c0c08000 ,
                        0xe0c0800000e0800020e0800040e0800060e0800080e08000a0e08000c0e08000 ,
                        0xe0e080000000c0002000c0004000c0006000c0008000c000a000c000c000c000 ,
                        0xe000c0000020c0002020c0004020c0006020c0008020c000a020c000c020c000 ,
                        0xe020c0000040c0002040c0004040c0006040c0008040c000a040c000c040c000 ,
                        0xe040c0000060c0002060c0004060c0006060c0008060c000a060c000c060c000 ,
                        0xe060c0000080c0002080c0004080c0006080c0008080c000a080c000c080c000 ,
                        0xe080c00000a0c00020a0c00040a0c00060a0c00080a0c000a0a0c000c0a0c000 ,
                        0xe0a0c00000c0c00020c0c00040c0c00060c0c00080c0c000a0c0c000c0c0c000 ,
                        0xe0c0c00000e0c00020e0c00040e0c00060e0c00080e0c000a0e0c000c0e0c000 ,
                        0xe0e0c00000000000
                    End
                    Picture ="C:\\software\\prg\\tink.gif"
                End
                Begin OptionGroup
                    Visible = NotDefault
                    BackStyle =1
                    OverlapFlags =255
                    Left =165
                    Top =1020
                    Width =2721
                    Height =901
                    TabIndex =15
                    BackColor =10079487
                    Name ="Frame58"
                    Begin
                        Begin Label
                            OverlapFlags =255
                            TextAlign =2
                            Left =2181
                            Top =1297
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
                    Visible = NotDefault
                    OverlapFlags =247
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =285
                    Top =1513
                    Width =846
                    Height =315
                    TabIndex =17
                    ForeColor =16711680
                    Name ="MEGHJAY"
                    ControlSource ="MEGHJAY"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextFontCharSet =0
                            Left =1203
                            Top =1513
                            Width =1275
                            Height =300
                            ForeColor =16711680
                            Name ="Label55"
                            Caption ="جايزه:"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =247
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =285
                    Top =1155
                    Width =846
                    Height =315
                    TabIndex =16
                    ForeColor =16711680
                    Name ="MEGHTA"
                    ControlSource ="MEGHTA"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextFontCharSet =0
                            Left =1198
                            Top =1155
                            Width =945
                            Height =240
                            ForeColor =16711680
                            Name ="Label57"
                            Caption ="به ازاي هر:"
                        End
                    End
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    Left =132
                    Top =5993
                    Width =1845
                    Height =360
                    TabIndex =29
                    Name ="Command55"
                    Caption ="ليست تخفيفات "
                    OnClick ="[Event Procedure]"
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
                    Left =165
                    Top =2113
                    Width =6906
                    Height =315
                    TabIndex =30
                    BackColor =10092543
                    Name ="PGID"
                    ControlSource ="PGID"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT PGID, PGNAME FROM PRICE_GRP"
                    ColumnWidths ="0"
                    OnNotInList ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextFontCharSet =0
                            Left =7264
                            Top =2085
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
                    Left =10736
                    Top =1425
                    Width =2805
                    Height =315
                    TabIndex =3
                    Name ="Barcode"
                    ControlSource ="Barcode"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextFontCharSet =0
                            Left =13695
                            Top =1425
                            Width =930
                            Height =285
                            Name ="Label60"
                            Caption ="بارکد:"
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
                    Left =9405
                    Top =2581
                    Width =2805
                    Height =315
                    TabIndex =31
                    Name ="sstid"
                    ControlSource ="sstid"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextFontCharSet =0
                            Left =12360
                            Top =2580
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
                    Left =4380
                    Top =2581
                    Width =2805
                    Height =315
                    TabIndex =32
                    Name ="mu"
                    ControlSource ="mu"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT IDD, NAME_MO FROM TCOD_VAHED_EXTENDED ORDER BY NAME_MO"
                    ColumnWidths ="0"
                    OnDblClick ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextFontCharSet =0
                            Left =7337
                            Top =2581
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
                    Left =1965
                    Top =2581
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
                            Left =2835
                            Top =2581
                            Width =1365
                            Height =240
                            Name ="Label66"
                            Caption ="درصد ارزش افزوده:"
                        End
                    End
                End
                Begin Subform
                    OverlapFlags =85
                    Left =60
                    Top =9538
                    Width =14565
                    Height =2430
                    TabIndex =34
                    Name ="FRM_RewardRules_Subform"
                    SourceObject ="Form.FRM_RewardRules_Subform"
                    LinkChildFields ="ProductID_Target"
                    LinkMasterFields ="CODE"
                End
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =2
                    Left =5542
                    Top =9255
                    Width =2925
                    Height =240
                    Name ="Label68"
                    Caption ="تعريف جايزه پلکاني"
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =247
                    IMESentenceMode =3
                    Left =5952
                    Top =194
                    TabIndex =35
                    Name ="UID"
                    ControlSource ="UID"
                    DefaultValue ="=UCurrentUserIDD()"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =8949
                            Top =194
                            Width =405
                            Height =240
                            Name ="Label69"
                            Caption ="UID:"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "STUF_DEF.cls"
