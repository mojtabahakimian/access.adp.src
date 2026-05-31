Version =20
VersionRequired =20
Begin Form
    KeyPreview = NotDefault
    DefaultView =2
    AllowUpdating =4
    ViewsAllowed =2
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =4380
    RowHeight =315
    ItemSuffix =46
    Left =405
    Top =2445
    Right =14850
    Bottom =7440
    HelpContextId =6003
    DatasheetGridlinesColor =12632256
    BeforeDelConfirm ="[Event Procedure]"
    RecSrcDt = Begin
        0x08977ba6f1a7e240
    End
    RecordSource ="INVO_LST"
    Caption ="INVO_LST"
    OnCurrent ="[Event Procedure]"
    OnDelete ="[Event Procedure]"
    BeforeUpdate ="[Event Procedure]"
    AfterUpdate ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    HelpFile ="dena.hlp"
    DatasheetFontName ="Tahoma"
    OnKeyPress ="[Event Procedure]"
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
            Height =0
            BackColor =-2147483633
            Name ="FormHeader"
        End
        Begin Section
            Height =8510
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =93
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =542
                    Top =113
                    Height =255
                    ColumnWidth =1417
                    ColumnOrder =0
                    Name ="NUMBER"
                    ControlSource ="NUMBER"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =3154
                            Top =113
                            Width =960
                            Height =240
                            Name ="Label1"
                            Caption ="شماره فاكتور"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =93
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =542
                    Top =453
                    Height =255
                    ColumnWidth =1417
                    ColumnOrder =1
                    TabIndex =1
                    Name ="TAG"
                    ControlSource ="TAG"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3484
                            Top =453
                            Width =630
                            Height =240
                            Name ="Label3"
                            Caption ="برچسب"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =542
                    Top =1814
                    Height =255
                    ColumnWidth =705
                    ColumnOrder =6
                    TabIndex =4
                    Name ="MEGH"
                    ControlSource ="MEGH"
                    AfterUpdate ="[Event Procedure]"
                    OnKeyPress ="[Event Procedure]"
                    Tag ="100"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3439
                            Top =1814
                            Width =675
                            Height =240
                            Name ="Label11"
                            Caption ="مقداركالا"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =568
                    Top =2437
                    Height =255
                    ColumnOrder =7
                    TabIndex =5
                    Name ="MEGH_MAR"
                    ControlSource ="MEGH_MAR"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3030
                            Top =2437
                            Width =1110
                            Height =240
                            Name ="Label13"
                            Caption ="مقدار مرجوعي"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =568
                    Top =3117
                    Height =255
                    ColumnWidth =1320
                    ColumnOrder =9
                    TabIndex =6
                    Name ="MABL"
                    ControlSource ="MABL"
                    Format ="#,##0.00;-#,##0.00"
                    ValidationRule ="Not Is Null"
                    AfterUpdate ="[Event Procedure]"
                    OnKeyPress ="[Event Procedure]"
                    Tag ="100"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3765
                            Top =3117
                            Width =375
                            Height =240
                            Name ="Label17"
                            Caption ="مبلغ"
                        End
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =568
                    Top =3457
                    Height =255
                    ColumnWidth =1140
                    ColumnOrder =10
                    TabIndex =7
                    Name ="MABL_K"
                    ControlSource ="MABL_K"
                    Format ="#,##0.00;-#,##0.00"
                    ValidationRule ="Not Is Null"
                    AfterUpdate ="[Event Procedure]"
                    OnExit ="[Event Procedure]"
                    OnKeyPress ="[Event Procedure]"
                    Tag ="10000"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3525
                            Top =3457
                            Width =615
                            Height =240
                            Name ="Label19"
                            Caption ="مبلغ كل"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =568
                    Top =3797
                    Height =255
                    ColumnOrder =14
                    TabIndex =8
                    Name ="FROM_A"
                    ControlSource ="FROM_A"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3435
                            Top =3797
                            Width =705
                            Height =240
                            Name ="Label21"
                            Caption ="FROM_A"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =568
                    Top =4138
                    Height =255
                    ColumnOrder =15
                    TabIndex =9
                    Name ="N_RASID"
                    ControlSource ="N_RASID"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3195
                            Top =4138
                            Width =945
                            Height =240
                            Name ="Label23"
                            Caption ="شماره رسيد"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =568
                    Top =4478
                    Height =255
                    ColumnOrder =16
                    TabIndex =10
                    Name ="MEGH_R"
                    ControlSource ="MEGH_R"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3270
                            Top =4478
                            Width =870
                            Height =240
                            Name ="Label25"
                            Caption ="مقدار رسيد"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =568
                    Top =4818
                    Height =255
                    ColumnOrder =17
                    TabIndex =11
                    Name ="RADAH"
                    ControlSource ="RADAH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3855
                            Top =4818
                            Width =285
                            Height =240
                            Name ="Label27"
                            Caption ="رده"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =568
                    Top =5158
                    Height =255
                    ColumnWidth =1417
                    ColumnOrder =19
                    TabIndex =12
                    Name ="SANAD_NO"
                    ControlSource ="SANAD_NO"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =3255
                            Top =5160
                            Width =1125
                            Height =405
                            Name ="Label29"
                            Caption ="ق. مصرف كننده"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =568
                    Top =5498
                    Height =255
                    ColumnOrder =20
                    TabIndex =13
                    Name ="CUST_NO"
                    ControlSource ="CUST_NO"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =3015
                            Top =5498
                            Width =1125
                            Height =240
                            Name ="Label31"
                            Caption ="شماره مشتري"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =568
                    Top =5838
                    Height =255
                    ColumnWidth =420
                    ColumnOrder =21
                    TabIndex =14
                    Name ="ANBARF"
                    ControlSource ="ANBARF"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3345
                            Top =5838
                            Width =795
                            Height =240
                            Name ="Label33"
                            Caption ="انبار فرعي"
                        End
                    End
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =3
                    Left =563
                    Top =1474
                    Width =1680
                    Height =255
                    ColumnWidth =5745
                    ColumnOrder =4
                    TabIndex =3
                    Name ="CODE"
                    ControlSource ="CODE"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT STUF_DEF.CODE, STUF_DEF.NAME, STUF_DEF.CODE FROM STUF_DEF ORDER BY STUF_D"
                        "EF.NAME"
                    ColumnWidths ="0;4320"
                    AfterUpdate ="[Event Procedure]"
                    OnEnter ="[Event Procedure]"
                    OnExit ="[Event Procedure]"
                    OnDblClick ="[Event Procedure]"
                    OnKeyUp ="[Event Procedure]"
                    Tag ="51"
                    OnNotInList ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3604
                            Top =1474
                            Width =510
                            Height =240
                            Name ="Label9"
                            Caption =" كالا"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =566
                    Top =2097
                    Height =255
                    ColumnWidth =735
                    ColumnOrder =8
                    TabIndex =16
                    Name ="MEGHk"
                    ControlSource ="MEGHk"
                    OnKeyPress ="[Event Procedure]"
                    Tag ="100"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2948
                            Top =2097
                            Width =1020
                            Height =240
                            Name ="Label38"
                            Caption ="مقدار كل "
                        End
                    End
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =568
                    Top =6179
                    Width =696
                    Height =255
                    ColumnWidth =1320
                    ColumnOrder =5
                    TabIndex =15
                    Name ="VAHED_K"
                    ControlSource ="VAHED_K"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT VAHEDS.VAHED, TCOD_VAHEDS.NAMES FROM TCOD_VAHEDS INNER JOIN VAHEDS ON TCO"
                        "D_VAHEDS.CODE = VAHEDS.VAHED GROUP BY VAHEDS.VAHED, TCOD_VAHEDS.NAMES ORDER BY T"
                        "COD_VAHEDS.NAMES"
                    ColumnWidths ="0"
                    AfterUpdate ="[Event Procedure]"
                    OnEnter ="[Event Procedure]"
                    OnExit ="[Event Procedure]"
                    Tag ="5"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3525
                            Top =6179
                            Width =615
                            Height =240
                            Name ="Label35"
                            Caption ="واحدكالا"
                        End
                    End
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =542
                    Top =793
                    Height =255
                    ColumnWidth =1545
                    ColumnOrder =3
                    TabIndex =2
                    Name ="ANBAR"
                    ControlSource ="ANBAR"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TCOD_ANBAR.CODE, TCOD_ANBAR.NAMES FROM TCOD_ANBAR ORDER BY TCOD_ANBAR.COD"
                        "E;"
                    ColumnWidths ="0"
                    AfterUpdate ="[Event Procedure]"
                    OnDblClick ="[Event Procedure]"
                    DefaultValue ="1"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3784
                            Top =793
                            Width =330
                            Height =240
                            Name ="Label5"
                            Caption ="انبار"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    ColumnHidden = NotDefault
                    OverlapFlags =247
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1644
                    Top =453
                    Width =1680
                    Height =255
                    ColumnOrder =22
                    TabIndex =17
                    Name ="CODEO"
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =226
                    Top =1133
                    ColumnWidth =1417
                    ColumnOrder =23
                    TabIndex =18
                    Name ="RADIF"
                    ControlSource ="RADIF"
                    DefaultValue ="[RecordsetClone].[RecordCount]+1"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3163
                            Top =1133
                            Width =465
                            Height =240
                            Name ="Label45"
                            Caption ="رديف"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =2
                    OverlapFlags =85
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =450
                    Top =6465
                    ColumnWidth =540
                    ColumnOrder =11
                    TabIndex =19
                    Name ="N_KOL"
                    ControlSource ="N_KOL"
                    ValidationRule ="<=100 And Not Is Null"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2974
                            Top =6465
                            Width =525
                            Height =240
                            Name ="Label50"
                            Caption ="تخفيف"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =450
                    Top =6805
                    ColumnWidth =900
                    ColumnOrder =13
                    TabIndex =20
                    Name ="N_MOIN"
                    ControlSource ="N_MOIN"
                    ValidationRule ="Not Is Null"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="0"
                    OnKeyPress ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2757
                            Top =6805
                            Width =870
                            Height =240
                            Name ="Label51"
                            Caption ="مبلغ تخفيف"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =247
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Top =3
                    ColumnWidth =990
                    ColumnOrder =18
                    TabIndex =21
                    Name ="IMBAA"
                    ControlSource ="IMBAA"
                    DefaultValue ="0"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =2302
                            Width =885
                            Height =240
                            Name ="Label53"
                            Caption ="ماليات ب.ا.ا"
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
                    Left =623
                    Top =2834
                    ColumnWidth =1530
                    ColumnOrder =24
                    TabIndex =22
                    Name ="MANDAH"
                    ControlSource ="MANDAH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3275
                            Top =2834
                            Width =750
                            Height =240
                            Name ="Label47"
                            Caption ="ملاحظات"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    DecimalPlaces =2
                    OverlapFlags =85
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =465
                    Top =7095
                    ColumnWidth =645
                    ColumnOrder =12
                    TabIndex =23
                    Name ="TKHN"
                    ControlSource ="TKHN"
                    ValidationRule ="<=100 And Not Is Null"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2985
                            Top =7095
                            Width =540
                            Height =240
                            Name ="Label56"
                            Caption ="%ت.ن"
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    IMESentenceMode =3
                    Left =420
                    Top =7530
                    ColumnWidth =525
                    ColumnOrder =2
                    TabIndex =24
                    Name ="JAY"
                    ControlSource ="JAY"
                    ConditionalFormat = Begin
                        0x010000006c000000020000000000000003000000000000000200000001000000 ,
                        0x33996600339966000000000002000000030000000500000001000000ffffff00 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x300000000000300000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3412
                            Top =7535
                            Width =420
                            Height =240
                            Name ="Label57"
                            Caption ="جايزه"
                        End
                    End
                End
                Begin ComboBox
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    IMESentenceMode =3
                    Left =420
                    Top =8260
                    ColumnWidth =3705
                    TabIndex =25
                    ConditionalFormat = Begin
                        0x010000006e000000020000000000000002000000000000000300000001000000 ,
                        0x33996600339966000000000002000000040000000600000001000000ffffff00 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x2d00310000000000300000000000
                    End
                    Name ="JAYO"
                    ControlSource ="JAYO"
                    RowSourceType ="Table/View/StoredProc"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3412
                            Top =8270
                            Width =855
                            Height =240
                            Name ="Label63"
                            Caption ="الگوي جايزه"
                        End
                    End
                End
            End
        End
        Begin FormFooter
            Height =1134
            BackColor =-2147483633
            Name ="FormFooter"
            Begin
                Begin TextBox
                    TabStop = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =680
                    Top =56
                    Height =255
                    Name ="SMABLK"
                    ControlSource ="=Sum([MABL_K])"
                End
            End
        End
    End
End
CodeBehindForm
' See "INVO_LST_PISHFROOSH_SUB2.cls"
