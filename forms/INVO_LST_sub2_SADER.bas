Version =20
VersionRequired =20
Begin Form
    AllowDeletions = NotDefault
    AllowAdditions = NotDefault
    KeyPreview = NotDefault
    DefaultView =2
    AllowUpdating =4
    ViewsAllowed =2
    TabularCharSet =178
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =4440
    RowHeight =315
    ItemSuffix =58
    Left =345
    Right =15090
    Bottom =5040
    HelpContextId =6001
    DatasheetForeColor =33554432
    DatasheetGridlinesColor =12632256
    BeforeDelConfirm ="[Event Procedure]"
    RecSrcDt = Begin
        0xb9c7ab017eb8e240
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
    OnError ="[Event Procedure]"
    AllowFormView =0
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
            Height =9127
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =93
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =612
                    Top =480
                    Height =255
                    ColumnWidth =1417
                    ColumnOrder =0
                    Name ="NUMBER"
                    ControlSource ="NUMBER"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3224
                            Top =480
                            Width =960
                            Height =240
                            Name ="Label1"
                            Caption ="شماره فاكتور"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =612
                    Top =820
                    Height =255
                    ColumnWidth =1417
                    ColumnOrder =1
                    TabIndex =1
                    Name ="TAG"
                    ControlSource ="TAG"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3554
                            Top =820
                            Width =630
                            Height =240
                            Name ="Label3"
                            Caption ="برچسب"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =599
                    Top =1814
                    Height =255
                    ColumnWidth =765
                    ColumnOrder =5
                    TabIndex =4
                    Name ="MEGH"
                    ControlSource ="MEGH"
                    AfterUpdate ="[Event Procedure]"
                    OnKeyPress ="[Event Procedure]"
                    Tag ="100"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3496
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
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =625
                    Top =2437
                    Height =255
                    ColumnOrder =6
                    TabIndex =5
                    Name ="MEGH_MAR"
                    ControlSource ="MEGH_MAR"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3087
                            Top =2437
                            Width =1110
                            Height =240
                            Name ="Label13"
                            Caption ="مقدار مرجوعي"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =625
                    Top =2777
                    Height =255
                    ColumnWidth =1417
                    ColumnOrder =25
                    TabIndex =6
                    Name ="MANDAH"
                    ControlSource ="MANDAH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3780
                            Top =2775
                            Width =660
                            Height =240
                            Name ="Label15"
                            Caption ="ملاحظات"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =625
                    Top =3117
                    Height =255
                    ColumnWidth =1050
                    ColumnOrder =10
                    TabIndex =7
                    Name ="MABL"
                    ControlSource ="MABL"
                    Format ="#,##0.00;-#,##0.00"
                    AfterUpdate ="[Event Procedure]"
                    OnEnter ="[Event Procedure]"
                    OnExit ="[Event Procedure]"
                    Tag ="253"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3822
                            Top =3117
                            Width =375
                            Height =240
                            Name ="Label17"
                            Caption ="مبلغ"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =625
                    Top =3457
                    Height =255
                    ColumnWidth =1155
                    ColumnOrder =11
                    TabIndex =8
                    Name ="MABL_K"
                    ControlSource ="MABL_K"
                    Format ="#,##0.00;-#,##0.00"
                    AfterUpdate ="[Event Procedure]"
                    OnExit ="[Event Procedure]"
                    Tag ="25300"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3582
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
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =625
                    Top =3797
                    Height =255
                    ColumnOrder =12
                    TabIndex =9
                    Name ="FROM_A"
                    ControlSource ="FROM_A"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3492
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
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =625
                    Top =4138
                    Height =255
                    ColumnOrder =13
                    TabIndex =10
                    Name ="N_RASID"
                    ControlSource ="N_RASID"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3252
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
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =625
                    Top =4478
                    Height =255
                    ColumnOrder =14
                    TabIndex =11
                    Name ="MEGH_R"
                    ControlSource ="MEGH_R"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3327
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
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =625
                    Top =4818
                    Height =255
                    ColumnOrder =15
                    TabIndex =12
                    Name ="RADAH"
                    ControlSource ="RADAH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3912
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
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =625
                    Top =5158
                    Height =255
                    ColumnOrder =16
                    TabIndex =13
                    Name ="SANAD_NO"
                    ControlSource ="SANAD_NO"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3312
                            Top =5158
                            Width =885
                            Height =240
                            Name ="Label29"
                            Caption ="شماره سند"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =625
                    Top =5498
                    Height =255
                    ColumnOrder =17
                    TabIndex =14
                    Name ="CUST_NO"
                    ControlSource ="CUST_NO"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3072
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
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =625
                    Top =5838
                    Height =255
                    ColumnWidth =420
                    ColumnOrder =18
                    TabIndex =15
                    Name ="ANBARF"
                    ControlSource ="ANBARF"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3402
                            Top =5838
                            Width =795
                            Height =240
                            Name ="Label33"
                            Caption ="انبار فرعي"
                        End
                    End
                End
                Begin ComboBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =3
                    Left =620
                    Top =1474
                    Width =1680
                    Height =255
                    ColumnWidth =3705
                    ColumnOrder =3
                    TabIndex =3
                    Name ="CODE"
                    ControlSource ="CODE"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT STUF_DEF.CODE, STUF_DEF.NAME, STUF_DEF.CODE FROM STUF_DEF ORDER BY STUF_D"
                        "EF.NAME"
                    ColumnWidths ="0;2268;567"
                    ValidationRule ="Is Not Null"
                    AfterUpdate ="[Event Procedure]"
                    OnEnter ="[Event Procedure]"
                    OnExit ="[Event Procedure]"
                    OnKeyUp ="[Event Procedure]"
                    Tag ="1"
                    OnNotInList ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3661
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
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =623
                    Top =2097
                    Height =255
                    ColumnWidth =795
                    ColumnOrder =7
                    TabIndex =17
                    Name ="MEGHk"
                    ControlSource ="MEGHk"
                    Tag ="100"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3005
                            Top =2097
                            Width =1020
                            Height =240
                            Name ="Label38"
                            Caption ="مقدار كل "
                        End
                    End
                End
                Begin ComboBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =625
                    Top =6179
                    Height =255
                    ColumnWidth =990
                    ColumnOrder =4
                    TabIndex =16
                    Name ="VAHED_K"
                    ControlSource ="VAHED_K"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT VAHEDS.VAHED, TCOD_VAHEDS.NAMES FROM TCOD_VAHEDS INNER JOIN VAHEDS ON TCO"
                        "D_VAHEDS.CODE = VAHEDS.VAHED GROUP BY VAHEDS.VAHED, TCOD_VAHEDS.NAMES ORDER BY T"
                        "COD_VAHEDS.NAMES"
                    ColumnWidths ="0"
                    ValidationRule ="Not Is Null"
                    AfterUpdate ="[Event Procedure]"
                    OnEnter ="[Event Procedure]"
                    OnExit ="[Event Procedure]"
                    Tag ="4"
                    OnNotInList ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3342
                            Top =6179
                            Width =855
                            Height =240
                            Name ="Label35"
                            Caption ="واحدكالا"
                        End
                    End
                End
                Begin ComboBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =627
                    Top =1160
                    Width =1686
                    Height =255
                    ColumnWidth =1560
                    ColumnOrder =2
                    TabIndex =2
                    Name ="ANBAR"
                    ControlSource ="ANBAR"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT     TCOD_ANBAR.CODE, TCOD_ANBAR.NAMES, OPANBACCESS.USERCO FROM  dbo.TCOD_"
                        "ANBAR INNER JOIN  dbo.OPANBACCESS ON dbo.TCOD_ANBAR.CODE = dbo.OPANBACCESS.ANBCO"
                        " WHERE (OPANBACCESS.USERCO = 2 ) ORDER BY TCOD_ANBAR.CODE"
                    ColumnWidths ="0"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="[Forms]![BASEKNOW]![DEFANB]"
                    OnNotInList ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3854
                            Top =1160
                            Width =330
                            Height =240
                            Name ="Label5"
                            Caption ="انبار"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =623
                    Top =6462
                    ColumnOrder =19
                    TabIndex =18
                    Name ="RADIF"
                    ControlSource ="RADIF"
                    DefaultValue ="[RecordsetClone].[RecordCount]+1"
                End
                Begin TextBox
                    Visible = NotDefault
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =644
                    Width =1680
                    Height =255
                    ColumnOrder =20
                    TabIndex =19
                    Name ="CODEO"
                End
                Begin TextBox
                    Visible = NotDefault
                    ColumnHidden = NotDefault
                    OverlapFlags =247
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =623
                    Top =283
                    Width =1680
                    Height =255
                    ColumnOrder =21
                    TabIndex =20
                    Name ="VH"
                End
                Begin TextBox
                    DecimalPlaces =2
                    OverlapFlags =85
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =623
                    Top =6803
                    ColumnWidth =555
                    ColumnOrder =22
                    TabIndex =21
                    Name ="N_KOL"
                    ControlSource ="N_KOL"
                    ValidationRule ="<=100"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3147
                            Top =6803
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
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =623
                    Top =7143
                    ColumnWidth =975
                    ColumnOrder =23
                    TabIndex =22
                    Name ="N_MOIN"
                    ControlSource ="N_MOIN"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2930
                            Top =7143
                            Width =870
                            Height =240
                            Name ="Label51"
                            Caption ="مبلغ تخفيف"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    ColumnHidden = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =85
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =623
                    Top =7503
                    ColumnWidth =975
                    ColumnOrder =24
                    TabIndex =23
                    Name ="IMBAA"
                    ControlSource ="IMBAA"
                    DefaultValue ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2925
                            Top =7500
                            Width =885
                            Height =240
                            Name ="Label53"
                            Caption ="ماليات ب.ا.ا"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =630
                    Top =7875
                    ColumnWidth =855
                    ColumnOrder =8
                    TabIndex =24
                    Name ="N_TAF"
                    ControlSource ="N_TAF"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2937
                            Top =7875
                            Width =1095
                            Height =240
                            Name ="Label54"
                            Caption ="Unit Price"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =630
                    Top =8215
                    ColumnWidth =1200
                    ColumnOrder =9
                    TabIndex =25
                    Name ="TOTALARZ"
                    ControlSource ="TOTALARZ"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3117
                            Top =8215
                            Width =915
                            Height =240
                            Name ="Label55"
                            Caption ="Line Total"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =510
                    Top =8730
                    TabIndex =26
                    Name ="AVRAGE"
                    ControlSource ="AVRAGE"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3177
                            Top =8730
                            Width =735
                            Height =240
                            Name ="Label56"
                            Caption ="AVRAGE:"
                        End
                    End
                End
            End
        End
        Begin FormFooter
            Height =645
            BackColor =-2147483633
            Name ="FormFooter"
            Begin
                Begin TextBox
                    TabStop = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =732
                    Top =15
                    Height =255
                    Name ="SMABLK"
                    ControlSource ="=Sum([MABL_K])"
                End
                Begin TextBox
                    TabStop = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2551
                    Top =113
                    Height =255
                    TabIndex =1
                    Name ="TOTALARZS"
                    ControlSource ="=Sum([TOTALARZ])"
                End
            End
        End
    End
End
CodeBehindForm
' See "INVO_LST_sub2_SADER.cls"
