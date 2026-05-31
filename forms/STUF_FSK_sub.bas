Version =20
VersionRequired =20
Begin Form
    DefaultView =2
    AllowUpdating =4
    ScrollBars =0
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =4705
    RowHeight =330
    ItemSuffix =29
    Left =7365
    Top =4920
    Right =21660
    Bottom =7080
    DatasheetGridlinesColor =12632256
    BeforeDelConfirm ="[Event Procedure]"
    RecSrcDt = Begin
        0x7b09edd5b95ee240
    End
    RecordSource ="STUF_FSK"
    Caption ="STUF_FSK"
    OnCurrent ="[Event Procedure]"
    OnDelete ="[Event Procedure]"
    AfterUpdate ="[Event Procedure]"
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
            Height =226
            BackColor =-2147483633
            Name ="FormHeader"
        End
        Begin Section
            Height =4705
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1029
                    Top =113
                    Width =1440
                    ColumnOrder =4
                    Name ="CODE"
                    ControlSource ="CODE"
                    FontName ="Arabic Transparent"
                    Tag ="5"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontFamily =2
                            Left =3845
                            Top =113
                            Width =750
                            Height =240
                            Name ="Label1"
                            Caption ="كد كالا"
                            FontName ="Titr"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =768
                    Top =1133
                    ColumnWidth =1020
                    ColumnOrder =1
                    TabIndex =2
                    Name ="MOGODI_A"
                    ControlSource ="MOGODI_A"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Arabic Transparent"
                    OnKeyPress ="[Event Procedure]"
                    Tag ="20"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontFamily =10
                            Left =3009
                            Top =1140
                            Width =1605
                            Height =285
                            Name ="Label7"
                            Caption ="مقدار"
                            FontName ="Arabic Transparent"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =794
                    Top =1474
                    ColumnWidth =1230
                    ColumnOrder =2
                    TabIndex =3
                    Name ="FI_A"
                    ControlSource ="FI_A"
                    Format ="Fixed"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Arabic Transparent"
                    OnKeyPress ="[Event Procedure]"
                    Tag ="20000"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontFamily =10
                            Left =3954
                            Top =1476
                            Width =675
                            Height =255
                            Name ="Label9"
                            Caption ="في "
                            FontName ="Arabic Transparent"
                        End
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =794
                    Top =1810
                    ColumnWidth =1215
                    ColumnOrder =3
                    TabIndex =4
                    Name ="MABL_A"
                    ControlSource ="MABL_A"
                    Format ="Fixed"
                    BeforeUpdate ="[Event Procedure]"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Arabic Transparent"
                    OnKeyPress ="[Event Procedure]"
                    Tag ="400000"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontFamily =10
                            Left =3714
                            Top =1806
                            Width =735
                            Height =255
                            Name ="Label11"
                            Caption ="مبلغ "
                            FontName ="Arabic Transparent"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =819
                    Top =2151
                    ColumnOrder =6
                    TabIndex =5
                    Name ="MANDAH_A"
                    ControlSource ="MANDAH_A"
                    FontName ="Arabic Transparent"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontFamily =10
                            Left =3476
                            Top =2151
                            Width =915
                            Height =240
                            Name ="Label13"
                            Caption ="مانده اوليه"
                            FontName ="Arabic Transparent"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =794
                    Top =2494
                    ColumnWidth =930
                    ColumnOrder =7
                    TabIndex =6
                    Name ="VAZ"
                    ControlSource ="VAZ"
                    FontName ="Arabic Transparent"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontFamily =10
                            Left =4006
                            Top =2494
                            Width =360
                            Height =240
                            Name ="Label15"
                            Caption ="وضعيت"
                            FontName ="Arabic Transparent"
                        End
                    End
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =1458
                    Top =453
                    Width =1011
                    ColumnWidth =3900
                    ColumnOrder =0
                    TabIndex =1
                    Name ="ANBAR"
                    ControlSource ="ANBAR"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TCOD_ANBAR.CODE, TCOD_ANBAR.NAMES, TCOD_ANBAR_KIND.ANB_KIND FROM TCOD_ANB"
                        "AR_KIND INNER JOIN TCOD_ANBAR ON TCOD_ANBAR_KIND.CODE = TCOD_ANBAR.KIND WHERE (T"
                        "COD_ANBAR.CODE > 0) ORDER BY TCOD_ANBAR.NAMES"
                    ColumnWidths ="0"
                    BeforeUpdate ="[Event Procedure]"
                    AfterUpdate ="[Event Procedure]"
                    OnEnter ="[Event Procedure]"
                    OnDblClick ="[Event Procedure]"
                    FontName ="Arabic Transparent"
                    Tag ="3"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontFamily =2
                            Left =3755
                            Top =453
                            Width =840
                            Height =240
                            Name ="Label3"
                            Caption ="انبار"
                            FontName ="Titr"
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
                    Top =2777
                    ColumnWidth =1410
                    ColumnOrder =5
                    TabIndex =7
                    Name ="POSITION"
                    ControlSource ="POSITION"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2532
                            Top =2777
                            Width =870
                            Height =240
                            Name ="Label24"
                            Caption ="موقعيت"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1077
                    Top =3174
                    ColumnOrder =9
                    TabIndex =8
                    Name ="B_SEF"
                    ControlSource ="B_SEF"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =3314
                            Top =3174
                            Width =630
                            Height =405
                            Name ="Label25"
                            Caption ="بهينه سفارش"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1077
                    Top =3514
                    ColumnOrder =8
                    TabIndex =9
                    Name ="N_SEF"
                    ControlSource ="N_SEF"
                    Begin
                        Begin Label
                            OverlapFlags =223
                            Left =3299
                            Top =3519
                            Width =660
                            Height =405
                            Name ="Label26"
                            Caption ="نقطه سفارش"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1077
                    Top =3854
                    TabIndex =10
                    Name ="MIN_M"
                    ControlSource ="MIN_M"
                    Begin
                        Begin Label
                            OverlapFlags =223
                            Left =3269
                            Top =3849
                            Width =705
                            Height =405
                            Name ="Label27"
                            Caption ="حداقل موجودي"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1077
                    Top =4194
                    TabIndex =11
                    Name ="MAX_M"
                    ControlSource ="MAX_M"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =3239
                            Top =4194
                            Width =750
                            Height =405
                            Name ="Label28"
                            Caption ="حداكثر موجودي"
                        End
                    End
                End
            End
        End
        Begin FormFooter
            Height =2445
            BackColor =-2147483633
            Name ="FormFooter"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1644
                    Top =226
                    Name ="SMOGODI_A"
                    ControlSource ="=Sum([MOGODI_A])"
                    FontName ="Arabic Transparent"
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1644
                    Top =586
                    TabIndex =1
                    Name ="SMABL_A"
                    ControlSource ="=Sum([MABL_A])"
                    FontName ="Arabic Transparent"
                End
            End
        End
    End
End
CodeBehindForm
' See "STUF_FSK_sub.cls"
