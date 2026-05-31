Version =20
VersionRequired =20
Begin Form
    OrderByOn = NotDefault
    DefaultView =2
    AllowUpdating =4
    ScrollBars =0
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =4629
    RowHeight =330
    ItemSuffix =25
    Left =5100
    Top =4140
    Right =14400
    Bottom =6180
    DatasheetGridlinesColor =12632256
    BeforeDelConfirm ="[Event Procedure]"
    RecSrcDt = Begin
        0x24999a722ef7e340
    End
    RecordSource ="TR_STUF_FSK"
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
            Height =4486
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1029
                    Width =1440
                    Name ="CODE"
                    ControlSource ="CODE"
                    FontName ="Arabic Transparent"
                    Tag ="5"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontFamily =2
                            Left =3845
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
                    Top =1020
                    TabIndex =1
                    Name ="MOGODI_A"
                    ControlSource ="MOGODI_A"
                    FontName ="Arabic Transparent"
                    Tag ="20"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontFamily =10
                            Left =3009
                            Top =1027
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
                    Top =1361
                    TabIndex =2
                    Name ="FI_A"
                    ControlSource ="FI_A"
                    Format ="Fixed"
                    FontName ="Arabic Transparent"
                    Tag ="20000"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontFamily =10
                            Left =3954
                            Top =1363
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
                    Top =1697
                    TabIndex =3
                    Name ="MABL_A"
                    ControlSource ="MABL_A"
                    Format ="Fixed"
                    FontName ="Arabic Transparent"
                    Tag ="400000"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontFamily =10
                            Left =3714
                            Top =1693
                            Width =735
                            Height =255
                            Name ="Label11"
                            Caption ="مبلغ "
                            FontName ="Arabic Transparent"
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
                    Left =819
                    Top =2038
                    TabIndex =4
                    Name ="MANDAH_A"
                    ControlSource ="MANDAH_A"
                    FontName ="Arabic Transparent"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontFamily =10
                            Left =3476
                            Top =2038
                            Width =915
                            Height =240
                            Name ="Label13"
                            Caption ="مانده اوليه"
                            FontName ="Arabic Transparent"
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
                    Left =794
                    Top =2381
                    TabIndex =5
                    Name ="VAZ"
                    ControlSource ="VAZ"
                    FontName ="Arabic Transparent"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontFamily =10
                            Left =4006
                            Top =2381
                            Width =360
                            Height =240
                            Name ="Label15"
                            Caption ="وضعيت"
                            FontName ="Arabic Transparent"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =1458
                    Top =340
                    Width =1011
                    TabIndex =6
                    Name ="ANBAR"
                    ControlSource ="ANBAR"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TCOD_ANBAR.CODE, TCOD_ANBAR.NAMES, TCOD_ANBAR_KIND.ANB_KIND FROM TCOD_ANB"
                        "AR_KIND INNER JOIN TCOD_ANBAR ON TCOD_ANBAR_KIND.CODE = TCOD_ANBAR.KIND WHERE (T"
                        "COD_ANBAR.CODE > 0) ORDER BY TCOD_ANBAR.NAMES"
                    ColumnWidths ="0"
                    FontName ="Arabic Transparent"
                    Tag ="3"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontFamily =2
                            Left =3755
                            Top =340
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
                    Top =2664
                    TabIndex =7
                    Name ="POSITION"
                    ControlSource ="POSITION"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2532
                            Top =2664
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
                    Top =3061
                    TabIndex =8
                    Name ="B_SEF"
                    ControlSource ="B_SEF"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =3314
                            Top =3061
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
                    Top =3401
                    TabIndex =9
                    Name ="N_SEF"
                    ControlSource ="N_SEF"
                    Begin
                        Begin Label
                            OverlapFlags =223
                            Left =3299
                            Top =3406
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
                    Top =3741
                    TabIndex =10
                    Name ="MIN_M"
                    ControlSource ="MIN_M"
                    Begin
                        Begin Label
                            OverlapFlags =223
                            Left =3269
                            Top =3736
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
                    Top =4081
                    TabIndex =11
                    Name ="MAX_M"
                    ControlSource ="MAX_M"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =3239
                            Top =4081
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
            Height =1186
            BackColor =-2147483633
            Name ="FormFooter"
        End
    End
End
CodeBehindForm
' See "TR_STUF_FSK_sub.cls"
