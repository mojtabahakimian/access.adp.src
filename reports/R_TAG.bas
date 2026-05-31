Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    OrderByOn = NotDefault
    DefaultView =0
    AllowUpdating =4
    TabularCharSet =178
    TabularFamily =42
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =5106
    DatasheetFontHeight =10
    ItemSuffix =229
    Left =270
    Top =780
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xa8c5919d50d5e340
    End
    RecordSource ="SELECT STUF_DEF.CODE, STUF_DEF.NAME, STUF_FSK.ANBAR, TCOD_ANBAR.NAMES, ANBGRD_LS"
        "T.GRD_NUM, TCOD_VAHEDS.NAMES AS vahed, STUF_FSK.POSITION FROM STUF_DEF INNER JOI"
        "N STUF_FSK ON STUF_DEF.CODE = STUF_FSK.CODE INNER JOIN TCOD_ANBAR ON STUF_FSK.AN"
        "BAR = TCOD_ANBAR.CODE INNER JOIN ANBGRD_LST ON STUF_DEF.CODE = ANBGRD_LST.CODE I"
        "NNER JOIN TCOD_VAHEDS ON STUF_DEF.VAHED = TCOD_VAHEDS.CODE"
    Caption ="چاپ شمارش اول"
    DatasheetFontName ="Arial (Arabic)"
    Begin
        Begin Label
            BackStyle =0
            TextFontCharSet =178
            TextAlign =1
            TextFontFamily =42
            FontSize =9
            FontWeight =700
            ForeColor =128
            FontName ="Arial (Arabic)"
        End
        Begin Rectangle
            BackStyle =0
            BorderWidth =1
            Width =850
            Height =850
        End
        Begin Line
            Width =1701
            BorderColor =128
        End
        Begin Image
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
            LabelX =230
            LabelY =-30
        End
        Begin CheckBox
            LabelX =230
            LabelY =-30
        End
        Begin OptionGroup
            BackStyle =1
            Width =1701
            Height =1701
        End
        Begin BoundObjectFrame
            Width =4536
            Height =2835
            LabelX =-1701
        End
        Begin TextBox
            OldBorderStyle =0
            TextFontCharSet =178
            TextFontFamily =42
            BackStyle =0
            Width =1701
            LabelX =-1701
            FontSize =9
            FontName ="Arial (Arabic)"
        End
        Begin ListBox
            TextFontCharSet =178
            TextFontFamily =42
            OldBorderStyle =0
            Width =1701
            Height =1417
            LabelX =-1701
            FontSize =9
            FontName ="Arial (Arabic)"
        End
        Begin ComboBox
            OldBorderStyle =0
            TextFontCharSet =178
            TextFontFamily =42
            BackStyle =0
            Width =1701
            LabelX =-1701
            FontSize =9
            FontName ="Arial (Arabic)"
        End
        Begin Subform
            OldBorderStyle =0
            Width =1701
            Height =1701
        End
        Begin UnboundObjectFrame
            Width =4536
            Height =2835
        End
        Begin BreakLevel
            ControlSource ="NAME"
        End
        Begin Section
            KeepTogether = NotDefault
            Height =7530
            Name ="Detail"
            Begin
                Begin TextBox
                    RunningSum =1
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2655
                    Top =1260
                    Width =1866
                    Height =300
                    FontSize =8
                    Name ="CODE"
                    ControlSource ="CODE"
                    FontName ="Titr"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =60
                    Top =915
                    Width =1086
                    Height =300
                    FontSize =8
                    TabIndex =1
                    Name ="vahed"
                    ControlSource ="vahed"
                    FontName ="Titr"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2025
                    Top =920
                    Width =2601
                    Height =300
                    FontSize =8
                    TabIndex =2
                    Name ="Text109"
                    ControlSource ="NAMES"
                    FontName ="Titr"
                    Begin
                        Begin Label
                            TextAlign =0
                            TextFontFamily =2
                            Left =4665
                            Top =915
                            Width =390
                            Height =315
                            FontSize =8
                            FontWeight =400
                            ForeColor =0
                            Name ="Label110"
                            Caption =" انبار:"
                            FontName ="Titr"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    TextAlign =3
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =170
                    Top =621
                    Width =786
                    Height =300
                    FontSize =8
                    TabIndex =3
                    Name ="GRD_NUM"
                    ControlSource ="GRD_NUM"
                    FontName ="Titr"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =4440
                    Top =1620
                    Width =630
                    Height =300
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="Label164"
                    Caption ="شرح كالا:"
                    FontName ="Titr"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =4534
                    Top =1262
                    Width =540
                    Height =300
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="Label168"
                    Caption ="كد كالا:"
                    FontName ="Titr"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =1200
                    Top =915
                    Width =675
                    Height =300
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="Label171"
                    Caption ="واحد كالا:"
                    FontName ="Titr"
                    Tag ="DetachedLabel"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Top =285
                    Width =5103
                    Height =300
                    FontSize =8
                    TabIndex =4
                    BorderColor =32768
                    Name ="Text90"
                    ControlSource ="=[Forms]![Baseknow]![WIDTH_D]"
                    FontName ="Titr"
                End
                Begin Label
                    TextAlign =0
                    TextFontFamily =2
                    Left =1890
                    Top =555
                    Width =1035
                    Height =300
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="Label185"
                    Caption ="تگ انبار گرداني"
                    FontName ="Titr"
                End
                Begin TextBox
                    DecimalPlaces =3
                    RunningSum =1
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    BackStyle =1
                    NumeralShapes =2
                    Left =150
                    Top =3600
                    Width =4266
                    Height =300
                    FontSize =8
                    TabIndex =5
                    Name ="Text187"
                    ControlSource ="NAME"
                    Format ="Standard"
                    FontName ="Titr"
                End
                Begin TextBox
                    RunningSum =1
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2655
                    Top =3255
                    Width =1866
                    Height =300
                    FontSize =8
                    TabIndex =6
                    Name ="Text188"
                    ControlSource ="CODE"
                    FontName ="Titr"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =60
                    Top =2916
                    Width =1086
                    Height =300
                    FontSize =8
                    TabIndex =7
                    Name ="Text189"
                    ControlSource ="vahed"
                    FontName ="Titr"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2055
                    Top =2915
                    Width =2586
                    Height =300
                    FontSize =8
                    TabIndex =8
                    Name ="Text190"
                    ControlSource ="NAMES"
                    FontName ="Titr"
                    Begin
                        Begin Label
                            TextAlign =0
                            TextFontFamily =2
                            Left =4680
                            Top =2910
                            Width =390
                            Height =300
                            FontSize =8
                            FontWeight =400
                            ForeColor =0
                            Name ="Label191"
                            Caption =" انبار:"
                            FontName ="Titr"
                        End
                    End
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =4046
                    Top =3963
                    Width =960
                    Height =300
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="Label192"
                    Caption ="تعداد شمارش"
                    FontName ="Titr"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =4440
                    Top =3600
                    Width =630
                    Height =300
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="Label193"
                    Caption ="شرح كالا:"
                    FontName ="Titr"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =4534
                    Top =3257
                    Width =540
                    Height =300
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="Label194"
                    Caption ="كد كالا:"
                    FontName ="Titr"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =1207
                    Top =2918
                    Width =675
                    Height =300
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="Label195"
                    Caption ="واحد كالا:"
                    FontName ="Titr"
                    Tag ="DetachedLabel"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Top =2310
                    Width =5043
                    Height =300
                    FontSize =8
                    TabIndex =9
                    BorderColor =32768
                    Name ="Text196"
                    ControlSource ="=[Forms]![Baseknow]![WIDTH_D]"
                    FontName ="Titr"
                End
                Begin Label
                    TextAlign =0
                    TextFontFamily =2
                    Left =1605
                    Top =2601
                    Width =1875
                    Height =300
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="Label197"
                    Caption ="تگ انبار گرداني- شمارش دوم"
                    FontName ="Titr"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =225
                    Top =3960
                    Width =1605
                    Height =300
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="Label198"
                    Caption ="نام و امضاء مسئول شمارش"
                    FontName ="Titr"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OldBorderStyle =1
                    BorderWidth =1
                    TextAlign =2
                    TextFontFamily =2
                    Left =4046
                    Top =4278
                    Width =945
                    Height =300
                    FontSize =8
                    FontWeight =400
                    ForeColor =16777215
                    Name ="Label199"
                    Caption ="."
                    FontName ="Titr"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =2
                    Top =2100
                    Width =5102
                    BorderColor =0
                    Name ="Line200"
                End
                Begin TextBox
                    DecimalPlaces =3
                    RunningSum =1
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    BackStyle =1
                    NumeralShapes =2
                    Left =180
                    Top =6155
                    Width =4266
                    Height =300
                    FontSize =8
                    TabIndex =10
                    Name ="Text201"
                    ControlSource ="NAME"
                    Format ="Standard"
                    FontName ="Titr"
                End
                Begin TextBox
                    RunningSum =1
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2685
                    Top =5790
                    Width =1866
                    Height =300
                    FontSize =8
                    TabIndex =11
                    Name ="Text202"
                    ControlSource ="CODE"
                    FontName ="Titr"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =60
                    Top =5445
                    Width =1086
                    Height =300
                    FontSize =8
                    TabIndex =12
                    Name ="Text203"
                    ControlSource ="vahed"
                    FontName ="Titr"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2070
                    Top =5450
                    Width =2586
                    Height =300
                    FontSize =8
                    TabIndex =13
                    Name ="Text204"
                    ControlSource ="NAMES"
                    FontName ="Titr"
                    Begin
                        Begin Label
                            TextAlign =0
                            TextFontFamily =2
                            Left =4695
                            Top =5445
                            Width =390
                            Height =300
                            FontSize =8
                            FontWeight =400
                            ForeColor =0
                            Name ="Label205"
                            Caption =" انبار:"
                            FontName ="Titr"
                        End
                    End
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =4046
                    Top =6513
                    Width =960
                    Height =300
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="Label206"
                    Caption ="تعداد شمارش"
                    FontName ="Titr"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =4457
                    Top =6150
                    Width =630
                    Height =300
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="Label207"
                    Caption ="شرح كالا:"
                    FontName ="Titr"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =4564
                    Top =5792
                    Width =540
                    Height =300
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="Label208"
                    Caption ="كد كالا:"
                    FontName ="Titr"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =1207
                    Top =5447
                    Width =675
                    Height =300
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="Label209"
                    Caption ="واحد كالا:"
                    FontName ="Titr"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =225
                    Top =6510
                    Width =1605
                    Height =300
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="Label212"
                    Caption ="نام و امضاء مسئول شمارش"
                    FontName ="Titr"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OldBorderStyle =1
                    BorderWidth =1
                    TextAlign =2
                    TextFontFamily =2
                    Left =4046
                    Top =6828
                    Width =945
                    Height =300
                    FontSize =8
                    FontWeight =400
                    ForeColor =16777215
                    Name ="Label213"
                    Caption ="."
                    FontName ="Titr"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =2
                    Top =4650
                    Width =5102
                    BorderColor =0
                    Name ="Line214"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Top =4845
                    Width =5043
                    Height =300
                    FontSize =8
                    TabIndex =14
                    BorderColor =32768
                    Name ="Text215"
                    ControlSource ="=[Forms]![Baseknow]![WIDTH_D]"
                    FontName ="Titr"
                End
                Begin Label
                    TextAlign =0
                    TextFontFamily =2
                    Left =1605
                    Top =5130
                    Width =1905
                    Height =300
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="Label216"
                    Caption ="تگ انبار گرداني- شمارش اول"
                    FontName ="Titr"
                End
                Begin TextBox
                    DecimalPlaces =3
                    RunningSum =1
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    BackStyle =1
                    NumeralShapes =2
                    Left =120
                    Top =1620
                    Width =4311
                    Height =300
                    FontSize =8
                    TabIndex =15
                    Name ="Text217"
                    ControlSource ="NAME"
                    Format ="Standard"
                    FontName ="Titr"
                End
                Begin Line
                    BorderWidth =2
                    Top =7530
                    Width =5106
                    Name ="Line218"
                End
                Begin TextBox
                    RunningSum =2
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =60
                    Top =1230
                    Width =696
                    Height =360
                    FontWeight =700
                    TabIndex =16
                    Name ="Text219"
                    ControlSource ="=1"
                    FontName ="Titr"
                End
                Begin Label
                    TextFontFamily =2
                    Left =793
                    Top =1247
                    Width =510
                    Height =345
                    FontWeight =400
                    ForeColor =0
                    Name ="Label220"
                    Caption ="سريال:"
                    FontName ="Titr"
                    Tag ="DetachedLabel"
                End
                Begin TextBox
                    RunningSum =2
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =56
                    Top =3231
                    Width =696
                    Height =360
                    FontWeight =700
                    TabIndex =17
                    Name ="Text225"
                    ControlSource ="=1"
                    FontName ="Titr"
                End
                Begin Label
                    TextFontFamily =2
                    Left =789
                    Top =3248
                    Width =510
                    Height =345
                    FontWeight =400
                    ForeColor =0
                    Name ="Label226"
                    Caption ="سريال:"
                    FontName ="Titr"
                    Tag ="DetachedLabel"
                End
                Begin TextBox
                    RunningSum =2
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =56
                    Top =5782
                    Width =696
                    Height =360
                    FontWeight =700
                    TabIndex =18
                    Name ="Text227"
                    ControlSource ="=1"
                    FontName ="Titr"
                End
                Begin Label
                    TextFontFamily =2
                    Left =789
                    Top =5799
                    Width =510
                    Height =345
                    FontWeight =400
                    ForeColor =0
                    Name ="Label228"
                    Caption ="سريال:"
                    FontName ="Titr"
                    Tag ="DetachedLabel"
                End
            End
        End
    End
End
CodeBehindForm
' See "R_TAG.cls"
