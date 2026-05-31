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
    GridX =64
    GridY =64
    Width =15760
    DatasheetFontHeight =10
    ItemSuffix =179
    Left =600
    Top =1890
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x52711df85311e640
    End
    RecordSource ="SELECT Q_LIST_FCT2.NUMBER, Q_LIST_FCT2.NAME, Q_LIST_FCT2.SumOfMABL_K, Q_LIST_FCT"
        "2.ADDRESS, Q_LIST_FCT2.DATE_N, Q_LIST_FCT2.TAG, Q_LIST_FCT2.AGHLAM, Q_LIST_FCT2."
        "hes, ISNULL(\"BED-BES\".BEST, 0) AS mand, CUST_HESAB.CITY, TOZIE.TDATE, TOZIE.TD"
        "RIVER, TOZIE.TCITY, TOZIE.TID, TOZIE.TMAMUR FROM Q_LIST_FCT2 INNER JOIN CUST_HES"
        "AB ON Q_LIST_FCT2.hes = CUST_HESAB.hes INNER JOIN TOZIE_SUB ON Q_LIST_FCT2.NUMBE"
        "R = TOZIE_SUB.NUMBER INNER JOIN TOZIE ON TOZIE_SUB.TID = TOZIE.TID LEFT OUTER JO"
        "IN \"BED-BES\" ON Q_LIST_FCT2.hes = \"BED-BES\".HES GROUP BY Q_LIST_FCT2.NUMBER,"
        " Q_LIST_FCT2.NAME, Q_LIST_FCT2.SumOfMABL_K, Q_LIST_FCT2.ADDRESS, Q_LIST_FCT2.DAT"
        "E_N, Q_LIST_FCT2.TAG, Q_LIST_FCT2.AGHLAM, Q_LIST_FCT2.hes, ISNULL(\"BED-BES\".BE"
        "ST, 0), CUST_HESAB.CITY, TOZIE.TDATE, TOZIE.TDRIVER, TOZIE.TCITY, TOZIE.TID, TOZ"
        "IE.TMAMUR HAVING (TOZIE.TID = 1) AND (Q_LIST_FCT2.TAG = 2)"
    Caption ="گزارش مامور پخش"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
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
            ControlSource ="NUMBER"
        End
        Begin BreakLevel
            ControlSource ="NAME"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =0
            Name ="ReportHeader"
        End
        Begin PageHeader
            Height =1350
            Name ="PageHeader"
            Begin
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =15322
                    Top =945
                    Width =390
                    Height =240
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="N_S Label"
                    Caption ="رديف"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="N_S_Label"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =13738
                    Top =945
                    Width =585
                    Height =240
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="SHARH Label"
                    Caption ="مشتري"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="SHARH_Label"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =11490
                    Top =796
                    Width =525
                    Height =539
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="BES Label"
                    Caption ="شماره فاكتور"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="BES_Label"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =9729
                    Top =796
                    Width =585
                    Height =539
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="Expr1 Label"
                    Caption ="مبــلغ فاكتور"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="Expr1_Label"
                End
                Begin Line
                    BorderWidth =2
                    Top =765
                    Width =15760
                    BorderColor =0
                    Name ="Line26"
                End
                Begin Line
                    BorderWidth =2
                    Top =1350
                    Width =15760
                    Name ="Line27"
                End
                Begin Line
                    BorderWidth =1
                    Left =15255
                    Top =781
                    Width =0
                    Height =569
                    BorderColor =0
                    Name ="Line28"
                End
                Begin Line
                    BorderWidth =1
                    Left =12064
                    Top =781
                    Width =0
                    Height =569
                    BorderColor =0
                    Name ="Line33"
                End
                Begin Line
                    BorderWidth =1
                    Left =11431
                    Top =781
                    Width =0
                    Height =569
                    BorderColor =0
                    Name ="Line34"
                End
                Begin Line
                    BorderWidth =1
                    Top =781
                    Width =0
                    Height =569
                    BorderColor =0
                    Name ="Line36"
                End
                Begin Line
                    BorderWidth =1
                    Left =15735
                    Top =781
                    Width =0
                    Height =569
                    BorderColor =0
                    Name ="Line37"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =7215
                    Top =945
                    Width =795
                    Height =240
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="Label73"
                    Caption ="كد مشتري"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =2
                    Left =7125
                    Top =781
                    Width =0
                    Height =569
                    BorderColor =0
                    Name ="Line76"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =6405
                    Top =945
                    Width =705
                    Height =240
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="Label82"
                    Caption ="وجه نقد"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =1
                    Left =5615
                    Top =781
                    Width =0
                    Height =569
                    BorderColor =0
                    Name ="Line84"
                End
                Begin Line
                    BorderWidth =1
                    Left =975
                    Top =781
                    Width =0
                    Height =569
                    BorderColor =0
                    Name ="Line88"
                End
                Begin Label
                    TextFontFamily =2
                    Left =14846
                    Top =61
                    Width =885
                    Height =240
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="Label100"
                    Caption ="مامور پخش:"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Top =88
                    Width =1256
                    Height =300
                    FontSize =8
                    Name ="Text116"
                    ControlSource ="=FARSIDATE(Date())"
                    FontName ="WeblogmaYekan"
                    InputMask ="####/##/##"
                End
                Begin Line
                    BorderWidth =1
                    Left =10559
                    Top =781
                    Width =0
                    Height =569
                    BorderColor =0
                    Name ="Line120"
                End
                Begin Label
                    TextFontFamily =2
                    Left =1301
                    Top =102
                    Width =435
                    Height =270
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="Label117"
                    Caption ="تاريخ:"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =1
                    Left =1891
                    Top =781
                    Width =0
                    Height =569
                    BorderColor =0
                    Name ="Line133"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =10725
                    Top =796
                    Width =555
                    Height =539
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="Label119"
                    Caption ="تاريخ فاكتور"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5295
                    Top =45
                    Width =4818
                    Height =270
                    FontSize =8
                    TabIndex =1
                    BorderColor =32768
                    Name ="Text90"
                    ControlSource ="=[Forms]![Baseknow]![WIDTH_D]"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextFontFamily =2
                    Left =13935
                    Top =420
                    Width =1800
                    Height =240
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="Label140"
                    Caption ="نام و نام خانوادگي راننده:"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =8460
                    Top =795
                    Width =585
                    Height =539
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="Label141"
                    Caption ="مانده حساب"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Left =9331
                    Top =781
                    Width =0
                    Height =569
                    BorderColor =0
                    Name ="Line142"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Left =8086
                    Top =781
                    Width =0
                    Height =569
                    BorderColor =0
                    Name ="Line147"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =4410
                    Top =945
                    Width =1215
                    Height =240
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="Label150"
                    Caption ="مبلغ چك"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =1
                    Left =4415
                    Top =781
                    Width =0
                    Height =569
                    BorderColor =0
                    Name ="Line151"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =3165
                    Top =945
                    Width =1245
                    Height =240
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="Label153"
                    Caption ="تاريخ چك"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =1
                    Left =3185
                    Top =781
                    Width =0
                    Height =569
                    BorderColor =0
                    Name ="Line154"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =1905
                    Top =945
                    Width =1275
                    Height =240
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="Label156"
                    Caption ="شماره چك"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =1005
                    Top =945
                    Width =855
                    Height =240
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="Label159"
                    Caption ="تخفيف"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =-15
                    Top =945
                    Width =975
                    Height =240
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="Label160"
                    Caption ="برگشتي"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextFontFamily =2
                    Left =6915
                    Top =360
                    Width =1575
                    Height =240
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="Label164"
                    Caption ="گزارش مامور پخش"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =1
                    Left =12840
                    Top =781
                    Width =0
                    Height =569
                    BorderColor =0
                    Name ="Line165"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =12060
                    Top =945
                    Width =765
                    Height =240
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="Label167"
                    Caption ="شهرستان"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =1
                    Left =6395
                    Top =781
                    Width =0
                    Height =569
                    BorderColor =0
                    Name ="Line169"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =5640
                    Top =945
                    Width =735
                    Height =240
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="Label171"
                    Caption ="كارت"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextFontFamily =2
                    Left =750
                    Top =420
                    Width =990
                    Height =240
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="Label172"
                    Caption ="شماره حواله:"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin TextBox
                    DecimalPlaces =0
                    BorderWidth =1
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =12165
                    Top =60
                    Width =2662
                    FontSize =8
                    TabIndex =2
                    Name ="TMAMUR"
                    ControlSource ="TMAMUR"
                    Format ="Fixed"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =0
                    BorderWidth =1
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =11280
                    Top =420
                    Width =2662
                    FontSize =8
                    TabIndex =3
                    Name ="TDRIVER"
                    ControlSource ="TDRIVER"
                    Format ="Fixed"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =0
                    BorderWidth =1
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =26
                    Top =433
                    Width =652
                    FontSize =8
                    TabIndex =4
                    Name ="TID"
                    ControlSource ="TID"
                    Format ="Fixed"
                    FontName ="WeblogmaYekan"
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =792
            Name ="Detail"
            Begin
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =12855
                    Top =15
                    Width =2370
                    Height =387
                    ColumnWidth =4650
                    FontSize =8
                    Name ="NAME"
                    ControlSource ="NAME"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =0
                    BorderWidth =1
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =11464
                    Top =15
                    Width =592
                    Height =387
                    FontSize =8
                    TabIndex =1
                    Name ="NUMBER"
                    ControlSource ="NUMBER"
                    Format ="Fixed"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    BackStyle =1
                    NumeralShapes =2
                    Left =9345
                    Top =15
                    Width =1191
                    Height =387
                    FontSize =8
                    TabIndex =2
                    BackColor =16777164
                    Name ="SumOfMABL_K"
                    ControlSource ="SumOfMABL_K"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    RunningSum =1
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =15274
                    Top =15
                    Width =471
                    Height =387
                    FontSize =8
                    TabIndex =3
                    Name ="N_S"
                    ControlSource ="=1"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    BorderWidth =1
                    Left =15255
                    Width =0
                    Height =417
                    BorderColor =0
                    Name ="Line38"
                End
                Begin Line
                    BorderWidth =1
                    Left =12064
                    Width =0
                    Height =417
                    BorderColor =0
                    Name ="Line43"
                End
                Begin Line
                    BorderWidth =1
                    Left =11430
                    Width =0
                    Height =417
                    BorderColor =0
                    Name ="Line44"
                End
                Begin Line
                    BorderWidth =1
                    Width =0
                    Height =792
                    BorderColor =0
                    Name ="Line46"
                End
                Begin Line
                    BorderWidth =1
                    Left =15735
                    Width =0
                    Height =792
                    BorderColor =0
                    Name ="Line47"
                End
                Begin Line
                    LineSlant = NotDefault
                    Top =780
                    Width =15760
                    BorderColor =4210752
                    Name ="Line48"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =7140
                    Top =15
                    Width =922
                    Height =387
                    FontSize =8
                    TabIndex =4
                    BackColor =16777164
                    Name ="hes"
                    ControlSource ="hes"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    BorderWidth =2
                    Left =7125
                    Width =0
                    Height =792
                    BorderColor =0
                    Name ="Line81"
                End
                Begin Line
                    BorderWidth =1
                    Left =5610
                    Width =0
                    Height =792
                    BorderColor =0
                    Name ="Line92"
                End
                Begin Line
                    BorderWidth =1
                    Left =975
                    Width =0
                    Height =792
                    BorderColor =0
                    Name ="Line96"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =10605
                    Top =15
                    Width =802
                    Height =387
                    FontSize =8
                    TabIndex =5
                    Name ="DATE_N"
                    ControlSource ="DATE_N"
                    FontName ="WeblogmaYekan"
                    InputMask ="####/##/##"
                End
                Begin Line
                    BorderWidth =1
                    Left =10560
                    Width =0
                    Height =417
                    BorderColor =0
                    Name ="Line122"
                End
                Begin Line
                    BorderWidth =1
                    Left =1892
                    Width =0
                    Height =792
                    BorderColor =0
                    Name ="Line134"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    BackStyle =1
                    NumeralShapes =2
                    Left =8115
                    Top =15
                    Width =1192
                    Height =387
                    FontSize =8
                    TabIndex =6
                    BackColor =13434828
                    Name ="mand"
                    ControlSource ="mand"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    BorderWidth =1
                    Left =9330
                    Width =0
                    Height =417
                    BorderColor =0
                    Name ="Line144"
                End
                Begin Line
                    BorderWidth =1
                    Left =8085
                    Width =0
                    Height =417
                    BorderColor =0
                    Name ="Line148"
                End
                Begin Line
                    BorderWidth =1
                    Left =4410
                    Width =0
                    Height =792
                    BorderColor =0
                    Name ="Line152"
                End
                Begin Line
                    BorderWidth =1
                    Left =3180
                    Width =0
                    Height =792
                    BorderColor =0
                    Name ="Line155"
                End
                Begin Line
                    BorderWidth =1
                    Left =12840
                    Width =0
                    Height =417
                    BorderColor =0
                    Name ="Line166"
                End
                Begin TextBox
                    DecimalPlaces =0
                    BorderWidth =1
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =12073
                    Top =15
                    Width =757
                    Height =387
                    FontSize =8
                    TabIndex =7
                    Name ="CITY"
                    ControlSource ="CITY"
                    Format ="Fixed"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    BorderWidth =1
                    Left =6390
                    Width =0
                    Height =792
                    BorderColor =0
                    Name ="Line170"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =7125
                    Top =405
                    Width =8635
                    BorderColor =4210752
                    Name ="Line177"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =7154
                    Top =389
                    Width =8565
                    Height =387
                    FontSize =8
                    TabIndex =8
                    Name ="ADDRESS"
                    ControlSource ="ADDRESS"
                    FontName ="WeblogmaYekan"
                End
            End
        End
        Begin PageFooter
            Height =0
            Name ="PageFooter"
        End
        Begin FormFooter
            KeepTogether = NotDefault
            Height =1134
            Name ="ReportFooter"
            Begin
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    BackStyle =1
                    NumeralShapes =2
                    Left =9345
                    Top =75
                    Width =1192
                    Height =295
                    FontSize =8
                    BackColor =16777164
                    Name ="SBES"
                    ControlSource ="=Sum([SumOfMABL_K])"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    BorderWidth =1
                    Width =0
                    Height =419
                    BorderColor =0
                    Name ="Line64"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Top =420
                    Width =15760
                    BorderColor =4210752
                    Name ="Line65"
                End
                Begin Line
                    BorderWidth =1
                    Left =15735
                    Width =0
                    Height =419
                    BorderColor =0
                    Name ="Line67"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =10904
                    Top =124
                    Width =660
                    Height =240
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="Label68"
                    Caption ="جمع كل:"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =2
                    Width =15760
                    BorderColor =0
                    Name ="Line71"
                End
                Begin Line
                    BorderWidth =1
                    Width =0
                    Height =419
                    BorderColor =0
                    Name ="Line114"
                End
                Begin Line
                    BorderWidth =2
                    Left =7126
                    Width =0
                    Height =419
                    BorderColor =0
                    Name ="Line137"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    BackStyle =1
                    NumeralShapes =2
                    Left =8115
                    Top =75
                    Width =1192
                    Height =295
                    FontSize =8
                    TabIndex =1
                    BackColor =13434828
                    Name ="Text145"
                    ControlSource ="=Sum([mand])"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    BorderWidth =1
                    Left =9332
                    Width =0
                    Height =419
                    BorderColor =0
                    Name ="Line146"
                End
                Begin Line
                    BorderWidth =1
                    Left =8086
                    Width =0
                    Height =419
                    BorderColor =0
                    Name ="Line149"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =5305
                    Top =106
                    Width =1125
                    Height =240
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="Label161"
                    Caption ="امضاء صندوقدار"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =1098
                    Top =106
                    Width =1155
                    Height =240
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="Label162"
                    Caption ="امضاء امور مالي"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =13978
                    Top =124
                    Width =1736
                    FontSize =8
                    TabIndex =2
                    Name ="Text24"
                    ControlSource ="=\"صفحه \" & [Page] & \" از \" & [Pages]"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    BorderWidth =1
                    Left =10559
                    Width =0
                    Height =419
                    BorderColor =0
                    Name ="Line163"
                End
            End
        End
    End
End
CodeBehindForm
' See "tozie_dasti.cls"
