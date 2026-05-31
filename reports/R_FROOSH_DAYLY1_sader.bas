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
    Width =10204
    DatasheetFontHeight =10
    ItemSuffix =100
    Left =600
    Top =1590
    DatasheetGridlinesColor =12632256
    Tag ="@Forms___F_MENU_KHFR___DT1 bigint=[Forms]![F_MENU_KHFR]![DT1],@Forms___F_MENU_KH"
        "FR___DT2 bigint=[Forms]![F_MENU_KHFR]![DT2],@Forms___F_MENU_KHFR___MMOIN NVARCHA"
        "R(20)=[Forms]![F_MENU_KHFR]![MMOIN]"
    RecSrcDt = Begin
        0x63d6645b9fe5e340
    End
    RecordSource ="SELECT HEAD_LST.NUMBER, HEAD_LST.NUMBER1, HEAD_LST.CUST_NO, HEAD_LST.SADER, SUM("
        "INVO_LST.MEGH) AS MEGHs, SUM(INVO_LST.MEGHk) AS MEGHks, SUM(INVO_LST.TOTALARZ) A"
        "S TOTALARZs, HEAD_LST.DATE_N, HEAD_LST_1.DATE_N AS DATE_Nh, CUST_HESAB.NAME, CUS"
        "T_HESAB.hes, HEAD_LST.ANBARF, HEAD_LST_1.ANBAR, 1 AS kk FROM HEAD_LST INNER JOIN"
        " INVO_LST ON HEAD_LST.NUMBER = INVO_LST.NUMBER AND HEAD_LST.TAG - 11 = INVO_LST."
        "TAG INNER JOIN HEAD_LST HEAD_LST_1 ON INVO_LST.NUMBER = HEAD_LST_1.NUMBER AND IN"
        "VO_LST.TAG = HEAD_LST_1.TAG INNER JOIN CUST_HESAB ON HEAD_LST.CUST_NO = CUST_HES"
        "AB.hes WHERE (HEAD_LST.TAG = 13) GROUP BY HEAD_LST.NUMBER, HEAD_LST.NUMBER1, HEA"
        "D_LST.CUST_NO, HEAD_LST.SADER, HEAD_LST.DATE_N, HEAD_LST_1.DATE_N, CUST_HESAB.NA"
        "ME, CUST_HESAB.hes, HEAD_LST.ANBARF, HEAD_LST_1.ANBAR HAVING (HEAD_LST.SADER = 1"
        ")"
    Caption ="فروش روزانه"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Arial (Arabic)"
    OnError ="[Event Procedure]"
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
            GroupFooter = NotDefault
            ControlSource ="KK"
        End
        Begin BreakLevel
            ControlSource ="DATE_N"
        End
        Begin BreakLevel
            ControlSource ="NUMBER1"
        End
        Begin PageHeader
            Height =1670
            Name ="PageHeader"
            Begin
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =9570
                    Top =857
                    Width =555
                    Height =765
                    FontSize =10
                    ForeColor =0
                    Name ="N_S Label"
                    Caption ="شماره رديف"
                    FontName ="Traffic"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="N_S_Label"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =8895
                    Top =863
                    Width =645
                    Height =765
                    FontSize =10
                    ForeColor =0
                    Name ="DATE_S Label"
                    Caption ="كد شخص"
                    FontName ="Traffic"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="DATE_S_Label"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =6540
                    Top =875
                    Width =2311
                    Height =765
                    FontSize =10
                    ForeColor =0
                    Name ="SHARH Label"
                    Caption ="نام شخص"
                    FontName ="Traffic"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="SHARH_Label"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =5220
                    Top =855
                    Width =636
                    Height =795
                    FontSize =10
                    ForeColor =0
                    Name ="BED Label"
                    Caption ="شماره حواله"
                    FontName ="Traffic"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="BED_Label"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =3195
                    Top =850
                    Width =936
                    Height =765
                    FontSize =10
                    ForeColor =0
                    Name ="BES Label"
                    Caption ="تاريخ حواله"
                    FontName ="Traffic"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="BES_Label"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =2370
                    Top =853
                    Width =756
                    Height =765
                    FontSize =10
                    ForeColor =0
                    Name ="MAND Label"
                    Caption ="كارتن"
                    FontName ="Traffic"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="MAND_Label"
                End
                Begin Line
                    BorderWidth =2
                    Left =15
                    Top =845
                    Width =10165
                    Name ="Line26"
                End
                Begin Line
                    BorderWidth =2
                    Left =15
                    Top =1670
                    Width =10165
                    Name ="Line27"
                End
                Begin Line
                    Left =9555
                    Top =863
                    Width =0
                    Height =794
                    Name ="Line28"
                End
                Begin Line
                    Left =8910
                    Top =866
                    Width =0
                    Height =794
                    Name ="Line29"
                End
                Begin Line
                    Left =6495
                    Top =860
                    Width =0
                    Height =794
                    Name ="Line32"
                End
                Begin Line
                    Left =4155
                    Top =875
                    Width =0
                    Height =794
                    Name ="Line33"
                End
                Begin Line
                    Left =3166
                    Top =860
                    Width =0
                    Height =794
                    Name ="Line34"
                End
                Begin Line
                    BorderWidth =2
                    Top =863
                    Width =0
                    Height =794
                    Name ="Line36"
                End
                Begin Line
                    BorderWidth =2
                    Left =10140
                    Top =863
                    Width =0
                    Height =794
                    Name ="Line37"
                End
                Begin Label
                    BackStyle =1
                    TextAlign =2
                    TextFontFamily =2
                    Left =3750
                    Width =2385
                    Height =540
                    FontSize =14
                    Name ="Label22"
                    Caption ="فروش روزانه  "
                    FontName ="Titr"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =75
                    Top =442
                    Width =1021
                    Height =360
                    FontSize =10
                    Name ="Text52"
                    ControlSource ="=[Forms]![F_MENU_DATE]![DT2]"
                    FontName ="Zar"
                    InputMask ="####/##/##"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1827
                    Top =440
                    Width =1086
                    Height =360
                    FontSize =10
                    TabIndex =1
                    Name ="Text53"
                    ControlSource ="=[Forms]![F_MENU_DATE]![DT1]"
                    FontName ="Zar"
                    InputMask ="####/##/##"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =1144
                    Top =436
                    Width =660
                    Height =360
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="Label54"
                    Caption ="تا تاريخ:"
                    FontName ="Zar"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =2947
                    Top =438
                    Width =690
                    Height =360
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="Label55"
                    Caption ="از تاريخ:"
                    FontName ="Zar"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =5850
                    Top =860
                    Width =690
                    Height =765
                    FontSize =10
                    ForeColor =0
                    Name ="Label72"
                    Caption ="شماره فاكتور"
                    FontName ="Traffic"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    Left =5835
                    Top =866
                    Width =0
                    Height =794
                    Name ="Line73"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =4185
                    Top =860
                    Width =1020
                    Height =765
                    FontSize =10
                    ForeColor =0
                    Name ="Label76"
                    Caption ="تاريخ فاكتور"
                    FontName ="Traffic"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    Left =5235
                    Top =866
                    Width =0
                    Height =794
                    Name ="Line77"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =45
                    Top =853
                    Width =1356
                    Height =765
                    FontSize =10
                    ForeColor =0
                    Name ="Label80"
                    Caption ="مبلغ ارزي"
                    FontName ="Traffic"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    Left =1456
                    Top =860
                    Width =0
                    Height =794
                    Name ="Line81"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =1515
                    Top =853
                    Width =801
                    Height =765
                    FontSize =10
                    ForeColor =0
                    Name ="Label92"
                    Caption ="متراژ"
                    FontName ="Traffic"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    Left =2341
                    Top =860
                    Width =0
                    Height =794
                    Name ="Line93"
                End
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Left =7370
                    Top =113
                    TabIndex =2
                    Name ="SADER"
                    ControlSource ="SADER"
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =419
            Name ="Detail"
            Begin
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =6539
                    Width =2310
                    Height =375
                    ColumnWidth =4650
                    FontSize =10
                    Name ="SHARH"
                    ControlSource ="NAME"
                    FontName ="Zar"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5267
                    Width =546
                    Height =375
                    FontSize =10
                    TabIndex =1
                    Name ="NUMBER"
                    ControlSource ="NUMBER"
                    Format ="Standard"
                    FontName ="Zar"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3197
                    Width =922
                    Height =375
                    FontSize =10
                    TabIndex =2
                    Name ="DATE_Nh"
                    ControlSource ="DATE_Nh"
                    FontName ="Zar"
                    InputMask ="####/##/##"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2372
                    Width =756
                    Height =375
                    FontSize =11
                    TabIndex =3
                    Name ="MEGHs"
                    ControlSource ="MEGHs"
                    Format ="General Number"
                    FontName ="Zar"
                End
                Begin TextBox
                    RunningSum =1
                    TextAlign =2
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =9576
                    Width =546
                    Height =375
                    FontSize =10
                    TabIndex =4
                    Name ="N_S"
                    ControlSource ="=1"
                    FontName ="Zar"
                End
                Begin Line
                    Left =9555
                    Width =0
                    Height =419
                    Name ="Line38"
                End
                Begin Line
                    Left =8910
                    Width =0
                    Height =419
                    Name ="Line39"
                End
                Begin Line
                    Left =6495
                    Width =0
                    Height =419
                    Name ="Line42"
                End
                Begin Line
                    Left =4155
                    Width =0
                    Height =419
                    Name ="Line43"
                End
                Begin Line
                    Left =3166
                    Width =0
                    Height =419
                    Name ="Line44"
                End
                Begin Line
                    BorderWidth =2
                    Width =0
                    Height =419
                    Name ="Line46"
                End
                Begin Line
                    BorderWidth =2
                    Left =10140
                    Width =0
                    Height =419
                    Name ="Line47"
                End
                Begin Line
                    LineSlant = NotDefault
                    Top =419
                    Width =10178
                    BorderColor =4210752
                    Name ="Line48"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =10
                    ReadingOrder =1
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =8895
                    Width =651
                    Height =375
                    FontSize =8
                    TabIndex =5
                    Name ="hes"
                    ControlSource ="hes"
                    FontName ="Zar"
                End
                Begin Line
                    Left =5835
                    Width =0
                    Height =419
                    Name ="Line74"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5820
                    Width =651
                    Height =375
                    FontSize =10
                    TabIndex =6
                    Name ="ANBARF"
                    ControlSource ="ANBARF"
                    FontName ="Zar"
                End
                Begin Line
                    Left =5235
                    Width =0
                    Height =419
                    Name ="Line78"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4200
                    Width =996
                    Height =375
                    FontSize =10
                    TabIndex =7
                    Name ="Text79"
                    ControlSource ="DATE_N"
                    FontName ="Zar"
                    InputMask ="####/##/##"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =47
                    Width =1356
                    Height =375
                    FontSize =11
                    TabIndex =8
                    Name ="TOTALARZs"
                    ControlSource ="TOTALARZs"
                    Format ="General Number"
                    FontName ="Zar"
                End
                Begin Line
                    Left =1456
                    Width =0
                    Height =419
                    Name ="Line83"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1470
                    Width =861
                    Height =390
                    FontSize =10
                    TabIndex =9
                    Name ="MEGHks"
                    ControlSource ="MEGHks"
                    Format ="General Number"
                    FontName ="Zar"
                End
                Begin Line
                    Left =2341
                    Width =0
                    Height =419
                    Name ="Line95"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    RunningSum =2
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5851
                    Width =1356
                    Height =375
                    FontSize =10
                    TabIndex =10
                    Name ="TOTALARZss"
                    ControlSource ="TOTALARZs"
                    Format ="Standard"
                    FontName ="Zar"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    RunningSum =2
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =8395
                    Width =756
                    Height =375
                    FontSize =8
                    TabIndex =11
                    Name ="MEGHss"
                    ControlSource ="MEGHs"
                    Format ="Standard"
                    FontName ="Zar"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    RunningSum =2
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =7540
                    Width =801
                    Height =375
                    FontSize =8
                    TabIndex =12
                    Name ="MEGHkss"
                    ControlSource ="MEGHks"
                    Format ="Standard"
                    FontName ="Zar"
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =958
            Name ="GroupFooter1"
            Begin
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5328
                    Top =623
                    Width =4421
                    Height =300
                    FontSize =10
                    Name ="Text23"
                    ControlSource ="=farsidate(Now())"
                    FontName ="Zar"
                    InputMask ="####/##/##"
                End
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =113
                    Top =566
                    Width =4421
                    Height =300
                    FontSize =10
                    TabIndex =1
                    Name ="Text24"
                    ControlSource ="=\"صفحه \" & [Page] & \" از \" & [Pages]"
                    FontName ="Zar"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2373
                    Width =756
                    Height =375
                    FontSize =8
                    TabIndex =2
                    Name ="SBED"
                    ControlSource ="=[MEGHss]"
                    Format ="General Number"
                    FontName ="Zar"
                    Tag ="=Sum([SumOfMABL_K]+[SumOfMABL_HAZ]-[SumOfTAKHFIF])"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1470
                    Width =862
                    Height =390
                    FontSize =8
                    TabIndex =3
                    Name ="SBES"
                    ControlSource ="=[MEGHkss]"
                    Format ="General Number"
                    FontName ="Zar"
                End
                Begin Line
                    Left =3167
                    Width =0
                    Height =419
                    Name ="Line60"
                End
                Begin Line
                    Left =2341
                    Width =0
                    Height =419
                    Name ="Line61"
                End
                Begin Line
                    BorderWidth =2
                    Width =0
                    Height =419
                    Name ="Line64"
                End
                Begin Line
                    LineSlant = NotDefault
                    Top =419
                    Width =10178
                    BorderColor =4210752
                    Name ="Line65"
                End
                Begin Line
                    BorderWidth =2
                    Left =10141
                    Width =0
                    Height =419
                    Name ="Line67"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =6066
                    Width =645
                    Height =375
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="Label68"
                    Caption ="جمع كل:"
                    FontName ="Zar"
                    Tag ="DetachedLabel"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =61
                    Width =1371
                    Height =375
                    FontSize =11
                    TabIndex =4
                    Name ="Text84"
                    ControlSource ="=[TOTALARZss]"
                    Format ="General Number"
                    FontName ="Zar"
                    Tag ="=Sum([SumOfMABL_HAV]+[SumOfMABL_VAR])"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =1457
                    Width =0
                    Height =419
                    Name ="Line85"
                End
                Begin Line
                    BorderWidth =2
                    Width =10165
                    Name ="Line86"
                End
            End
        End
        Begin PageFooter
            Height =0
            Name ="PageFooter"
        End
    End
End
CodeBehindForm
' See "R_FROOSH_DAYLY1_sader.cls"
