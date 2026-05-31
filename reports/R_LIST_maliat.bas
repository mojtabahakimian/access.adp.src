Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
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
    Width =15431
    DatasheetFontHeight =10
    ItemSuffix =221
    Left =600
    Top =1095
    DatasheetGridlinesColor =12632256
    Tag ="@Forms___Baseknow___SAGHFH=[Forms]![Baseknow]![SAGHFH],@Forms___F_MENU_BIM___MMO"
        "=Forms![F_MENU_BIM]![MMO] ,@Forms___Baseknow___HOLA=Forms![Baseknow]![HOLA] ,@Fo"
        "rms___Baseknow___HSAY=Forms![Baseknow]![HSAY] ,@Forms___Baseknow___HJAZ=Forms![B"
        "aseknow]![HJAZ] ,@Forms___Baseknow___HNAH=Forms![Baseknow]![HNAH] ,@Forms___Base"
        "know___HCON=Forms![Baseknow]![HCON] ,@Forms___Baseknow___HKHA=Forms![Baseknow]!["
        "HKHA],@HOKMDATE=Forms![F_MENU_BIM]![MODATE],@Forms___Baseknow___HSANP=Forms![Bas"
        "eknow]![SANAVP]"
    RecSrcDt = Begin
        0x9530ac585423e640
    End
    RecordSource ="SELECT PLIST_MALIAT.WDATE, PLIST_MALIAT.CODE, PLIST_MALIAT.DAYS, PLIST_MALIAT.SA"
        "LARY_DAYLY, PLIST_MALIAT.DAYS * PLIST_MALIAT.SALARY_DAYLY AS MOSALARY, PLIST_MAL"
        "IAT.HOME, PLIST_MALIAT.CHILDREN, PLIST_MALIAT.CONDITIONS, PLIST_MALIAT.BON, PLIS"
        "T_MALIAT.JAZB, PLIST_MALIAT.SAYER, PLIST_MALIAT.EZAFAH, PLIST_MALIAT.PADASH, PLI"
        "ST_MALIAT.MAZAYA, PLIST_MALIAT.MAZAYAKOL, PLIST_MALIAT.MALIAT, PERSONEL.BIMEH_NU"
        "M, PERSONEL.PNAME + N' ' + PERSONEL.PFAMILY AS NAM, PERSONEL.PFAMILY, PERSONEL.K"
        "HNOWNUM, PERSONEL.FATHER, PERSONEL.SEX, PERSONEL.WSDATE, PERSONEL.WEDATE, PERSON"
        "EL.JOB, PLIST_MALIAT.MAZAYAKOL AS MAZMASHVAGH, PLIST_MALIAT.BIMAH, dbo.Umonth(PL"
        "IST_MALIAT.WDATE) AS mm, PERSONEL.MELLICOD FROM PLIST_MALIAT INNER JOIN PERSONEL"
        " ON PLIST_MALIAT.CODE = PERSONEL.CODE WHERE (PERSONEL.BIMEH_NUM > 0)"
    Caption ="ليست بيمه و ماليات پرسنل"
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
        Begin Tab
            TextFontCharSet =178
            TextFontFamily =2
            Width =5103
            Height =3402
            FontName ="Arial (Arabic)"
        End
        Begin BreakLevel
            ControlSource ="CODE"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =0
            Name ="ReportHeader"
        End
        Begin PageHeader
            Height =1732
            Name ="PageHeader"
            Begin
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =14930
                    Top =847
                    Width =450
                    Height =810
                    FontSize =7
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
                    Left =12031
                    Top =850
                    Width =2851
                    Height =810
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="SHARH Label"
                    Caption ="نام و نام خانوادگي"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="SHARH_Label"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =11370
                    Top =825
                    Width =600
                    Height =810
                    FontSize =7
                    FontWeight =400
                    ForeColor =0
                    Name ="BES Label"
                    Caption ="شماره شناسنامه"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="BES_Label"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =10718
                    Top =825
                    Width =480
                    Height =300
                    FontSize =7
                    FontWeight =400
                    ForeColor =0
                    Name ="Expr1 Label"
                    Caption ="نام پدر"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="Expr1_Label"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =2
                    Left =17
                    Top =735
                    Width =15349
                    BorderColor =0
                    Name ="Line26"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =2
                    Left =17
                    Top =1695
                    Width =15365
                    BorderColor =0
                    Name ="Line27"
                End
                Begin Line
                    Left =14935
                    Top =739
                    Width =0
                    Height =993
                    Name ="Line29"
                End
                Begin Line
                    Left =11325
                    Top =735
                    Width =0
                    Height =993
                    Name ="Line34"
                End
                Begin Line
                    BorderWidth =2
                    Left =15388
                    Top =739
                    Width =0
                    Height =993
                    BorderColor =0
                    Name ="Line37"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =8413
                    Width =4575
                    Height =390
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="Label22"
                    Caption ="صورت دستمزد/حقوق ومزاياي ماهانه                            ماه"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =9383
                    Top =795
                    Width =1065
                    Height =405
                    FontSize =7
                    FontWeight =400
                    ForeColor =0
                    Name ="Label72"
                    Caption ="شماره بيمه شده"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =8572
                    Top =795
                    Width =600
                    Height =405
                    FontSize =7
                    FontWeight =400
                    ForeColor =0
                    Name ="Label73"
                    Caption ="شغل"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    Left =9379
                    Top =735
                    Width =0
                    Height =974
                    Name ="Line75"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =8095
                    Top =915
                    Width =240
                    Height =315
                    FontSize =7
                    FontWeight =400
                    ForeColor =0
                    Name ="Label82"
                    Caption ="مرد"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =7890
                    Top =915
                    Width =210
                    Height =405
                    FontSize =7
                    FontWeight =400
                    ForeColor =0
                    Name ="Label83"
                    Caption ="زن"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    Left =8116
                    Top =735
                    Width =0
                    Height =978
                    Name ="Line84"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =6898
                    Width =1521
                    Height =405
                    ColumnWidth =1110
                    FontSize =10
                    Name ="WDATE1"
                    ControlSource ="=UTMONTH([WDATE])"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =6458
                    Width =405
                    Height =390
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="Label116"
                    Caption ="سال"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5685
                    Width =741
                    Height =405
                    FontSize =10
                    TabIndex =1
                    Name ="Text117"
                    ControlSource ="=uyear([WDATE])"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =7425
                    Top =765
                    Width =405
                    Height =870
                    FontSize =7
                    FontWeight =400
                    ForeColor =0
                    Name ="Label118"
                    Caption ="روزهاي كاركرد"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    Left =7836
                    Top =739
                    Width =0
                    Height =993
                    Name ="Line119"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =6831
                    Top =765
                    Width =540
                    Height =870
                    FontSize =7
                    FontWeight =400
                    ForeColor =0
                    Name ="Label122"
                    Caption ="دستمزد روزانه"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    Left =7401
                    Top =739
                    Width =0
                    Height =993
                    Name ="Line123"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =5961
                    Top =765
                    Width =810
                    Height =870
                    FontSize =7
                    FontWeight =400
                    ForeColor =0
                    Name ="Label126"
                    Caption ="دستمزد ماهانه"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    Left =6801
                    Top =739
                    Width =0
                    Height =993
                    Name ="Line127"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =5077
                    Top =765
                    Width =825
                    Height =870
                    FontSize =7
                    FontWeight =400
                    ForeColor =0
                    Name ="Label130"
                    Caption ="مزاياي ماهانه مشمول بيمه"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    Left =5942
                    Top =735
                    Width =0
                    Height =993
                    Name ="Line131"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =4117
                    Top =795
                    Width =885
                    Height =885
                    FontSize =7
                    FontWeight =400
                    ForeColor =0
                    Name ="Label134"
                    Caption ="جمع دستمزد و مزاياي ماهانه مشمول  بيمه"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    Left =5032
                    Top =739
                    Width =0
                    Height =993
                    Name ="Line135"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =3195
                    Top =795
                    Width =870
                    Height =915
                    FontSize =7
                    FontWeight =400
                    ForeColor =0
                    Name ="Label138"
                    Caption ="جمع دستمزد و مزاياي ماهانه مشمول وعير مشمول بيمه"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    Left =4102
                    Top =739
                    Width =0
                    Height =993
                    Name ="Line139"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =2534
                    Top =795
                    Width =600
                    Height =885
                    FontSize =7
                    FontWeight =400
                    ForeColor =0
                    Name ="Label142"
                    Caption ="حق بيمه سهم بيمه شده"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    Left =3164
                    Top =739
                    Width =0
                    Height =993
                    Name ="Line143"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =1378
                    Top =795
                    Width =1080
                    Height =885
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="Label146"
                    Caption ="ماليات"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    Left =2489
                    Top =739
                    Width =0
                    Height =993
                    Name ="Line147"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =52
                    Top =795
                    Width =1245
                    Height =885
                    FontSize =7
                    FontWeight =400
                    ForeColor =0
                    Name ="Label150"
                    Caption ="ملاحطات"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    Left =1337
                    Top =739
                    Width =0
                    Height =993
                    Name ="Line151"
                End
                Begin Line
                    Left =12007
                    Top =739
                    Width =0
                    Height =993
                    Name ="Line158"
                End
                Begin Line
                    Left =8385
                    Top =735
                    Width =0
                    Height =978
                    Name ="Line178"
                End
                Begin Line
                    Left =10510
                    Top =735
                    Width =0
                    Height =993
                    Name ="Line180"
                End
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =236
                    Top =60
                    Width =4226
                    Height =300
                    FontSize =10
                    TabIndex =2
                    Name ="Text24"
                    ControlSource ="=\"صفحه \" & [Page] & \" از \" & [Pages]"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextFontFamily =2
                    Left =14578
                    Top =375
                    Width =840
                    Height =345
                    FontWeight =400
                    ForeColor =0
                    Name ="Label184"
                    Caption ="شماره كارگاه:"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    NumeralShapes =2
                    Left =12240
                    Top =345
                    Width =2316
                    Height =330
                    TabIndex =3
                    Name ="Text185"
                    ControlSource ="=[Forms]![Baseknow]![UNIVERSITY_CO]"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextFontFamily =2
                    Left =11458
                    Top =375
                    Width =660
                    Height =345
                    FontWeight =400
                    ForeColor =0
                    Name ="Label186"
                    Caption ="نام كارگاه:"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    Left =8883
                    Top =380
                    Width =2556
                    Height =345
                    TabIndex =4
                    Name ="Text187"
                    ControlSource ="=[Forms]![Baseknow]![NAME]"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextFontFamily =2
                    Left =8068
                    Top =375
                    Width =810
                    Height =345
                    FontWeight =400
                    ForeColor =0
                    Name ="Label190"
                    Caption ="نام كار فرما:"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    Left =6618
                    Top =380
                    Width =1431
                    Height =345
                    TabIndex =5
                    Name ="Text191"
                    ControlSource ="=[Forms]![Baseknow]![MANAGER]"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextFontFamily =2
                    Left =5713
                    Top =375
                    Width =870
                    Height =345
                    FontWeight =400
                    ForeColor =0
                    Name ="Label192"
                    Caption ="نشاني كارگاه:"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    Left =2224
                    Top =380
                    Width =3426
                    Height =345
                    TabIndex =6
                    Name ="Text193"
                    ControlSource ="=[Forms]![Baseknow]![TFADDRESS]"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextFontFamily =2
                    Left =1031
                    Top =375
                    Width =1650
                    Height =315
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="Label194"
                    Caption ="نام شعبه تامين اجتماعي:"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    Left =71
                    Top =390
                    Width =921
                    Height =345
                    TabIndex =7
                    Name ="Text195"
                    ControlSource ="=[Forms]![Baseknow]![AMINAMVAL]"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    BorderWidth =2
                    Left =11
                    Top =735
                    Width =0
                    Height =993
                    BorderColor =0
                    Name ="Line206"
                End
                Begin TextBox
                    Visible = NotDefault
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =3915
                    TabIndex =8
                    Name ="WDATE"
                    ControlSource ="WDATE"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =2099
                    Top =1142
                    TabIndex =9
                    Name ="mm"
                    ControlSource ="mm"
                    FontName ="WeblogmaYekan"
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =435
            Name ="Detail"
            Begin
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =12015
                    Width =2895
                    Height =390
                    ColumnWidth =4650
                    FontSize =8
                    Name ="NAM"
                    ControlSource ="NAM"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =3
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =11340
                    Width =652
                    Height =390
                    FontSize =8
                    TabIndex =1
                    Name ="KHNOWNUM"
                    ControlSource ="KHNOWNUM"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =10533
                    Width =771
                    Height =390
                    FontSize =8
                    TabIndex =2
                    Name ="FATHER"
                    ControlSource ="FATHER"
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
                    Left =14951
                    Width =456
                    Height =390
                    FontSize =8
                    TabIndex =3
                    Name ="N_S"
                    ControlSource ="=1"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    Left =14935
                    Width =0
                    Height =426
                    Name ="Line39"
                End
                Begin Line
                    Left =12007
                    Width =0
                    Height =426
                    Name ="Line43"
                End
                Begin Line
                    Left =11325
                    Width =0
                    Height =426
                    Name ="Line44"
                End
                Begin Line
                    BorderWidth =2
                    Left =15388
                    Width =0
                    Height =426
                    BorderColor =0
                    Name ="Line47"
                End
                Begin Line
                    Top =420
                    Width =15409
                    BorderColor =4210752
                    Name ="Line48"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =9427
                    Width =1027
                    Height =435
                    FontSize =8
                    TabIndex =4
                    Name ="BIMEH_NUM"
                    ControlSource ="BIMEH_NUM"
                    Format ="#"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =8407
                    Width =996
                    Height =390
                    FontSize =8
                    TabIndex =5
                    Name ="JOB"
                    ControlSource ="JOB"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    Left =10510
                    Width =0
                    Height =426
                    Name ="Line79"
                End
                Begin Line
                    Left =9381
                    Width =0
                    Height =426
                    Name ="Line80"
                End
                Begin Line
                    Left =8379
                    Width =0
                    Height =426
                    Name ="Line81"
                End
                Begin Line
                    Left =8110
                    Width =0
                    Height =426
                    Name ="Line92"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =7440
                    Width =381
                    Height =390
                    FontSize =8
                    TabIndex =6
                    Name ="DAYS"
                    ControlSource ="DAYS"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    Left =7838
                    Width =0
                    Height =426
                    Name ="Line121"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =6817
                    Width =561
                    Height =390
                    FontSize =7
                    TabIndex =7
                    Name ="SALARY_DAYLY"
                    ControlSource ="SALARY_DAYLY"
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    Left =7401
                    Width =0
                    Height =426
                    Name ="Line125"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5947
                    Width =831
                    Height =390
                    FontSize =8
                    TabIndex =8
                    Name ="mosalary"
                    ControlSource ="mosalary"
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    Left =6802
                    Width =0
                    Height =426
                    Name ="Line129"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5048
                    Width =876
                    Height =390
                    FontSize =8
                    TabIndex =9
                    Name ="MAZAYA"
                    ControlSource ="MAZAYA"
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    Left =5942
                    Width =0
                    Height =426
                    Name ="Line133"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4109
                    Width =891
                    Height =390
                    FontSize =8
                    TabIndex =10
                    Name ="MAZAYAKOL"
                    ControlSource ="MAZAYAKOL"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    Left =5032
                    Width =0
                    Height =426
                    Name ="Line137"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3179
                    Width =906
                    Height =390
                    FontSize =8
                    TabIndex =11
                    Name ="MAZMASHVAGH"
                    ControlSource ="MAZMASHVAGH"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    Left =4102
                    Width =0
                    Height =426
                    Name ="Line141"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2505
                    Width =621
                    Height =390
                    FontSize =7
                    TabIndex =12
                    Name ="BIMAH"
                    ControlSource ="BIMAH"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    Left =3163
                    Width =0
                    Height =426
                    Name ="Line145"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1364
                    Width =1101
                    Height =390
                    FontSize =8
                    TabIndex =13
                    Name ="MALIAT"
                    ControlSource ="MALIAT"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    Left =2488
                    Width =0
                    Height =426
                    Name ="Line149"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =50
                    Width =1236
                    Height =390
                    FontSize =8
                    TabIndex =14
                    Name ="ghabel"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    Left =1335
                    Width =0
                    Height =426
                    Name ="Line153"
                End
                Begin CheckBox
                    OldBorderStyle =0
                    ReadingOrder =1
                    Left =8101
                    Top =105
                    Height =285
                    TabIndex =15
                    Name ="Check159"
                    ControlSource ="SEX"
                End
                Begin CheckBox
                    OldBorderStyle =0
                    Left =7921
                    Top =105
                    Height =285
                    TabIndex =16
                    Name ="Check161"
                    ControlSource ="=Not [SEX]"
                End
                Begin Line
                    BorderWidth =2
                    Left =11
                    Width =0
                    Height =426
                    BorderColor =0
                    Name ="Line207"
                End
            End
        End
        Begin PageFooter
            Height =0
            Name ="PageFooter"
        End
        Begin FormFooter
            KeepTogether = NotDefault
            CanGrow = NotDefault
            Height =2116
            Name ="ReportFooter"
            Begin
                Begin Line
                    Left =6802
                    Width =0
                    Height =599
                    BorderColor =4210752
                    Name ="Line60"
                End
                Begin Line
                    Left =15
                    Top =585
                    Width =15394
                    BorderColor =4210752
                    Name ="Line65"
                End
                Begin Line
                    BorderWidth =2
                    Left =15388
                    Width =0
                    Height =599
                    BorderColor =4210752
                    Name ="Line67"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =7559
                    Width =1380
                    Height =375
                    FontSize =7
                    FontWeight =400
                    BorderColor =4210752
                    ForeColor =0
                    Name ="Label68"
                    Caption ="جمع كل:"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =2
                    Left =11
                    Width =15379
                    Name ="Line71"
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    DecimalPlaces =0
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =6824
                    Top =60
                    Width =576
                    Height =480
                    FontSize =7
                    Name ="Text162"
                    ControlSource ="=Sum([SALARY_DAYLY])"
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    DecimalPlaces =0
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5954
                    Top =60
                    Width =846
                    Height =480
                    FontSize =7
                    TabIndex =1
                    Name ="Text163"
                    ControlSource ="=Sum([mosalary])"
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    Left =7402
                    Width =0
                    Height =599
                    BorderColor =4210752
                    Name ="Line164"
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    DecimalPlaces =0
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5054
                    Top =60
                    Width =876
                    Height =480
                    FontSize =7
                    TabIndex =2
                    Name ="Text165"
                    ControlSource ="=Sum([MAZAYA])"
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    Left =5942
                    Width =0
                    Height =599
                    BorderColor =4210752
                    Name ="Line166"
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    DecimalPlaces =0
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4109
                    Top =60
                    Width =921
                    Height =495
                    FontSize =7
                    TabIndex =3
                    Name ="Text167"
                    ControlSource ="=Sum([MAZAYAKOL])"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =5040
                    Width =0
                    Height =599
                    BorderColor =4210752
                    Name ="Line168"
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    DecimalPlaces =0
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3194
                    Top =60
                    Width =891
                    Height =495
                    FontSize =7
                    TabIndex =4
                    Name ="Text169"
                    ControlSource ="=Sum([MAZMASHVAGH])"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    Left =4102
                    Width =0
                    Height =599
                    BorderColor =4210752
                    Name ="Line170"
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    DecimalPlaces =0
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2490
                    Top =60
                    Width =651
                    Height =480
                    FontSize =7
                    TabIndex =5
                    Name ="SBIMP"
                    ControlSource ="=Sum([BIMAH])"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    DecimalPlaces =0
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1379
                    Top =60
                    Width =1086
                    Height =480
                    FontSize =7
                    TabIndex =6
                    Name ="Text173"
                    ControlSource ="=Sum([MALIAT])"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    Left =2488
                    Width =0
                    Height =599
                    BorderColor =4210752
                    Name ="Line174"
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    DecimalPlaces =0
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =45
                    Top =60
                    Width =1222
                    Height =495
                    FontSize =7
                    TabIndex =7
                    Name ="Text175"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    BorderWidth =2
                    Left =11
                    Width =0
                    Height =599
                    BorderColor =4210752
                    Name ="Line208"
                End
                Begin Rectangle
                    Left =8865
                    Top =615
                    Width =3345
                    Height =1457
                    Name ="Box210"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =12330
                    Top =615
                    Width =1635
                    Height =360
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="Label211"
                    Caption ="مهر و امضاء كارفرما:"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    Left =1350
                    Width =0
                    Height =606
                    Name ="Line218"
                End
                Begin Line
                    Left =3165
                    Width =0
                    Height =599
                    BorderColor =4210752
                    Name ="Line219"
                End
            End
        End
    End
End
CodeBehindForm
' See "R_LIST_maliat.cls"
