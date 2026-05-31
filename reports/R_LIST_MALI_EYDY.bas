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
    Width =14629
    DatasheetFontHeight =10
    ItemSuffix =235
    Left =600
    Top =330
    DatasheetGridlinesColor =12632256
    Tag ="@Forms___Baseknow___SAGHFH=[Forms]![Baseknow]![SAGHFH],@Forms___F_MENU_BIM___MMO"
        "=Forms![F_MENU_BIM]![MMO] ,@Forms___Baseknow___HOLA=Forms![Baseknow]![HOLA] ,@Fo"
        "rms___Baseknow___HSAY=Forms![Baseknow]![HSAY] ,@Forms___Baseknow___HJAZ=Forms![B"
        "aseknow]![HJAZ] ,@Forms___Baseknow___HNAH=Forms![Baseknow]![HNAH] ,@Forms___Base"
        "know___HCON=Forms![Baseknow]![HCON] ,@Forms___Baseknow___HKHA=Forms![Baseknow]!["
        "HKHA],@HOKMDATE=Forms![F_MENU_BIM]![MODATE],@Forms___Baseknow___HSANP=Forms![Bas"
        "eknow]![SANAVP]"
    RecSrcDt = Begin
        0x94b281ab3172e340
    End
    RecordSource ="SELECT PERSONEL.CODE, PERSONEL.BIMEH_NUM, PERSONEL.PNAME + N' ' + PERSONEL.PFAMI"
        "LY AS PNPF, PERSONEL.PFAMILY, PERSONEL.JOB, SALARY_EYDY.ROOZ, SALARY_EYDY.DARAMA"
        "D, SALARY_EYDY.MALIAT, SALARY_EYDY.EYDY, SALARY_EYDY.SANAVAT, SALARY_EYDY.MOAFIY"
        "AT, SALARY_EYDY.DARMASH, SALARY_EYDY.MALIKOL, SALARY_EYDY.EPM, SALARY_EYDY.ROOZ "
        "+ SALARY_EYDY.EPM AS ROZHA, SALARY_EYDY.EYDYMAL, SALARY_EYDY.EYDY + SALARY_EYDY."
        "DARAMAD AS JDAR FROM SALARY_EYDY INNER JOIN PERSONEL ON SALARY_EYDY.CODE = PERSO"
        "NEL.CODE WHERE (PERSONEL.BIMEH_NUM > 50)"
    Caption ="ليست عيدي پاداش"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Arial (Arabic)"
    RecordSourceQualifier ="dbo"
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
            Height =1882
            Name ="PageHeader"
            Begin
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =14139
                    Top =997
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
                    Left =11630
                    Top =1000
                    Width =2431
                    Height =810
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="SHARH Label"
                    Caption ="نام و نام خانوادگي"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="SHARH_Label"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =2
                    Left =30
                    Top =885
                    Width =14569
                    BorderColor =0
                    Name ="Line26"
                End
                Begin Line
                    BorderWidth =2
                    Top =1845
                    Width =14600
                    BorderColor =0
                    Name ="Line27"
                End
                Begin Line
                    Left =14144
                    Top =889
                    Width =0
                    Height =993
                    Name ="Line29"
                End
                Begin Line
                    BorderWidth =2
                    Left =14597
                    Top =889
                    Width =0
                    Height =993
                    BorderColor =0
                    Name ="Line37"
                End
                Begin Label
                    TextFontFamily =2
                    Left =5610
                    Top =390
                    Width =3030
                    Height =390
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="Label22"
                    Caption ="ليست مجموع درآمد سالانه و عيدي پاداش"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =10695
                    Top =1005
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
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =9900
                    Top =955
                    Width =585
                    Height =885
                    FontSize =7
                    FontWeight =400
                    ForeColor =0
                    Name ="Label85"
                    Caption ="مدت كاركرد در سال"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =8700
                    Top =915
                    Width =1125
                    Height =915
                    FontSize =7
                    FontWeight =400
                    ForeColor =0
                    Name ="Label86"
                    Caption ="جمع حقوق دوازده ماهه طبق ليستهاي ارسالي"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    Left =9851
                    Top =889
                    Width =0
                    Height =993
                    Name ="Line88"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =5185
                    Top =390
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
                    TextAlign =3
                    TextFontFamily =2
                    NumeralShapes =2
                    Left =4440
                    Top =390
                    Width =741
                    Height =405
                    FontSize =10
                    Name ="Text117"
                    ControlSource ="=[Forms]![BASEKNOW]![YEA]"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =7470
                    Top =915
                    Width =1140
                    Height =915
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="Label118"
                    Caption =" عيدي و پاداش آخر سال"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    Left =8651
                    Top =889
                    Width =0
                    Height =993
                    Name ="Line119"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =6288
                    Top =915
                    Width =1110
                    Height =870
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="Label122"
                    Caption ="جمع كل درآمد"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    Left =7428
                    Top =889
                    Width =0
                    Height =993
                    Name ="Line123"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =4983
                    Top =915
                    Width =1200
                    Height =870
                    FontSize =7
                    FontWeight =400
                    ForeColor =0
                    Name ="Label126"
                    Caption ="معافيت ماليات سالانه به نسبت كاركرد"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    Left =6243
                    Top =889
                    Width =0
                    Height =993
                    Name ="Line127"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =3633
                    Top =915
                    Width =1215
                    Height =870
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="Label130"
                    Caption ="درآمد مشمول ماليات"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    Left =4903
                    Top =885
                    Width =0
                    Height =993
                    Name ="Line131"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =2445
                    Top =945
                    Width =1110
                    Height =885
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="Label134"
                    Caption ="ماليات كل سال"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    Left =3600
                    Top =889
                    Width =0
                    Height =993
                    Name ="Line135"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =1264
                    Top =945
                    Width =1095
                    Height =885
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="Label138"
                    Caption ="ماليهاتهاي پرداختي 12 ماهه طبق ليستهاي ارسالي"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    Left =2412
                    Top =889
                    Width =0
                    Height =993
                    Name ="Line139"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =90
                    Top =945
                    Width =1095
                    Height =915
                    FontSize =7
                    FontWeight =400
                    ForeColor =0
                    Name ="Label142"
                    Caption ="تتمه (ماليات عيدي پاداش آخر سال و معوقه ها)"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    Left =1218
                    Top =889
                    Width =0
                    Height =993
                    Name ="Line143"
                End
                Begin Line
                    Left =11606
                    Top =889
                    Width =0
                    Height =993
                    Name ="Line158"
                End
                Begin Line
                    Left =10521
                    Top =885
                    Width =0
                    Height =993
                    Name ="Line179"
                End
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =70
                    Top =453
                    Width =3086
                    Height =300
                    FontSize =10
                    TabIndex =1
                    Name ="Text24"
                    ControlSource ="=\"صفحه \" & [Page] & \" از \" & [Pages]"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =2
                    Left =4198
                    Top =30
                    Width =4836
                    Height =465
                    FontSize =12
                    TabIndex =2
                    Name ="Text187"
                    ControlSource ="=[Forms]![Baseknow]![NAME]"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    BorderWidth =2
                    Top =885
                    Width =0
                    Height =993
                    BorderColor =0
                    Name ="Line206"
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =426
            Name ="Detail"
            Begin
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =50
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =11625
                    Width =2475
                    Height =390
                    ColumnWidth =4650
                    FontSize =8
                    Name ="PNPF"
                    ControlSource ="PNPF"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    RunningSum =1
                    TextAlign =2
                    TextFontFamily =50
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =14171
                    Width =456
                    Height =390
                    FontSize =8
                    TabIndex =1
                    Name ="N_S"
                    ControlSource ="=1"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    Left =14144
                    Width =0
                    Height =426
                    Name ="Line39"
                End
                Begin Line
                    Left =11606
                    Width =0
                    Height =426
                    Name ="Line43"
                End
                Begin Line
                    BorderWidth =2
                    Left =14608
                    Width =0
                    Height =426
                    BorderColor =0
                    Name ="Line47"
                End
                Begin Line
                    Top =420
                    Width =14629
                    BorderColor =4210752
                    Name ="Line48"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =2
                    TextFontFamily =50
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =10590
                    Width =996
                    Height =390
                    FontSize =8
                    TabIndex =2
                    Name ="JOB"
                    ControlSource ="JOB"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =2
                    TextFontFamily =50
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =9850
                    Width =637
                    Height =390
                    FontSize =8
                    TabIndex =3
                    Name ="ROZHA"
                    ControlSource ="ROZHA"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =2
                    TextFontFamily =50
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =8695
                    Width =1117
                    Height =390
                    FontSize =8
                    TabIndex =4
                    Name ="DARAMAD"
                    ControlSource ="DARAMAD"
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    Left =10532
                    Width =0
                    Height =426
                    Name ="Line95"
                End
                Begin Line
                    Left =9851
                    Width =0
                    Height =426
                    Name ="Line96"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =2
                    TextFontFamily =50
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =7486
                    Width =1116
                    Height =390
                    FontSize =8
                    TabIndex =5
                    Name ="EYDY"
                    ControlSource ="EYDY"
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    Left =8649
                    Width =0
                    Height =426
                    Name ="Line121"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =2
                    TextFontFamily =50
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =6270
                    Width =1131
                    Height =390
                    FontSize =8
                    TabIndex =6
                    Name ="JDAR"
                    ControlSource ="=[EYDY]+[DARAMAD]"
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    Left =7424
                    Width =0
                    Height =426
                    Name ="Line125"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =2
                    TextFontFamily =50
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4965
                    Width =1221
                    Height =390
                    FontSize =8
                    TabIndex =7
                    Name ="MOAFIYAT"
                    ControlSource ="MOAFIYAT"
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    Left =6240
                    Width =0
                    Height =426
                    Name ="Line129"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =2
                    TextFontFamily =50
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3630
                    Width =1236
                    Height =390
                    FontSize =8
                    TabIndex =8
                    Name ="DARMASH"
                    ControlSource ="DARMASH"
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    Left =4899
                    Width =0
                    Height =426
                    Name ="Line133"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =2
                    TextFontFamily =50
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2433
                    Width =1116
                    Height =390
                    FontSize =8
                    TabIndex =9
                    Name ="MALIKOL"
                    ControlSource ="MALIKOL"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    Left =3596
                    Width =0
                    Height =426
                    Name ="Line137"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =2
                    TextFontFamily =50
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1245
                    Width =1131
                    Height =390
                    FontSize =8
                    TabIndex =10
                    Name ="MALIAT"
                    ControlSource ="MALIAT"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    Left =2408
                    Width =0
                    Height =426
                    Name ="Line141"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =2
                    TextFontFamily =50
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =75
                    Width =1101
                    Height =390
                    FontSize =8
                    TabIndex =11
                    Name ="EYDYMAL"
                    ControlSource ="EYDYMAL"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    Left =1213
                    Width =0
                    Height =426
                    Name ="Line145"
                End
                Begin Line
                    BorderWidth =2
                    Width =0
                    Height =426
                    BorderColor =0
                    Name ="Line207"
                End
            End
        End
        Begin PageFooter
            Height =407
            OnFormat ="[Event Procedure]"
            Name ="PageFooter"
            Begin
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =1975
                    Width =1755
                    Height =375
                    FontSize =10
                    FontWeight =400
                    BorderColor =4210752
                    ForeColor =0
                    Name ="Label231"
                    Caption ="مهر شركت و امضاء "
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin TextBox
                    Visible = NotDefault
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =947
                    Top =16
                    Width =612
                    Height =347
                    Name ="Text233"
                    ControlSource ="=[Page]"
                    FontName ="WeblogmaYekan"
                End
            End
        End
        Begin FormFooter
            KeepTogether = NotDefault
            Height =810
            OnFormat ="[Event Procedure]"
            Name ="ReportFooter"
            Begin
                Begin Line
                    Left =11
                    Top =405
                    Width =14614
                    BorderColor =4210752
                    Name ="Line65"
                End
                Begin Line
                    BorderWidth =2
                    Left =14608
                    Width =0
                    Height =419
                    BorderColor =4210752
                    Name ="Line67"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =50
                    Left =12525
                    Width =1380
                    Height =375
                    FontSize =8
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
                    Width =14599
                    Name ="Line71"
                End
                Begin Line
                    BorderWidth =2
                    Width =0
                    Height =419
                    BorderColor =4210752
                    Name ="Line208"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =2
                    TextFontFamily =50
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =8695
                    Width =1117
                    Height =390
                    FontSize =8
                    Name ="Text215"
                    ControlSource ="=Sum([DARAMAD])"
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    Left =9849
                    Width =0
                    Height =426
                    Name ="Line216"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =2
                    TextFontFamily =50
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =7486
                    Width =1116
                    Height =390
                    FontSize =8
                    TabIndex =1
                    Name ="Text217"
                    ControlSource ="=Sum([EYDY])"
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    Left =8649
                    Width =0
                    Height =426
                    Name ="Line218"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =2
                    TextFontFamily =50
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =6270
                    Width =1131
                    Height =390
                    FontSize =8
                    TabIndex =2
                    Name ="Text219"
                    ControlSource ="=Sum([JDAR])"
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    Left =7424
                    Width =0
                    Height =426
                    Name ="Line220"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =2
                    TextFontFamily =50
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4965
                    Width =1221
                    Height =390
                    FontSize =8
                    TabIndex =3
                    Name ="Text221"
                    ControlSource ="=Sum([MOAFIYAT])"
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    Left =6240
                    Width =0
                    Height =426
                    Name ="Line222"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =2
                    TextFontFamily =50
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3630
                    Width =1236
                    Height =390
                    FontSize =8
                    TabIndex =4
                    Name ="Text223"
                    ControlSource ="=Sum([DARMASH])"
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    Left =4899
                    Width =0
                    Height =426
                    Name ="Line224"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =2
                    TextFontFamily =50
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2433
                    Width =1116
                    Height =390
                    FontSize =8
                    TabIndex =5
                    Name ="Text225"
                    ControlSource ="=Sum([MALIKOL])"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    Left =3596
                    Width =0
                    Height =426
                    Name ="Line226"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =2
                    TextFontFamily =50
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1245
                    Width =1131
                    Height =390
                    FontSize =8
                    TabIndex =6
                    Name ="Text227"
                    ControlSource ="=Sum([MALIAT])"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    Left =2408
                    Width =0
                    Height =426
                    Name ="Line228"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =2
                    TextFontFamily =50
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =60
                    Width =1116
                    Height =390
                    FontSize =8
                    TabIndex =7
                    Name ="Text229"
                    ControlSource ="=Sum([EYDYMAL])"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    Left =1213
                    Width =0
                    Height =426
                    Name ="Line230"
                End
                Begin Label
                    Visible = NotDefault
                    TextAlign =2
                    TextFontFamily =2
                    Left =2100
                    Top =435
                    Width =1755
                    Height =375
                    FontSize =10
                    FontWeight =400
                    BorderColor =4210752
                    ForeColor =0
                    Name ="Label232"
                    Caption ="مهر شركت و امضاء "
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
            End
        End
    End
End
CodeBehindForm
' See "R_LIST_MALI_EYDY.cls"
