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
    Width =10683
    DatasheetFontHeight =10
    ItemSuffix =241
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
        0x245c79173772e340
    End
    RecordSource ="SELECT PERSONEL.CODE, PERSONEL.BIMEH_NUM, PERSONEL.PNAME + N' ' + PERSONEL.PFAMI"
        "LY AS PNPF, PERSONEL.PFAMILY, PERSONEL.JOB, SALARY_EYDY.ROOZ, SALARY_EYDY.DARAMA"
        "D, SALARY_EYDY.MALIAT, SALARY_EYDY.EYDY, SALARY_EYDY.SANAVAT, SALARY_EYDY.MOAFIY"
        "AT, SALARY_EYDY.DARMASH, SALARY_EYDY.MALIKOL, SALARY_EYDY.EPM, SALARY_EYDY.ROOZ "
        "+ SALARY_EYDY.EPM AS ROZHA, SALARY_EYDY.EYDYMAL, SALARY_EYDY.EYDY + SALARY_EYDY."
        "DARAMAD AS JDAR, SALARY_EYDY.SANAVAT + SALARY_EYDY.EYDY - SALARY_EYDY.EYDYMAL AS"
        " MAND FROM SALARY_EYDY INNER JOIN PERSONEL ON SALARY_EYDY.CODE = PERSONEL.CODE"
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
                    Left =10134
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
                    Left =7625
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
                    Left =6
                    Top =885
                    Width =10594
                    BorderColor =0
                    Name ="Line26"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =2
                    Left =6
                    Top =1845
                    Width =10610
                    BorderColor =0
                    Name ="Line27"
                End
                Begin Line
                    Left =10139
                    Top =889
                    Width =0
                    Height =993
                    Name ="Line29"
                End
                Begin Line
                    BorderWidth =2
                    Left =10592
                    Top =889
                    Width =0
                    Height =993
                    BorderColor =0
                    Name ="Line37"
                End
                Begin Label
                    TextFontFamily =2
                    Left =3870
                    Top =390
                    Width =4755
                    Height =390
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="Label22"
                    Caption ="ليست مجموع درآمد سالانه و عيدي پاداش و حق سنوات"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =6690
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
                    Left =5895
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
                    Left =4695
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
                    Left =5846
                    Top =889
                    Width =0
                    Height =993
                    Name ="Line88"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =3445
                    Top =391
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
                    Left =2700
                    Top =391
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
                    Left =3465
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
                    Left =4646
                    Top =889
                    Width =0
                    Height =993
                    Name ="Line119"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =2283
                    Top =915
                    Width =1110
                    Height =870
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="Label122"
                    Caption ="حق سنوات"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    Left =3423
                    Top =889
                    Width =0
                    Height =993
                    Name ="Line123"
                End
                Begin Line
                    Left =2238
                    Top =889
                    Width =0
                    Height =993
                    Name ="Line127"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =1110
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
                    Left =7601
                    Top =889
                    Width =0
                    Height =993
                    Name ="Line158"
                End
                Begin Line
                    Left =6516
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
                    Left =132
                    Top =451
                    Width =1391
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
                    Left =2863
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
                Begin Line
                    Left =1083
                    Top =889
                    Width =0
                    Height =993
                    Name ="Line235"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =15
                    Top =945
                    Width =1050
                    Height =915
                    FontSize =7
                    FontWeight =400
                    ForeColor =0
                    Name ="Label236"
                    Caption ="مانده قابل پرداخت عيدي پاداش و حق سنوات"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
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
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =7620
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
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =10166
                    Width =456
                    Height =390
                    FontSize =8
                    TabIndex =1
                    Name ="N_S"
                    ControlSource ="=1"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    Left =10150
                    Width =0
                    Height =426
                    Name ="Line39"
                End
                Begin Line
                    Left =7612
                    Width =0
                    Height =426
                    Name ="Line43"
                End
                Begin Line
                    BorderWidth =2
                    Left =10603
                    Width =0
                    Height =426
                    BorderColor =0
                    Name ="Line47"
                End
                Begin Line
                    Top =420
                    Width =10654
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
                    Left =6585
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
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5845
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
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4690
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
                    Left =6527
                    Width =0
                    Height =426
                    Name ="Line95"
                End
                Begin Line
                    Left =5859
                    Width =0
                    Height =426
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
                    Left =3481
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
                    Left =4644
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
                    Left =2265
                    Width =1131
                    Height =390
                    FontSize =8
                    TabIndex =6
                    Name ="SANAVAT"
                    ControlSource ="SANAVAT"
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    Left =3419
                    Width =0
                    Height =426
                    Name ="Line125"
                End
                Begin Line
                    Left =2235
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
                    Left =1095
                    Width =1101
                    Height =390
                    FontSize =8
                    TabIndex =7
                    Name ="EYDYMAL"
                    ControlSource ="EYDYMAL"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    BorderWidth =2
                    Width =0
                    Height =426
                    BorderColor =0
                    Name ="Line207"
                End
                Begin Line
                    Left =1080
                    Width =0
                    Height =426
                    Name ="Line237"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Width =1056
                    Height =390
                    FontSize =8
                    TabIndex =8
                    Name ="MAND"
                    ControlSource ="MAND"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
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
                    Left =2373
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
                    Left =1345
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
                    Width =10639
                    BorderColor =4210752
                    Name ="Line65"
                End
                Begin Line
                    BorderWidth =2
                    Left =10603
                    Width =0
                    Height =419
                    BorderColor =4210752
                    Name ="Line67"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =8520
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
                    Width =10624
                    Name ="Line71"
                End
                Begin Line
                    BorderWidth =2
                    Width =0
                    Height =397
                    BorderColor =4210752
                    Name ="Line208"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4690
                    Width =1117
                    Height =390
                    FontSize =8
                    Name ="Text215"
                    ControlSource ="=Sum([DARAMAD])"
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    Left =5844
                    Width =0
                    Height =426
                    Name ="Line216"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3481
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
                    Left =4644
                    Width =0
                    Height =426
                    Name ="Line218"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2265
                    Width =1131
                    Height =390
                    FontSize =8
                    TabIndex =2
                    Name ="Text219"
                    ControlSource ="=Sum([SANAVAT])"
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    Left =3419
                    Width =0
                    Height =426
                    Name ="Line220"
                End
                Begin Line
                    Left =2235
                    Width =0
                    Height =426
                    Name ="Line222"
                End
                Begin Line
                    Width =0
                    Height =426
                    Name ="Line224"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1110
                    Width =1086
                    Height =390
                    FontSize =8
                    TabIndex =3
                    Name ="Text229"
                    ControlSource ="=Sum([EYDYMAL])"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    Visible = NotDefault
                    TextAlign =2
                    TextFontFamily =2
                    Left =1513
                    Top =433
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
                Begin Line
                    Left =1080
                    Width =0
                    Height =426
                    Name ="Line239"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =15
                    Width =1041
                    Height =390
                    FontSize =8
                    TabIndex =4
                    Name ="Text240"
                    ControlSource ="=Sum([MAND])"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
            End
        End
    End
End
CodeBehindForm
' See "R_LIST_MALI_EYDY_SANAWAT.cls"
