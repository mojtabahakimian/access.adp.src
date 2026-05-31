Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    FilterOn = NotDefault
    DefaultView =0
    AllowUpdating =4
    TabularCharSet =178
    TabularFamily =42
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =10566
    DatasheetFontHeight =10
    ItemSuffix =116
    Left =600
    Top =1890
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xb3a50d52cd64e340
    End
    RecordSource ="SELECT DEED_HED.N_S, DEED_HED.DATE_S, DEED_HED.SHARH_S, DEED_DTL.HES_K, TOTA_HES"
        ".NAME, SUM(DEED_DTL.BED) AS sbed, SUM(DEED_DTL.BES) AS sbes, DEED_HED.BASE FROM "
        "DEED_DTL INNER JOIN DEED_HED ON DEED_DTL.N_S = DEED_HED.N_S INNER JOIN TOTA_HES "
        "ON DEED_DTL.HES_K = TOTA_HES.NUMBER GROUP BY DEED_HED.N_S, DEED_HED.DATE_S, DEED"
        "_HED.SHARH_S, DEED_DTL.HES_K, TOTA_HES.NAME, DEED_HED.BASE"
    Caption ="سند حسابداري"
    OnOpen ="[Event Procedure]"
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
        Begin Tab
            TextFontCharSet =178
            TextFontFamily =2
            Width =5103
            Height =3402
            FontName ="Arial (Arabic)"
        End
        Begin BreakLevel
            GroupFooter = NotDefault
            ControlSource ="N_S"
        End
        Begin BreakLevel
            SortOrder = NotDefault
            ControlSource ="sbed"
        End
        Begin BreakLevel
            SortOrder = NotDefault
            ControlSource ="sbes"
        End
        Begin PageHeader
            Height =2375
            Name ="PageHeader"
            Begin
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =9465
                    Top =1980
                    Width =1035
                    Height =345
                    FontSize =10
                    ForeColor =0
                    Name ="N_S Label"
                    Caption ="شماره حساب "
                    FontName ="Zar"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="N_S_Label"
                End
                Begin Line
                    BorderWidth =2
                    Left =11
                    Top =1950
                    Width =10510
                    Name ="Line26"
                End
                Begin Line
                    BorderWidth =2
                    Left =15
                    Top =2370
                    Width =10510
                    Name ="Line27"
                End
                Begin Line
                    Left =9540
                    Top =1971
                    Width =0
                    Height =404
                    Name ="Line29"
                End
                Begin Line
                    Left =6538
                    Top =1965
                    Width =0
                    Height =404
                    Name ="Line32"
                End
                Begin Line
                    BorderWidth =2
                    Top =1968
                    Width =0
                    Height =404
                    Name ="Line36"
                End
                Begin Line
                    BorderWidth =2
                    Left =10496
                    Top =1965
                    Width =0
                    Height =404
                    Name ="Line37"
                End
                Begin Label
                    BackStyle =1
                    TextAlign =2
                    TextFontFamily =10
                    Left =-53
                    Width =10545
                    Height =675
                    FontSize =18
                    FontWeight =900
                    Name ="Label22"
                    Caption ="سند حسابداري "
                    FontName ="Titr"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =4095
                    Top =1980
                    Width =1155
                    Height =345
                    FontSize =10
                    ForeColor =0
                    Name ="Label72"
                    Caption ="شــــــــــرح"
                    FontName ="Zar"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =1643
                    Top =1965
                    Width =915
                    Height =345
                    FontSize =10
                    ForeColor =0
                    Name ="Label80"
                    Caption ="مبلغ بدهكار"
                    FontName ="Zar"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    Left =2834
                    Top =1965
                    Width =0
                    Height =404
                    Name ="Line81"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =7672
                    Top =1980
                    Width =810
                    Height =345
                    FontSize =10
                    ForeColor =0
                    Name ="Label90"
                    Caption ="نام حساب "
                    FontName ="Zar"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =9375
                    Top =1080
                    Width =915
                    Height =405
                    FontSize =10
                    ForeColor =0
                    Name ="Label92"
                    Caption ="تاريخ سند:"
                    FontName ="Traffic"
                    Tag ="DetachedLabel"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =8055
                    Top =1080
                    Width =1311
                    Height =420
                    FontSize =12
                    FontWeight =700
                    Name ="Text93"
                    ControlSource ="DATE_S"
                    FontName ="Zar"
                    InputMask ="####/##/##"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =9375
                    Top =675
                    Width =1005
                    Height =405
                    FontSize =10
                    ForeColor =0
                    Name ="Label97"
                    Caption ="شماره سند:"
                    FontName ="Traffic"
                    Tag ="DetachedLabel"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =8097
                    Top =680
                    Width =1251
                    Height =420
                    FontSize =12
                    FontWeight =700
                    TabIndex =1
                    Name ="Text98"
                    ControlSource ="N_S"
                    FontName ="Zar"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =9377
                    Top =1500
                    Width =900
                    Height =405
                    FontSize =10
                    ForeColor =0
                    Name ="Label99"
                    Caption ="شرح سند:"
                    FontName ="Traffic"
                    Tag ="DetachedLabel"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =195
                    Top =1505
                    Width =9171
                    Height =420
                    FontSize =12
                    FontWeight =700
                    TabIndex =2
                    Name ="Text100"
                    ControlSource ="SHARH_S"
                    FontName ="Zar"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =225
                    Top =1965
                    Width =930
                    Height =345
                    FontSize =10
                    ForeColor =0
                    Name ="Label101"
                    Caption ="مبلغ بستانكار"
                    FontName ="Zar"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    Left =1410
                    Top =1965
                    Width =0
                    Height =404
                    Name ="Line102"
                End
                Begin Label
                    Visible = NotDefault
                    TextAlign =2
                    TextFontFamily =10
                    Left =1536
                    Top =690
                    Width =930
                    Height =405
                    FontSize =10
                    ForeColor =0
                    Name ="Label112"
                    Caption ="شماره مبنا:"
                    FontName ="Traffic"
                    Tag ="DetachedLabel"
                End
                Begin TextBox
                    Visible = NotDefault
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =225
                    Top =693
                    Width =1251
                    Height =420
                    FontSize =12
                    FontWeight =700
                    TabIndex =3
                    Name ="Text113"
                    ControlSource ="BASE"
                    FontName ="Zar"
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            CanGrow = NotDefault
            Height =735
            Name ="Detail"
            Begin
                Begin TextBox
                    CanShrink = NotDefault
                    TextAlign =2
                    TextFontFamily =10
                    ReadingOrder =1
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =9521
                    Width =951
                    Height =690
                    FontSize =10
                    Name ="HES_K"
                    ControlSource ="HES_K"
                    FontName ="Zar"
                End
                Begin Line
                    Left =9540
                    Width =0
                    Height =734
                    Name ="Line39"
                End
                Begin Line
                    Left =6538
                    Width =0
                    Height =734
                    Name ="Line42"
                End
                Begin Line
                    BorderWidth =2
                    Width =0
                    Height =734
                    Name ="Line46"
                End
                Begin Line
                    BorderWidth =2
                    Left =10496
                    Width =0
                    Height =734
                    Name ="Line47"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Top =735
                    Width =10523
                    BorderColor =4210752
                    Name ="Line48"
                End
                Begin TextBox
                    CanShrink = NotDefault
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2859
                    Width =3636
                    Height =690
                    FontSize =10
                    TabIndex =1
                    Name ="Text75"
                    ControlSource ="SHARH_S"
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
                    Left =1425
                    Width =1356
                    Height =375
                    FontSize =10
                    TabIndex =2
                    Name ="sbed"
                    ControlSource ="sbed"
                    Format ="Standard"
                    FontName ="Traffic"
                End
                Begin Line
                    Left =2834
                    Width =0
                    Height =734
                    Name ="Line83"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =15
                    Width =1356
                    Height =375
                    FontSize =10
                    TabIndex =3
                    Name ="sbes"
                    ControlSource ="sbes"
                    Format ="Standard"
                    FontName ="Traffic"
                End
                Begin Line
                    Left =1410
                    Width =0
                    Height =734
                    Name ="Line104"
                End
                Begin TextBox
                    CanShrink = NotDefault
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =6570
                    Width =2955
                    Height =690
                    FontSize =10
                    TabIndex =4
                    Name ="Text108"
                    ControlSource ="NAME"
                    FontName ="Zar"
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            ForceNewPage =2
            Height =1020
            Name ="GroupFooter1"
            Begin
                Begin Line
                    BorderWidth =2
                    Width =0
                    Height =419
                    Name ="Line64"
                End
                Begin Line
                    LineSlant = NotDefault
                    Top =419
                    Width =10493
                    BorderColor =4210752
                    Name ="Line65"
                End
                Begin Line
                    BorderWidth =2
                    Left =10482
                    Width =0
                    Height =419
                    Name ="Line67"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =8946
                    Width =1485
                    Height =375
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="Label68"
                    Caption ="جمع عمليات به حروف:"
                    FontName ="Zar"
                    Tag ="DetachedLabel"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1440
                    Width =1386
                    Height =375
                    FontSize =10
                    FontWeight =700
                    Name ="SSAN"
                    ControlSource ="=Sum([sBED])"
                    Format ="Standard"
                    FontName ="Traffic"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =2834
                    Width =0
                    Height =419
                    Name ="Line85"
                End
                Begin Line
                    BorderWidth =2
                    Width =10480
                    Name ="Line96"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2867
                    Width =6021
                    Height =375
                    FontSize =10
                    TabIndex =1
                    Name ="Text105"
                    ControlSource ="=ALPHANUM([SSAN]) & \" \" & \"ريال\""
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
                    Left =15
                    Width =1386
                    Height =375
                    FontSize =10
                    FontWeight =700
                    TabIndex =2
                    Name ="Text114"
                    ControlSource ="=Sum([sBEs])"
                    Format ="Standard"
                    FontName ="Traffic"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =1417
                    Width =0
                    Height =419
                    Name ="Line115"
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
' See "R_SANAD_PRINT2.cls"
