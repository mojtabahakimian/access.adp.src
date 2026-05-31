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
    GridY =10
    Width =10204
    DatasheetFontHeight =10
    ItemSuffix =77
    Left =600
    Top =1500
    HelpContextId =500203
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xc10e2731b3bee340
    End
    RecordSource ="SELECT AI_DAFTAR.N_S, AI_DAFTAR.HES_M, AI_DAFTAR.DATE_S, AI_DAFTAR.HES_K, AI_DAF"
        "TAR.SHARH, AI_DAFTAR.BED, AI_DAFTAR.BES, AI_DAFTAR.BED - AI_DAFTAR.BES AS MAND, "
        "AI_DAFTAR.ID, TDETA_HES.NAME + N' - ' + AI_DAFTAR.HES AS tname, TDETA_HES.ADDRES"
        "S, AI_DAFTAR.HES, AI_DAFTAR.DIFF * AI_DAFTAR.BED AS DD FROM AI_DAFTAR INNER JOIN"
        " TDETA_HES ON AI_DAFTAR.HES_K = TDETA_HES.N_KOL AND AI_DAFTAR.HES_M = TDETA_HES."
        "NUMBER AND AI_DAFTAR.HES_T = TDETA_HES.TNUMBER"
    Caption ="صورتحساب مشتريان جهت تسويه"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    HelpFile ="dena.hlp"
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
            ControlSource ="TNAME"
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            GroupFooter = NotDefault
            ControlSource ="HES_M"
        End
        Begin BreakLevel
            ControlSource ="DATE_S"
        End
        Begin BreakLevel
            SortOrder = NotDefault
            ControlSource ="BED"
        End
        Begin BreakLevel
            SortOrder = NotDefault
            ControlSource ="BES"
        End
        Begin PageHeader
            Height =0
            Name ="PageHeader"
            Begin
                Begin Line
                    BorderWidth =2
                    Width =10165
                    BorderColor =0
                    Name ="Line25"
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =1669
            BreakLevel =1
            Name ="GroupHeader0"
            Begin
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =9463
                    Top =856
                    Width =675
                    Height =765
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="N_S Label"
                    Caption ="شماره سند"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="N_S_Label"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =8475
                    Top =862
                    Width =945
                    Height =765
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="DATE_S Label"
                    Caption ="تاريخ سند"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="DATE_S_Label"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =4935
                    Top =859
                    Width =3496
                    Height =765
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="SHARH Label"
                    Caption ="شرح"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="SHARH_Label"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =3435
                    Top =862
                    Width =1461
                    Height =765
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="BED Label"
                    Caption ="بدهكار"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="BED_Label"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =1950
                    Top =859
                    Width =1461
                    Height =765
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="BES Label"
                    Caption ="بستانكار"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="BES_Label"
                End
                Begin Label
                    TextFontFamily =2
                    Left =1575
                    Top =859
                    Width =352
                    Height =765
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="Expr1 Label"
                    Caption ="تش"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="Expr1_Label"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =30
                    Top =862
                    Width =1491
                    Height =765
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="MAND Label"
                    Caption ="مانده"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="MAND_Label"
                End
                Begin Line
                    BorderWidth =2
                    Left =15
                    Top =844
                    Width =10165
                    Name ="Line26"
                End
                Begin Line
                    BorderWidth =2
                    Left =15
                    Top =1669
                    Width =10165
                    Name ="Line27"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =1
                    NumeralShapes =2
                    Left =3789
                    Top =2
                    Width =5896
                    Height =375
                    FontSize =8
                    Name ="tname"
                    ControlSource ="tname"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    Left =9435
                    Top =862
                    Width =0
                    Height =794
                    Name ="Line28"
                End
                Begin Line
                    Left =8460
                    Top =865
                    Width =0
                    Height =794
                    Name ="Line29"
                End
                Begin Line
                    Left =4905
                    Top =865
                    Width =0
                    Height =794
                    Name ="Line32"
                End
                Begin Line
                    Left =3420
                    Top =874
                    Width =0
                    Height =794
                    Name ="Line33"
                End
                Begin Line
                    Left =1919
                    Top =859
                    Width =0
                    Height =794
                    Name ="Line34"
                End
                Begin Line
                    Left =1545
                    Top =862
                    Width =0
                    Height =794
                    Name ="Line35"
                End
                Begin Line
                    BorderWidth =2
                    Top =862
                    Width =0
                    Height =794
                    Name ="Line36"
                End
                Begin Line
                    BorderWidth =2
                    Left =10140
                    Top =862
                    Width =0
                    Height =794
                    Name ="Line37"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =9654
                    Width =495
                    Height =360
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="DETA_HES.NAME Label"
                    Caption ="حساب:"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="DETA_HES_NAME_Label"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =345
                    Top =402
                    Width =9231
                    Height =375
                    FontSize =8
                    TabIndex =1
                    Name ="ADDRESS"
                    ControlSource ="ADDRESS"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =9658
                    Top =390
                    Width =510
                    Height =375
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="Label75"
                    Caption ="آدرس:"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
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
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4935
                    Width =3495
                    Height =375
                    ColumnWidth =4650
                    FontSize =8
                    Name ="SHARH"
                    ControlSource ="SHARH"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3435
                    Width =1461
                    Height =375
                    FontSize =8
                    TabIndex =1
                    Name ="BED"
                    ControlSource ="BED"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1950
                    Width =1462
                    Height =375
                    FontSize =8
                    TabIndex =2
                    Name ="BES"
                    ControlSource ="BES"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1575
                    Width =351
                    Height =375
                    FontSize =8
                    TabIndex =3
                    Name ="Expr1"
                    ControlSource ="=IIf([MAND]>0,\"بد\",\"بس\")"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =30
                    Width =1491
                    Height =375
                    FontSize =8
                    TabIndex =4
                    Name ="MAN"
                    ControlSource ="=Abs([MAND])"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =9409
                    Width =741
                    Height =375
                    FontSize =8
                    TabIndex =5
                    Name ="N_S"
                    ControlSource ="N_S"
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
                    Left =8480
                    Width =897
                    Height =375
                    FontSize =8
                    TabIndex =6
                    Name ="DATE_SS"
                    ControlSource ="DATE_S"
                    FontName ="WeblogmaYekan"
                    InputMask ="####/##/##"
                End
                Begin Line
                    Left =9435
                    Width =0
                    Height =419
                    Name ="Line38"
                End
                Begin Line
                    Left =8460
                    Width =0
                    Height =419
                    Name ="Line39"
                End
                Begin Line
                    Left =4905
                    Width =0
                    Height =419
                    Name ="Line42"
                End
                Begin Line
                    Left =3420
                    Width =0
                    Height =419
                    Name ="Line43"
                End
                Begin Line
                    Left =1919
                    Width =0
                    Height =419
                    Name ="Line44"
                End
                Begin Line
                    Left =1545
                    Width =0
                    Height =419
                    Name ="Line45"
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
                    Visible = NotDefault
                    RunningSum =1
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5524
                    Width =321
                    Height =375
                    FontSize =8
                    TabIndex =7
                    Name ="MAND"
                    ControlSource ="MAND"
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    TextFontFamily =2
                    Left =6462
                    Top =56
                    Width =1071
                    FontSize =8
                    TabIndex =8
                    Name ="DATE_S"
                    ControlSource ="DATE_S"
                    FontName ="WeblogmaYekan"
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =510
            BreakLevel =1
            Name ="GroupFooter1"
            Begin
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3435
                    Width =1462
                    Height =375
                    FontSize =8
                    Name ="SBED"
                    ControlSource ="=Sum([BED])"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1950
                    Width =1476
                    Height =375
                    FontSize =8
                    TabIndex =1
                    Name ="SBES"
                    ControlSource ="=Sum([BES])"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1560
                    Width =351
                    Height =375
                    FontSize =8
                    TabIndex =2
                    Name ="Text58"
                    ControlSource ="=IIf([SBED]-[SBES]>0,\"بد\",\"بس\")"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =45
                    Width =1491
                    Height =375
                    FontSize =8
                    TabIndex =3
                    Name ="Text59"
                    ControlSource ="=Abs([SBED]-[SBES])"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    Left =4906
                    Width =0
                    Height =419
                    Name ="Line60"
                End
                Begin Line
                    Left =3420
                    Width =0
                    Height =419
                    Name ="Line61"
                End
                Begin Line
                    Left =1919
                    Width =0
                    Height =419
                    Name ="Line62"
                End
                Begin Line
                    Left =1545
                    Width =0
                    Height =419
                    Name ="Line63"
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
                    TextFontFamily =2
                    Left =5102
                    Width =645
                    Height =375
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
                    Width =10165
                    Name ="Line71"
                End
                Begin Line
                    BorderWidth =3
                    Top =510
                    Width =10204
                    BorderColor =6723891
                    Name ="Line73"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =6360
                    Top =-4
                    Width =2151
                    Height =450
                    FontSize =8
                    TabIndex =4
                    Name ="Text78"
                    ControlSource ="=Sum([Dd])/IIf(Sum([BED])=0,1,Sum([BED]))"
                    Format ="#\"  روز\""
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextAlign =2
                            TextFontFamily =2
                            Left =8543
                            Top =15
                            Width =1485
                            Height =450
                            FontSize =8
                            FontWeight =400
                            ForeColor =0
                            Name ="Label76"
                            Caption ="راس مدت تا امروز:"
                            FontName ="WeblogmaYekan"
                            Tag ="DetachedLabel"
                        End
                    End
                End
            End
        End
        Begin PageFooter
            Height =300
            Name ="PageFooter"
            Begin
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =170
                    Width =4421
                    Height =300
                    FontSize =8
                    Name ="Text24"
                    ControlSource ="=\"صفحه \" & [Page] & \" از \" & [Pages]"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5385
                    Width =4421
                    Height =300
                    FontSize =8
                    TabIndex =1
                    Name ="Text23"
                    ControlSource ="=farsidate(Now())"
                    FontName ="WeblogmaYekan"
                    InputMask ="####/##/##"
                End
            End
        End
    End
End
CodeBehindForm
' See "R_DAFTAR_AI.cls"
