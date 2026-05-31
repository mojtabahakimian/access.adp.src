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
    ItemSuffix =99
    Left =600
    Top =1890
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x38d451ea19a3e240
    End
    RecordSource ="OPGET_HED_REP"
    Caption ="دريافت پرداخت روزانه"
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
        Begin BreakLevel
            GroupFooter = NotDefault
            ControlSource ="DT"
        End
        Begin PageHeader
            Height =1755
            Name ="PageHeader"
            Begin
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =9525
                    Top =917
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
                    Left =8409
                    Top =920
                    Width =960
                    Height =405
                    FontSize =10
                    ForeColor =0
                    Name ="DATE_S Label"
                    Caption ="نوع عمليات"
                    FontName ="Traffic"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="DATE_S_Label"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =6270
                    Top =900
                    Width =810
                    Height =405
                    FontSize =10
                    ForeColor =0
                    Name ="SHARH Label"
                    Caption ="از حساب "
                    FontName ="Traffic"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="SHARH_Label"
                End
                Begin Line
                    BorderWidth =2
                    Left =15
                    Top =902
                    Width =10165
                    Name ="Line26"
                End
                Begin Line
                    BorderWidth =2
                    Left =15
                    Top =1727
                    Width =10165
                    Name ="Line27"
                End
                Begin Line
                    Left =8327
                    Top =923
                    Width =0
                    Height =794
                    Name ="Line29"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =4725
                    Top =917
                    Width =0
                    Height =794
                    Name ="Line32"
                End
                Begin Line
                    BorderWidth =2
                    Top =920
                    Width =0
                    Height =794
                    Name ="Line36"
                End
                Begin Line
                    BorderWidth =2
                    Left =10155
                    Top =917
                    Width =0
                    Height =794
                    Name ="Line37"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =2985
                    Width =4395
                    Height =750
                    FontSize =20
                    FontWeight =900
                    ForeColor =0
                    Name ="Label22"
                    Caption ="چاپ صورتحساب عملكرد خزانه"
                    FontName ="Titr"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =1486
                    Top =917
                    Width =3225
                    Height =765
                    FontSize =10
                    ForeColor =0
                    Name ="Label72"
                    Caption ="شرح"
                    FontName ="Traffic"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =45
                    Top =910
                    Width =1356
                    Height =765
                    FontSize =10
                    ForeColor =0
                    Name ="Label80"
                    Caption ="مبلغ"
                    FontName ="Traffic"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    Left =1456
                    Top =917
                    Width =0
                    Height =794
                    Name ="Line81"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =8402
                    Top =1322
                    Width =975
                    Height =405
                    FontSize =10
                    ForeColor =0
                    Name ="Label86"
                    Caption ="نحوه "
                    FontName ="Traffic"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    Left =4740
                    Top =1307
                    Width =4734
                    Name ="Line87"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =6255
                    Top =1350
                    Width =840
                    Height =405
                    FontSize =10
                    ForeColor =0
                    Name ="Label90"
                    Caption ="به حساب "
                    FontName ="Traffic"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =9694
                    Top =397
                    Width =480
                    Height =405
                    FontSize =10
                    ForeColor =0
                    Name ="Label92"
                    Caption ="تاريخ"
                    FontName ="Traffic"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    Left =9465
                    Top =920
                    Width =0
                    Height =794
                    Name ="Line94"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =8320
                    Top =397
                    Width =1311
                    Height =375
                    FontSize =10
                    Name ="Text93"
                    ControlSource ="DT"
                    FontName ="Zar"
                    InputMask ="####/##/##"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =1343
                    Top =450
                    Width =615
                    Height =405
                    FontSize =10
                    ForeColor =0
                    Name ="Label97"
                    Caption ="شماره:"
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
                    Left =56
                    Top =453
                    Width =1281
                    Height =360
                    FontSize =10
                    FontWeight =700
                    TabIndex =1
                    Name ="ID"
                    ControlSource ="ID"
                    FontName ="Zar"
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            CanGrow = NotDefault
            Height =779
            Name ="Detail"
            Begin
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4770
                    Width =3525
                    Height =375
                    ColumnWidth =4650
                    FontSize =10
                    Name ="SHARH"
                    ControlSource ="=[TOTA_HES.Name] & \"-\" & [DETA_HES.Name] & \"-\" & [tDETA_HES.Name]"
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
                    Left =9495
                    Width =621
                    Height =375
                    FontSize =10
                    TabIndex =1
                    Name ="N_S"
                    ControlSource ="=1"
                    FontName ="Zar"
                End
                Begin Line
                    Left =8327
                    Width =0
                    Height =779
                    Name ="Line39"
                End
                Begin Line
                    Left =4725
                    Width =0
                    Height =779
                    Name ="Line42"
                End
                Begin Line
                    BorderWidth =2
                    Width =0
                    Height =779
                    Name ="Line46"
                End
                Begin Line
                    BorderWidth =2
                    Left =10155
                    Width =0
                    Height =779
                    Name ="Line47"
                End
                Begin Line
                    LineSlant = NotDefault
                    Top =779
                    Width =10178
                    BorderColor =4210752
                    Name ="Line48"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =8342
                    Width =1071
                    Height =375
                    FontSize =10
                    TabIndex =2
                    Name ="Text71"
                    ControlSource ="TCOD_DPS.NAMES"
                    FontName ="Zar"
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1485
                    Width =3231
                    Height =750
                    FontSize =10
                    TabIndex =3
                    Name ="Text75"
                    ControlSource ="SHARH"
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
                    Left =47
                    Width =1356
                    Height =375
                    FontSize =10
                    TabIndex =4
                    Name ="MABL"
                    ControlSource ="MABL"
                    Format ="Standard"
                    FontName ="Zar"
                End
                Begin Line
                    Left =1456
                    Width =0
                    Height =779
                    Name ="Line83"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =8325
                    Top =390
                    Width =1101
                    Height =375
                    FontSize =10
                    TabIndex =5
                    Name ="Text88"
                    ControlSource ="TCOD_DPSKIND.NAMES"
                    FontName ="Zar"
                End
                Begin Line
                    Left =4740
                    Top =375
                    Width =4749
                    Name ="Line89"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4770
                    Top =390
                    Width =3525
                    Height =375
                    FontSize =10
                    TabIndex =6
                    Name ="Text91"
                    ControlSource ="=[TOTA_HES_1.Name] & \"-\" & [DETA_HES_1.Name] & \"-\" & [tDETA_HES_1.Name]"
                    FontName ="Zar"
                End
                Begin Line
                    Left =9465
                    Width =0
                    Height =779
                    Name ="Line95"
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
                    ControlSource ="=FARSIDATE(Now())"
                    FontName ="Zar"
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
                    Left =1644
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
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =61
                    Width =1371
                    Height =375
                    FontSize =10
                    TabIndex =2
                    Name ="Text84"
                    ControlSource ="=Sum([MABL])"
                    Format ="Standard"
                    FontName ="Zar"
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
                    Name ="Line96"
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
' See "OR_DB_PRINT.cls"
