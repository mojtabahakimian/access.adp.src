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
    Width =10021
    DatasheetFontHeight =10
    ItemSuffix =194
    Left =270
    Top =780
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xae970818f3f7e340
    End
    RecordSource ="SELECT STUF_DEF.CODE, STUF_DEF.NAME, TCOD_ANBAR.NAMES, ANBGRD_LST.GRD_NUM, TCOD_"
        "VAHEDS.NAMES AS vahed, ANBGRD_LST.MOG, ANBGRD_LST.NUM1, ANBGRD_LST.MOG - ANBGRD_"
        "LST.NUM1 AS ekht, ANBGRD_HEAD.GRD_ANBAR, STUF_FSK.POSITION FROM ANBGRD_HEAD INNE"
        "R JOIN TCOD_ANBAR ON ANBGRD_HEAD.GRD_ANBAR = TCOD_ANBAR.CODE INNER JOIN STUF_DEF"
        " INNER JOIN ANBGRD_LST ON STUF_DEF.CODE = ANBGRD_LST.CODE INNER JOIN TCOD_VAHEDS"
        " ON STUF_DEF.VAHED = TCOD_VAHEDS.CODE ON ANBGRD_HEAD.GRD_NUM = ANBGRD_LST.GRD_NU"
        "M INNER JOIN STUF_FSK ON ANBGRD_LST.CODE = STUF_FSK.CODE AND ANBGRD_HEAD.GRD_ANB"
        "AR = STUF_FSK.ANBAR"
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
            GroupHeader = NotDefault
            ControlSource ="POSITION"
        End
        Begin BreakLevel
            ControlSource ="NAME"
        End
        Begin PageHeader
            Height =545
            Name ="PageHeader"
            Begin
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2326
                    Top =5
                    Width =4761
                    Height =540
                    FontSize =14
                    Name ="Text109"
                    ControlSource ="NAMES"
                    FontName ="Titr"
                    Begin
                        Begin Label
                            TextAlign =0
                            TextFontFamily =10
                            Left =7126
                            Width =2895
                            Height =540
                            FontSize =14
                            ForeColor =0
                            Name ="Label110"
                            Caption ="موجودي كالاها در انبار:"
                            FontName ="Titr"
                        End
                    End
                End
                Begin Label
                    TextAlign =0
                    TextFontFamily =10
                    Left =170
                    Width =1320
                    Height =540
                    FontSize =14
                    ForeColor =0
                    Name ="Label170"
                    Caption ="شمارش اول"
                    FontName ="Titr"
                End
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Left =1474
                    Top =56
                    Width =786
                    Height =450
                    TabIndex =1
                    Name ="GRD_NUM"
                    ControlSource ="GRD_NUM"
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            RepeatSection = NotDefault
            Height =422
            Name ="GroupHeader0"
            Begin
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =7230
                    Top =15
                    Width =510
                    Height =375
                    FontSize =10
                    ForeColor =0
                    Name ="N_S Label"
                    Caption ="رديف"
                    FontName ="Traffic"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="N_S_Label"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =735
                    Width =585
                    Height =375
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="MAND Label"
                    Caption ="شمارش"
                    FontName ="Traffic"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="MAND_Label"
                End
                Begin Line
                    BorderWidth =2
                    Width =7768
                    Name ="Line26"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =2
                    Top =392
                    Width =7768
                    Name ="Line27"
                End
                Begin Line
                    Left =2070
                    Top =15
                    Width =0
                    Height =404
                    Name ="Line34"
                End
                Begin Line
                    BorderWidth =2
                    Top =18
                    Width =0
                    Height =404
                    Name ="Line36"
                End
                Begin Line
                    BorderWidth =2
                    Left =7755
                    Top =15
                    Width =0
                    Height =404
                    Name ="Line37"
                End
                Begin Line
                    Left =7155
                    Top =15
                    Width =0
                    Height =404
                    Name ="Line162"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =2835
                    Top =30
                    Width =3495
                    Height =330
                    FontSize =10
                    ForeColor =0
                    Name ="Label164"
                    Caption ="نام كالا"
                    FontName ="Traffic"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    Left =6438
                    Width =0
                    Height =404
                    Name ="Line166"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =6495
                    Top =15
                    Width =630
                    Height =405
                    FontSize =10
                    ForeColor =0
                    Name ="Label168"
                    Caption ="كد"
                    FontName ="Traffic"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =2040
                    Top =15
                    Width =675
                    Height =375
                    FontSize =10
                    ForeColor =0
                    Name ="Label171"
                    Caption ="واحد"
                    FontName ="Traffic"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    Left =2762
                    Top =15
                    Width =0
                    Height =404
                    Name ="Line172"
                End
                Begin Line
                    Left =1350
                    Top =15
                    Width =0
                    Height =404
                    Name ="Line185"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =1425
                    Top =15
                    Width =585
                    Height =315
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="Label186"
                    Caption ="موجودي"
                    FontName ="Traffic"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    Left =675
                    Width =0
                    Height =419
                    Name ="Line191"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =75
                    Width =585
                    Height =375
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="Label192"
                    Caption ="اختلاف"
                    FontName ="Traffic"
                    Tag ="DetachedLabel"
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =419
            OnFormat ="[Event Procedure]"
            Name ="Detail"
            Begin
                Begin TextBox
                    RunningSum =1
                    TextAlign =2
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =7185
                    Top =15
                    Width =561
                    Height =360
                    FontSize =10
                    Name ="N_S"
                    ControlSource ="=1"
                    FontName ="Traffic"
                End
                Begin Line
                    Left =2070
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
                    Left =7755
                    Width =0
                    Height =419
                    Name ="Line47"
                End
                Begin Line
                    LineSlant = NotDefault
                    Top =405
                    Width =7768
                    BorderColor =4210752
                    Name ="Line48"
                End
                Begin Line
                    Left =7155
                    Width =0
                    Height =419
                    Name ="Line163"
                End
                Begin TextBox
                    DecimalPlaces =3
                    RunningSum =1
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    BackStyle =1
                    NumeralShapes =2
                    Left =2820
                    Top =15
                    Width =3606
                    Height =360
                    FontSize =10
                    TabIndex =1
                    Name ="NAME1"
                    ControlSource ="NAME"
                    Format ="Standard"
                    FontName ="Traffic"
                End
                Begin Line
                    Left =6435
                    Width =0
                    Height =419
                    Name ="Line167"
                End
                Begin TextBox
                    RunningSum =1
                    TextAlign =2
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =6480
                    Top =15
                    Width =651
                    Height =360
                    FontSize =10
                    TabIndex =2
                    Name ="CODE"
                    ControlSource ="CODE"
                    FontName ="Traffic"
                End
                Begin Line
                    Left =2762
                    Width =0
                    Height =419
                    Name ="Line173"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2100
                    Top =30
                    Width =636
                    Height =315
                    FontSize =8
                    TabIndex =3
                    Name ="vahed"
                    ControlSource ="vahed"
                    FontName ="Traffic"
                End
                Begin Line
                    Left =1351
                    Width =0
                    Height =419
                    Name ="Line187"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1395
                    Top =30
                    Width =651
                    Height =315
                    FontSize =8
                    TabIndex =4
                    Name ="Text188"
                    ControlSource ="MOG"
                    FontName ="Traffic"
                End
                Begin Line
                    Left =675
                    Width =0
                    Height =419
                    Name ="Line189"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =705
                    Top =30
                    Width =606
                    Height =315
                    FontSize =8
                    TabIndex =5
                    Name ="Text190"
                    ControlSource ="NUM1"
                    FontName ="Traffic"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =45
                    Top =30
                    Width =606
                    Height =315
                    FontSize =8
                    TabIndex =6
                    Name ="Text193"
                    ControlSource ="ekht"
                    FontName ="Traffic"
                End
            End
        End
        Begin PageFooter
            Height =780
            Name ="PageFooter"
            Begin
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Width =1076
                    Height =300
                    FontSize =10
                    Name ="Text24"
                    ControlSource ="=\"صفحه \" & [Page] & \" از \" & [Pages]"
                    FontName ="Zar"
                End
                Begin Label
                    TextAlign =0
                    TextFontFamily =2
                    Left =9240
                    Width =660
                    Height =330
                    FontSize =8
                    ForeColor =0
                    Name ="Label180"
                    Caption ="شمارشگر1"
                    FontName ="Titr"
                End
                Begin Label
                    TextAlign =0
                    TextFontFamily =2
                    Left =7715
                    Width =690
                    Height =330
                    FontSize =8
                    ForeColor =0
                    Name ="Label181"
                    Caption ="شمارشگر2"
                    FontName ="Titr"
                End
                Begin Label
                    TextAlign =0
                    TextFontFamily =2
                    Left =6179
                    Width =735
                    Height =330
                    FontSize =8
                    ForeColor =0
                    Name ="Label182"
                    Caption ="اپراتور ثبت"
                    FontName ="Titr"
                End
                Begin Label
                    TextAlign =0
                    TextFontFamily =2
                    Left =4650
                    Width =525
                    Height =330
                    FontSize =8
                    ForeColor =0
                    Name ="Label183"
                    Caption ="مديريت"
                    FontName ="Titr"
                End
                Begin Label
                    TextAlign =0
                    TextFontFamily =2
                    Left =2715
                    Width =1395
                    Height =330
                    FontSize =8
                    ForeColor =0
                    Name ="Label184"
                    Caption ="تاريخ و ساعت شمارش:"
                    FontName ="Titr"
                End
            End
        End
    End
End
CodeBehindForm
' See "r_anbargrd.cls"
