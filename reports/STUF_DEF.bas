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
    Width =9467
    DatasheetFontHeight =10
    ItemSuffix =27
    Left =930
    Top =1440
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xbdd6b8d979f1e340
    End
    RecordSource ="SELECT STUF_DEF.NAME + ' - ' + STUF_DEF.CODE AS NAM, STUF_DEF.MABL_F, 1 AS Expr1"
        ", STUF_DEF.B_SEF, MOGUDI_KOL_ANBARHA.MANDAH FROM STUF_DEF INNER JOIN MOGUDI_KOL_"
        "ANBARHA ON STUF_DEF.CODE = MOGUDI_KOL_ANBARHA.CODE"
    Caption ="ليست كالاها جهت ويزيت"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Arial (Arabic)"
    Begin
        Begin Label
            FontItalic = NotDefault
            BackStyle =0
            TextFontCharSet =178
            TextAlign =1
            TextFontFamily =42
            FontSize =11
            FontWeight =700
            ForeColor =8388608
            FontName ="Arial (Arabic)"
        End
        Begin Rectangle
            BackStyle =0
            BorderWidth =1
            Width =850
            Height =850
            BorderColor =8388608
        End
        Begin Line
            Width =1701
            BorderColor =8388608
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
            FontName ="Arial (Arabic)"
        End
        Begin ListBox
            TextFontCharSet =178
            TextFontFamily =42
            OldBorderStyle =0
            Width =1701
            Height =1417
            LabelX =-1701
            FontName ="Arial (Arabic)"
        End
        Begin ComboBox
            OldBorderStyle =0
            TextFontCharSet =178
            TextFontFamily =42
            BackStyle =0
            Width =1701
            LabelX =-1701
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
            ControlSource ="Expr1"
        End
        Begin BreakLevel
            ControlSource ="NAM"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =0
            Name ="ReportHeader"
        End
        Begin PageHeader
            Height =1035
            Name ="PageHeader"
            Begin
                Begin Label
                    FontItalic = NotDefault
                    TextFontFamily =2
                    Left =4785
                    Width =1380
                    Height =540
                    FontSize =14
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label88"
                    Caption ="ليست اجناس"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1020
                    Top =495
                    Width =8283
                    Height =540
                    FontSize =14
                    BorderColor =32768
                    Name ="Text90"
                    ControlSource ="=Forms!Baseknow!WIDTH_D"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =2
                    Left =56
                    Top =56
                    Width =3005
                    Height =480
                    FontSize =11
                    TabIndex =1
                    Name ="Text24"
                    ControlSource ="=farsidate(Now())"
                    FontName ="WeblogmaYekan"
                    InputMask ="####/##/##"
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            RepeatSection = NotDefault
            NewRowOrCol =1
            Height =510
            Name ="GroupHeader2"
            Begin
                Begin Label
                    FontItalic = NotDefault
                    TextAlign =2
                    TextFontFamily =2
                    Left =2312
                    Top =45
                    Width =1365
                    Height =420
                    FontSize =12
                    FontWeight =400
                    Name ="NAME Label"
                    Caption ="نام كالا"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="NAME_Label"
                End
                Begin Label
                    FontItalic = NotDefault
                    TextAlign =2
                    TextFontFamily =2
                    Left =60
                    Top =46
                    Width =1185
                    Height =420
                    FontSize =12
                    FontWeight =400
                    Name ="MABL_F Label"
                    Caption ="قيمت"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="MABL_F_Label"
                End
                Begin Label
                    FontItalic = NotDefault
                    TextAlign =2
                    TextFontFamily =2
                    Left =4537
                    Top =45
                    Width =510
                    Height =420
                    FontSize =12
                    FontWeight =400
                    Name ="Label16"
                    Caption ="رديف"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    LineSlant = NotDefault
                    Top =510
                    Width =5135
                    BorderColor =0
                    Name ="Line20"
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =526
            Name ="Detail"
            Begin
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1264
                    Width =3225
                    Height =495
                    FontSize =10
                    BorderColor =4210752
                    Name ="NAM"
                    ControlSource ="NAM"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =2
                    TextFontFamily =2
                    NumeralShapes =2
                    Left =45
                    Width =1155
                    Height =495
                    FontSize =11
                    TabIndex =1
                    BorderColor =4210752
                    Name ="MABL_F"
                    ControlSource ="MABL_F"
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    Left =4532
                    Width =0
                    Height =525
                    BorderColor =0
                    Name ="Line12"
                End
                Begin Line
                    Left =1237
                    Width =0
                    Height =525
                    BorderColor =0
                    Name ="Line13"
                End
                Begin Line
                    Width =0
                    Height =525
                    BorderColor =0
                    Name ="Line15"
                End
                Begin TextBox
                    RunningSum =2
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4545
                    Width =570
                    Height =495
                    FontSize =12
                    TabIndex =2
                    BorderColor =4210752
                    Name ="Text17"
                    ControlSource ="=1"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    Left =5130
                    Width =0
                    Height =525
                    BorderColor =0
                    Name ="Line18"
                End
                Begin Line
                    LineSlant = NotDefault
                    Top =525
                    Width =5150
                    BorderColor =0
                    Name ="Line14"
                End
                Begin TextBox
                    Visible = NotDefault
                    TextFontFamily =2
                    Left =1757
                    Top =226
                    TabIndex =3
                    Name ="MANDAH"
                    ControlSource ="MANDAH"
                    FontName ="WeblogmaYekan"
                End
            End
        End
        Begin PageFooter
            Height =488
            Name ="PageFooter"
            Begin
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =2
                    Left =3855
                    Top =-7
                    Width =3005
                    Height =480
                    FontSize =10
                    Name ="Text21"
                    ControlSource ="=\"صفحه \" & [Page] & \" از\" & [Pages]"
                    FontName ="WeblogmaYekan"
                End
            End
        End
        Begin FormFooter
            KeepTogether = NotDefault
            CanGrow = NotDefault
            Height =420
            Name ="ReportFooter"
            Begin
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TextAlign =2
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =53
                    Width =9311
                    Height =420
                    FontSize =12
                    Name ="Text25"
                    ControlSource ="=Forms!F_MENU_AJNAS!MATN"
                    FontName ="WeblogmaYekan"
                End
            End
        End
    End
End
CodeBehindForm
' See "STUF_DEF.cls"
