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
    ItemSuffix =80
    Left =600
    Top =1110
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xe9e446061ca9e240
    End
    InputParameters ="@Forms___F_MENU_ANBAR___DT2 BIGINT =[Forms]![F_MENU_ANBAR]![DT2],@Forms___F_MENU"
        "_ANBAR___MANBAR nvarchar(10)=[FormS]![F_MENU_ANBAR]![MANBAR]"
    RecordSource ="AKMOGUDI_KOL_ANBAR"
    Caption ="موجودي كالا به تفكيك انبار"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Arial (Arabic)"
    OnError ="[Event Procedure]"
    UnknownProp = {258 ,66 ,12 ,4 ,280 } Begin
        0x400046006f0072006d0073005f005f005f0046005f004d0045004e0055005f00 ,
        0x41004e004200410052005f005f005f0044005400320020004200490047004900 ,
        0x4e00540020003d005b0046006f0072006d0073005d0021005b0046005f004d00 ,
        0x45004e0055005f0041004e004200410052005d0021005b004400540032005d00 ,
        0x2c00400046006f0072006d0073005f005f005f0046005f004d0045004e005500 ,
        0x5f0041004e004200410052005f005f005f004d0041004e004200410052002000 ,
        0x6e00760061007200630068006100720028003100300029003d005b0046006f00 ,
        0x72006d0053005d0021005b0046005f004d0045004e0055005f0041004e004200 ,
        0x410052005d0021005b004d0041004e004200410052005d00
    End
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
        Begin BreakLevel
            GroupFooter = NotDefault
            ControlSource ="ANBAR"
        End
        Begin BreakLevel
            ControlSource ="NAME"
        End
        Begin PageHeader
            Height =1575
            Name ="PageHeader"
            Begin
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =9463
                    Top =762
                    Width =675
                    Height =765
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
                    TextFontFamily =10
                    Left =8475
                    Top =768
                    Width =945
                    Height =765
                    FontSize =10
                    ForeColor =0
                    Name ="DATE_S Label"
                    Caption ="كد كالا"
                    FontName ="Traffic"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="DATE_S_Label"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =2694
                    Top =765
                    Width =5746
                    Height =765
                    FontSize =10
                    ForeColor =0
                    Name ="SHARH Label"
                    Caption ="نام كالا"
                    FontName ="Traffic"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="SHARH_Label"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =1200
                    Top =768
                    Width =1416
                    Height =765
                    FontSize =10
                    ForeColor =0
                    Name ="BED Label"
                    Caption ="واحد كالا"
                    FontName ="Traffic"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="BED_Label"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =64
                    Top =765
                    Width =1116
                    Height =765
                    FontSize =10
                    ForeColor =0
                    Name ="BES Label"
                    Caption ="موجودي"
                    FontName ="Traffic"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="BES_Label"
                End
                Begin Line
                    BorderWidth =2
                    Left =15
                    Top =750
                    Width =10165
                    Name ="Line26"
                End
                Begin Line
                    BorderWidth =2
                    Left =15
                    Top =1575
                    Width =10165
                    Name ="Line27"
                End
                Begin Line
                    Left =9435
                    Top =768
                    Width =0
                    Height =794
                    Name ="Line28"
                End
                Begin Line
                    Left =8460
                    Top =771
                    Width =0
                    Height =794
                    Name ="Line29"
                End
                Begin Line
                    Left =2664
                    Top =771
                    Width =0
                    Height =794
                    Name ="Line32"
                End
                Begin Line
                    Left =1189
                    Top =780
                    Width =0
                    Height =794
                    Name ="Line33"
                End
                Begin Line
                    BorderWidth =2
                    Top =768
                    Width =0
                    Height =794
                    Name ="Line36"
                End
                Begin Line
                    BorderWidth =2
                    Left =10140
                    Top =768
                    Width =0
                    Height =794
                    Name ="Line37"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =349
                    Width =9516
                    Height =750
                    FontSize =16
                    Name ="ANBARDF"
                    ControlSource ="=\"موجودي انبار:\" & [ANBARN] & \"  تا تاريخ : \" & Format([Forms]![F_MENU_ANBAR"
                        "]![dt2],\"00\\/00\\/00\")"
                    FontName ="Titr"
                End
                Begin TextBox
                    Visible = NotDefault
                    Left =4705
                    Top =1077
                    TabIndex =1
                    Name ="ANBARN"
                    ControlSource ="ANBARN"
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
                    Left =2694
                    Width =5745
                    Height =375
                    ColumnWidth =4650
                    FontSize =10
                    FontWeight =700
                    Name ="NAME"
                    ControlSource ="NAME"
                    FontName ="Traffic"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1200
                    Width =1416
                    Height =375
                    FontSize =10
                    FontWeight =700
                    TabIndex =1
                    Name ="NAMES"
                    ControlSource ="NAMES"
                    Format ="Standard"
                    FontName ="Traffic"
                End
                Begin TextBox
                    DecimalPlaces =3
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =64
                    Width =1117
                    Height =375
                    FontSize =10
                    FontWeight =700
                    TabIndex =2
                    Name ="MAND"
                    ControlSource ="MAND"
                    Format ="Standard"
                    FontName ="Traffic"
                End
                Begin TextBox
                    RunningSum =1
                    TextAlign =2
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =9409
                    Width =741
                    Height =375
                    FontSize =10
                    FontWeight =700
                    TabIndex =3
                    Name ="N_S"
                    ControlSource ="=1"
                    FontName ="Traffic"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =8480
                    Width =897
                    Height =375
                    FontSize =10
                    FontWeight =700
                    TabIndex =4
                    Name ="CODE"
                    ControlSource ="CODE"
                    FontName ="Traffic"
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
                    Left =2664
                    Width =0
                    Height =419
                    Name ="Line42"
                End
                Begin Line
                    Left =1189
                    Width =0
                    Height =419
                    Name ="Line43"
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
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            ForceNewPage =2
            Height =923
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
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =60
                    Width =1117
                    Height =390
                    FontSize =10
                    TabIndex =2
                    Name ="SBED"
                    ControlSource ="=Sum([MAND])"
                    Format ="Standard"
                    FontName ="Traffic"
                End
                Begin Line
                    Left =1200
                    Width =0
                    Height =419
                    Name ="Line60"
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
                    Left =1305
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
                Begin Line
                    BorderWidth =2
                    Width =10165
                    Name ="Line74"
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
' See "R_AK_MOGUDI_ANBAR2.cls"
