Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    AllowUpdating =4
    TabularCharSet =178
    TabularFamily =98
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =10608
    DatasheetFontHeight =10
    ItemSuffix =46
    Left =600
    Top =1710
    DatasheetGridlinesColor =12632256
    MaxRecords =100000
    RecSrcDt = Begin
        0x61748d1e6e30e440
    End
    InputParameters ="@F_MENU_DATE_DT1= Forms![F_MENU_DATE]![DT1],@F_MENU_DATE_DT2= Forms![F_MENU_DATE"
        "]![DT2],@dep = Forms![GOZARESH_FROOSH_MAHSUL]![DEPATMAN]"
    RecordSource ="FROOSH_COUNT_baz"
    Caption ="گزارش فروش به تفكيك واحد"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial"
    UnknownProp = {258 ,66 ,12 ,4 ,272 } Begin
        0x400046005f004d0045004e0055005f0044004100540045005f00440054003100 ,
        0x3d00200046006f0072006d00730021005b0046005f004d0045004e0055005f00 ,
        0x44004100540045005d0021005b004400540031005d002c00400046005f004d00 ,
        0x45004e0055005f0044004100540045005f004400540032003d00200046006f00 ,
        0x72006d00730021005b0046005f004d0045004e0055005f004400410054004500 ,
        0x5d0021005b004400540032005d002c00400064006500700020003d0020004600 ,
        0x6f0072006d00730021005b0047004f005a00410052004500530048005f004600 ,
        0x52004f004f00530048005f004d0041004800530055004c005d0021005b004400 ,
        0x45005000410054004d0041004e005d00
    End
    RecordSourceQualifier ="dbo"
    Begin
        Begin Label
            BackStyle =0
            TextFontFamily =2
            FontName ="Arial"
        End
        Begin Rectangle
            Width =850
            Height =850
        End
        Begin Line
            Width =1701
        End
        Begin TextBox
            FELineBreak = NotDefault
            OldBorderStyle =0
            TextFontFamily =2
            Width =1701
            LabelX =-1701
            FontName ="Arial"
            AsianLineBreak =255
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =0
            Name ="ReportHeader"
        End
        Begin PageHeader
            Height =1177
            Name ="PageHeaderSection"
            Begin
                Begin TextBox
                    TextFontCharSet =178
                    TextAlign =3
                    IMESentenceMode =3
                    Left =5216
                    Top =170
                    Width =3741
                    Height =390
                    FontSize =10
                    Name ="DEPNAME"
                    ControlSource ="DEPNAME"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextFontCharSet =178
                            Left =8957
                            Top =170
                            Width =1590
                            Height =390
                            FontSize =10
                            Name ="Label0"
                            Caption ="گزارش فروش  واحد:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin Label
                    TextFontCharSet =178
                    Left =3628
                    Top =796
                    Width =780
                    Height =330
                    Name ="Label10"
                    Caption ="تعداد فاكتور"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextFontCharSet =178
                    Left =1474
                    Top =796
                    Width =480
                    Height =330
                    Name ="Label1"
                    Caption ="كد كالا"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextFontCharSet =178
                    Left =7540
                    Top =796
                    Width =495
                    Height =330
                    Name ="Label2"
                    Caption ="نام كالا"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextFontCharSet =178
                    Left =4705
                    Top =796
                    Width =615
                    Height =330
                    Name ="Label3"
                    Caption ="مقدار كل"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextFontCharSet =178
                    Left =2494
                    Top =796
                    Width =540
                    Height =330
                    Name ="Label11"
                    Caption ="مبلغ كل"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextFontCharSet =178
                    TextAlign =2
                    Left =9519
                    Top =796
                    Width =1020
                    Height =330
                    Name ="Label15"
                    Caption ="شماره فني"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextFontCharSet =178
                    Left =390
                    Top =796
                    Width =585
                    Height =330
                    Name ="Label12"
                    Caption ="موجودي"
                    FontName ="WeblogmaYekan"
                End
                Begin Rectangle
                    BackStyle =0
                    BorderWidth =2
                    Top =795
                    Width =10608
                    Height =345
                    Name ="Box13"
                End
                Begin Line
                    BorderWidth =1
                    Left =9467
                    Top =793
                    Width =0
                    Height =382
                    Name ="Line14"
                End
                Begin Line
                    BorderWidth =1
                    Left =5612
                    Top =793
                    Width =0
                    Height =382
                    Name ="Line15"
                End
                Begin Line
                    BorderWidth =1
                    Left =4530
                    Top =780
                    Width =0
                    Height =382
                    Name ="Line16"
                End
                Begin Line
                    BorderWidth =1
                    Left =3465
                    Top =780
                    Width =0
                    Height =382
                    Name ="Line17"
                End
                Begin Line
                    BorderWidth =1
                    Left =2100
                    Top =780
                    Width =0
                    Height =382
                    Name ="Line18"
                End
                Begin Line
                    BorderWidth =1
                    Left =1185
                    Top =795
                    Width =0
                    Height =382
                    Name ="Line19"
                End
                Begin TextBox
                    FELineBreak = NotDefault
                    TextFontCharSet =178
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    BackStyle =0
                    NumeralShapes =2
                    Left =56
                    Top =289
                    Width =1021
                    Height =360
                    FontSize =10
                    TabIndex =1
                    Name ="Text52"
                    ControlSource ="=[Forms]![F_MENU_date]![DT2]"
                    FontName ="WeblogmaYekan"
                    InputMask ="##/##/##"
                    AsianLineBreak =0
                End
                Begin TextBox
                    FELineBreak = NotDefault
                    TextFontCharSet =178
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    BackStyle =0
                    NumeralShapes =2
                    Left =1808
                    Top =287
                    Width =1086
                    Height =360
                    FontSize =10
                    TabIndex =2
                    Name ="Text53"
                    ControlSource ="=[Forms]![F_MENU_date]![DT1]"
                    FontName ="WeblogmaYekan"
                    InputMask ="##/##/##"
                    AsianLineBreak =0
                End
                Begin Label
                    TextFontCharSet =178
                    TextAlign =2
                    Left =1125
                    Top =283
                    Width =660
                    Height =360
                    FontSize =10
                    Name ="Label54"
                    Caption ="تا تاريخ:"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextFontCharSet =178
                    TextAlign =2
                    Left =2928
                    Top =285
                    Width =690
                    Height =360
                    FontSize =10
                    Name ="Label55"
                    Caption ="از تاريخ:"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =382
            Name ="Detail"
            Begin
                Begin TextBox
                    TextFontCharSet =178
                    TextAlign =2
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    BackStyle =0
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =3503
                    Width =1011
                    Height =345
                    Name ="TEDAD"
                    ControlSource ="TEDAD"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    TextFontCharSet =178
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    BackStyle =0
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1215
                    Width =831
                    Height =345
                    TabIndex =1
                    Name ="CODE"
                    ControlSource ="CODE"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    TextFontCharSet =178
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    BackStyle =0
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =5645
                    Width =3801
                    Height =345
                    TabIndex =2
                    Name ="NAME"
                    ControlSource ="NAME"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    TextFontCharSet =178
                    TextAlign =2
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    BackStyle =0
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =4565
                    Width =1026
                    Height =345
                    TabIndex =3
                    Name ="SMEGHk"
                    ControlSource ="SMEGHk"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    BackStyle =0
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2126
                    Width =1281
                    Height =345
                    TabIndex =4
                    Name ="SMABL"
                    ControlSource ="SMABL"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    TextFontCharSet =178
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    BackStyle =0
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =9495
                    Width =1056
                    Height =345
                    TabIndex =5
                    Name ="N_FANI"
                    ControlSource ="N_FANI"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    BorderWidth =1
                    SpecialEffect =3
                    Left =9467
                    Width =0
                    Height =382
                    Name ="Line20"
                End
                Begin Line
                    BorderWidth =1
                    SpecialEffect =3
                    Left =5612
                    Width =0
                    Height =382
                    Name ="Line21"
                End
                Begin Line
                    BorderWidth =1
                    SpecialEffect =3
                    Left =4530
                    Width =0
                    Height =382
                    Name ="Line22"
                End
                Begin Line
                    BorderWidth =1
                    SpecialEffect =3
                    Left =3465
                    Width =0
                    Height =382
                    Name ="Line23"
                End
                Begin Line
                    BorderWidth =1
                    SpecialEffect =3
                    Left =2100
                    Width =0
                    Height =382
                    Name ="Line24"
                End
                Begin Line
                    BorderWidth =1
                    SpecialEffect =3
                    Left =1185
                    Width =0
                    Height =382
                    Name ="Line25"
                End
                Begin Line
                    BorderWidth =1
                    Top =375
                    Width =10605
                    Name ="Line26"
                End
                Begin Line
                    BorderWidth =2
                    Width =0
                    Height =382
                    Name ="Line43"
                End
                Begin Line
                    BorderWidth =2
                    Left =10590
                    Width =0
                    Height =382
                    Name ="Line44"
                End
                Begin Line
                    BorderWidth =2
                    Width =0
                    Height =382
                    Name ="Line45"
                End
            End
        End
        Begin PageFooter
            Height =0
            Name ="PageFooterSection"
        End
        Begin FormFooter
            KeepTogether = NotDefault
            Height =680
            Name ="ReportFooter"
            Begin
                Begin TextBox
                    TextFontCharSet =178
                    TextAlign =2
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    BackStyle =0
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =3462
                    Width =1011
                    Height =345
                    Name ="Text28"
                    ControlSource ="=Sum([TEDAD])"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    TextFontCharSet =178
                    TextAlign =2
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    BackStyle =0
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =4569
                    Width =1026
                    Height =345
                    TabIndex =1
                    Name ="Text31"
                    ControlSource ="=Sum([SMEGHk])"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    BackStyle =0
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2085
                    Width =1311
                    Height =345
                    TabIndex =2
                    Name ="Text32"
                    ControlSource ="=Sum([SMABL])"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    BorderWidth =1
                    Left =5610
                    Top =28
                    Width =0
                    Height =382
                    Name ="Line33"
                End
                Begin Line
                    BorderWidth =1
                    Left =4528
                    Top =15
                    Width =0
                    Height =382
                    Name ="Line34"
                End
                Begin Line
                    BorderWidth =1
                    Left =3463
                    Width =0
                    Height =382
                    Name ="Line35"
                End
                Begin Line
                    BorderWidth =1
                    Top =30
                    Width =0
                    Height =382
                    Name ="Line37"
                End
                Begin Label
                    TextFontCharSet =178
                    Left =6120
                    Width =615
                    Height =330
                    Name ="Label38"
                    Caption ="جمع كل:"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    Top =390
                    Width =10605
                    Name ="Line40"
                End
                Begin Line
                    BorderWidth =1
                    Left =10590
                    Width =0
                    Height =382
                    Name ="Line41"
                End
                Begin Line
                    BorderWidth =1
                    Width =0
                    Height =382
                    Name ="Line42"
                End
            End
        End
    End
End
CodeBehindForm
' See "FROOSH_VAHEDS.cls"
