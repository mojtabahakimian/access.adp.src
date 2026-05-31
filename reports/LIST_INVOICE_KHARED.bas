Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    FilterOn = NotDefault
    OrderByOn = NotDefault
    DefaultView =0
    AllowUpdating =4
    TabularCharSet =178
    TabularFamily =42
    DateGrouping =1
    Orientation =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =10611
    DatasheetFontHeight =10
    ItemSuffix =161
    Left =270
    Top =390
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x744ed95e9a46e340
    End
    InputParameters ="@Forms___F_MENU_KHFR___DT1 =[Forms]![F_MENU_KHFR]![DT1] , @Forms___F_MENU_KHFR__"
        "_DT2  =[Forms]![F_MENU_KHFR]![DT2] ,  @Forms___F_MENU_KHFR___MMOIN  nvarchar(20)"
        "=[Forms]![F_MENU_KHFR]![MMOIN]"
    RecordSource ="QSL_LIST_INVOICE_KHAREDG"
    Caption ="فاكتور خريد"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Arial (Arabic)"
    OnError ="[Event Procedure]"
    UnknownProp = {258 ,66 ,12 ,4 ,380 } Begin
        0x400046006f0072006d0073005f005f005f0046005f004d0045004e0055005f00 ,
        0x4b004800460052005f005f005f0044005400310020003d005b0046006f007200 ,
        0x6d0073005d0021005b0046005f004d0045004e0055005f004b00480046005200 ,
        0x5d0021005b004400540031005d0020002c002000400046006f0072006d007300 ,
        0x5f005f005f0046005f004d0045004e0055005f004b004800460052005f005f00 ,
        0x5f00440054003200200020003d005b0046006f0072006d0073005d0021005b00 ,
        0x46005f004d0045004e0055005f004b004800460052005d0021005b0044005400 ,
        0x32005d0020002c0020002000400046006f0072006d0073005f005f005f004600 ,
        0x5f004d0045004e0055005f004b004800460052005f005f005f004d004d004f00 ,
        0x49004e00200020006e0076006100720063006800610072002800320030002900 ,
        0x3d005b0046006f0072006d0073005d0021005b0046005f004d0045004e005500 ,
        0x5f004b004800460052005d0021005b004d004d004f0049004e005d00
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
        Begin PageBreak
            Width =283
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            GroupFooter = NotDefault
            ControlSource ="NUMBER"
        End
        Begin BreakLevel
            ControlSource ="INUMBER"
        End
        Begin BreakLevel
            ControlSource ="DATE_N"
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =2970
            Name ="GroupHeader2"
            Begin
                Begin Label
                    OverlapFlags =81
                    TextAlign =2
                    TextFontFamily =2
                    Left =4275
                    Width =1965
                    Height =540
                    FontSize =8
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label88"
                    Caption ="فاكتور خريـــد"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    OverlapFlags =211
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1875
                    Top =495
                    Width =5613
                    Height =540
                    FontSize =8
                    BorderColor =32768
                    Name ="Text90"
                    ControlSource ="=[Forms]![Baseknow]![WIDTH_D]"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    OverlapFlags =81
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =60
                    Top =540
                    Width =1116
                    Height =450
                    FontSize =8
                    TabIndex =1
                    BorderColor =32768
                    Name ="DATE_N"
                    ControlSource ="DATE_N"
                    FontName ="WeblogmaYekan"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =83
                            TextAlign =3
                            TextFontFamily =2
                            Left =1213
                            Top =540
                            Width =660
                            Height =480
                            FontSize =8
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label92"
                            Caption ="تاريخ :"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =83
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =7485
                    Top =510
                    Width =1686
                    Height =435
                    FontSize =8
                    TabIndex =2
                    BorderColor =32768
                    Name ="HEAD_LST.NUMBER"
                    ControlSource ="NUMBER"
                    FontName ="WeblogmaYekan"
                    EventProcPrefix ="HEAD_LST_NUMBER"
                    Begin
                        Begin Label
                            OverlapFlags =81
                            TextAlign =3
                            TextFontFamily =2
                            Left =9222
                            Top =510
                            Width =1275
                            Height =480
                            FontSize =8
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label93"
                            Caption ="شماره فاكتور:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin Line
                    BorderWidth =3
                    OverlapFlags =115
                    Top =1050
                    Width =10611
                    BorderColor =0
                    Name ="Line94"
                End
                Begin TextBox
                    OverlapFlags =115
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1875
                    Top =1062
                    Width =7611
                    Height =450
                    FontSize =8
                    TabIndex =3
                    BorderColor =32768
                    Name ="BESTANKARA"
                    ControlSource ="CNAME"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            OverlapFlags =115
                            TextAlign =3
                            TextFontFamily =2
                            Left =9525
                            Top =1065
                            Width =1005
                            Height =480
                            FontSize =8
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label95"
                            Caption ="فروشنده:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =243
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =120
                    Top =1050
                    Width =981
                    Height =450
                    FontSize =8
                    TabIndex =4
                    BorderColor =32768
                    Name ="BESN"
                    ControlSource ="BESN"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            OverlapFlags =243
                            TextAlign =3
                            TextFontFamily =2
                            Left =1092
                            Top =1050
                            Width =405
                            Height =480
                            FontSize =8
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label96"
                            Caption ="كد:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =243
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1875
                    Top =1524
                    Width =7821
                    Height =450
                    FontSize =8
                    TabIndex =5
                    BorderColor =32768
                    Name ="ADDRESS"
                    ControlSource ="ADDRESS"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            OverlapFlags =243
                            TextAlign =3
                            TextFontFamily =2
                            Left =9732
                            Top =1524
                            Width =765
                            Height =480
                            FontSize =8
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label97"
                            Caption ="آدرس:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =81
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =60
                    Top =1515
                    Width =1011
                    Height =450
                    FontSize =8
                    TabIndex =6
                    BorderColor =32768
                    Name ="TEL"
                    ControlSource ="TEL"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            OverlapFlags =243
                            TextAlign =3
                            TextFontFamily =2
                            Left =1092
                            Top =1515
                            Width =555
                            Height =480
                            FontSize =8
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label98"
                            Caption ="تلفن:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin Line
                    BorderWidth =3
                    OverlapFlags =81
                    Top =2445
                    Width =10611
                    BorderColor =0
                    Name ="Line99"
                End
                Begin Line
                    BorderWidth =3
                    OverlapFlags =83
                    Left =10575
                    Top =2460
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line130"
                End
                Begin Line
                    OverlapFlags =83
                    Left =10005
                    Top =2460
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line131"
                End
                Begin Line
                    OverlapFlags =83
                    Left =9435
                    Top =2460
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line132"
                End
                Begin Line
                    OverlapFlags =83
                    Left =5625
                    Top =2460
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line133"
                End
                Begin Line
                    OverlapFlags =83
                    Left =4485
                    Top =2460
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line134"
                End
                Begin Line
                    OverlapFlags =83
                    Left =3660
                    Top =2460
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line135"
                End
                Begin Line
                    OverlapFlags =83
                    Left =2715
                    Top =2460
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line136"
                End
                Begin Line
                    OverlapFlags =83
                    Left =1215
                    Top =2460
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line137"
                End
                Begin Line
                    BorderWidth =3
                    OverlapFlags =83
                    Left =30
                    Top =2460
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line138"
                End
                Begin Label
                    OverlapFlags =81
                    TextAlign =2
                    TextFontFamily =2
                    Left =330
                    Top =2505
                    Width =585
                    Height =405
                    FontSize =8
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label139"
                    Caption ="مبــلغ"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    OverlapFlags =81
                    TextAlign =2
                    TextFontFamily =2
                    Left =1689
                    Top =2501
                    Width =585
                    Height =405
                    FontSize =8
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label140"
                    Caption ="فـــي"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    OverlapFlags =81
                    TextAlign =2
                    TextFontFamily =2
                    Left =2805
                    Top =2505
                    Width =810
                    Height =405
                    FontSize =8
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label141"
                    Caption ="مقدار كل"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    OverlapFlags =81
                    TextAlign =2
                    TextFontFamily =2
                    Left =3810
                    Top =2505
                    Width =525
                    Height =405
                    FontSize =8
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label142"
                    Caption ="مقدار"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    OverlapFlags =81
                    TextAlign =2
                    TextFontFamily =2
                    Left =4687
                    Top =2505
                    Width =750
                    Height =405
                    FontSize =8
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label143"
                    Caption ="واحد كالا"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    OverlapFlags =81
                    TextAlign =2
                    TextFontFamily =2
                    Left =6727
                    Top =2505
                    Width =1710
                    Height =405
                    FontSize =8
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label144"
                    Caption ="شــــــــــرح  كالا"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    OverlapFlags =81
                    TextAlign =2
                    TextFontFamily =2
                    Left =9510
                    Top =2505
                    Width =375
                    Height =405
                    FontSize =8
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label145"
                    Caption ="انبار"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    OverlapFlags =83
                    TextAlign =2
                    TextFontFamily =2
                    Left =10011
                    Top =2501
                    Width =540
                    Height =405
                    FontSize =8
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label146"
                    Caption ="رديف"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    BorderWidth =3
                    OverlapFlags =83
                    Top =2970
                    Width =10611
                    BorderColor =0
                    Name ="Line147"
                End
                Begin TextBox
                    OverlapFlags =243
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =60
                    Top =1995
                    Width =9426
                    Height =450
                    FontSize =8
                    TabIndex =7
                    BorderColor =32768
                    Name ="MOLAH"
                    ControlSource ="MOLAH"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            OverlapFlags =243
                            TextAlign =3
                            TextFontFamily =2
                            Left =9510
                            Top =1995
                            Width =1035
                            Height =480
                            FontSize =8
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label155"
                            Caption ="ملاحظات:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =450
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =81
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =9467
                    Top =16
                    Width =504
                    Height =405
                    FontSize =8
                    FontWeight =700
                    BorderColor =32768
                    Name ="INVO_LST.ANBAR"
                    ControlSource ="IANBAR"
                    FontName ="WeblogmaYekan"
                    EventProcPrefix ="INVO_LST_ANBAR"
                End
                Begin TextBox
                    DecimalPlaces =0
                    RunningSum =1
                    OverlapFlags =81
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =10021
                    Top =16
                    Width =504
                    Height =405
                    FontSize =8
                    FontWeight =700
                    TabIndex =1
                    BorderColor =32768
                    Name ="RADIF"
                    ControlSource ="=1"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    OverlapFlags =81
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5637
                    Top =16
                    Width =3774
                    Height =405
                    FontSize =8
                    FontWeight =700
                    TabIndex =2
                    BorderColor =32768
                    Name ="CODE"
                    ControlSource ="NAME"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =2
                    OverlapFlags =81
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3663
                    Top =16
                    Width =789
                    Height =405
                    FontSize =8
                    FontWeight =700
                    TabIndex =3
                    BorderColor =32768
                    Name ="MEGH"
                    ControlSource ="MEGH"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =2
                    OverlapFlags =81
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2730
                    Top =16
                    Width =894
                    Height =405
                    FontSize =8
                    FontWeight =700
                    TabIndex =4
                    BorderColor =32768
                    Name ="MEGHk"
                    ControlSource ="MEGHk"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =81
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1234
                    Top =16
                    Width =1449
                    Height =405
                    FontSize =8
                    FontWeight =700
                    TabIndex =5
                    BorderColor =32768
                    Name ="MABL"
                    ControlSource ="MABL"
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =81
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =60
                    Top =16
                    Width =1119
                    Height =405
                    FontSize =8
                    FontWeight =700
                    TabIndex =6
                    BorderColor =32768
                    Name ="MABL_K"
                    ControlSource ="MABL_K"
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    OverlapFlags =81
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4485
                    Top =15
                    Width =1110
                    Height =405
                    FontSize =8
                    FontWeight =700
                    TabIndex =7
                    BorderColor =32768
                    Name ="VAHED_K"
                    ControlSource ="TNAMES"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    OverlapFlags =81
                    Top =450
                    Width =10596
                    BorderColor =0
                    Name ="Line100"
                End
                Begin Line
                    BorderWidth =3
                    OverlapFlags =83
                    Left =10575
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line101"
                End
                Begin Line
                    BorderWidth =1
                    OverlapFlags =83
                    Left =10005
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line102"
                End
                Begin Line
                    BorderWidth =1
                    OverlapFlags =83
                    Left =9435
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line103"
                End
                Begin Line
                    BorderWidth =1
                    OverlapFlags =83
                    Left =5625
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line104"
                End
                Begin Line
                    BorderWidth =1
                    OverlapFlags =83
                    Left =4485
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line105"
                End
                Begin Line
                    BorderWidth =1
                    OverlapFlags =83
                    Left =3660
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line106"
                End
                Begin Line
                    BorderWidth =1
                    OverlapFlags =83
                    Left =2715
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line107"
                End
                Begin Line
                    BorderWidth =1
                    OverlapFlags =83
                    Left =1215
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line108"
                End
                Begin Line
                    BorderWidth =2
                    OverlapFlags =83
                    Left =15
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line109"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    RunningSum =1
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =277
                    Width =624
                    Height =405
                    FontSize =8
                    FontWeight =700
                    TabIndex =8
                    BorderColor =32768
                    Name ="SMABL_K"
                    ControlSource ="MABL_K"
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =3288
            OnFormat ="[Event Procedure]"
            Name ="GroupFooter3"
            Begin
                Begin Line
                    BorderWidth =3
                    OverlapFlags =81
                    Top =30
                    Width =10555
                    BorderColor =0
                    Name ="Line110"
                End
                Begin TextBox
                    OverlapFlags =211
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3518
                    Top =95
                    Width =5796
                    Height =750
                    FontSize =8
                    FontWeight =700
                    BorderColor =32768
                    Name ="Text111"
                    ControlSource ="=ALPHANUM([GHABEL]) & \" \" & \"ريال\""
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            OverlapFlags =81
                            TextAlign =3
                            TextFontFamily =2
                            Left =9297
                            Top =83
                            Width =1200
                            Height =405
                            FontSize =8
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label112"
                            Caption ="مبلغ به حروف:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =81
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    BackStyle =1
                    NumeralShapes =2
                    Left =60
                    Top =1710
                    Width =1761
                    Height =435
                    FontSize =8
                    FontWeight =700
                    TabIndex =1
                    Name ="NPAR"
                    ControlSource ="=[M_NAGHD]+[MABL_VAR]+[MABL_HAV]+[mab]"
                    Format ="#,##0.00;-#,##0.00"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            OverlapFlags =81
                            TextAlign =0
                            TextFontFamily =2
                            Left =1876
                            Top =1717
                            Width =1440
                            Height =435
                            FontSize =8
                            ForeColor =0
                            Name ="Label86"
                            Caption ="جمع مبالغ پرداختي:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =81
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    BackStyle =1
                    NumeralShapes =2
                    Left =61
                    Top =90
                    Width =1761
                    Height =435
                    FontSize =8
                    FontWeight =700
                    TabIndex =2
                    Name ="JF"
                    ControlSource ="=[SMABL_K]"
                    Format ="#,##0.00;-#,##0.00"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            OverlapFlags =81
                            TextAlign =0
                            TextFontFamily =2
                            Left =1881
                            Top =90
                            Width =1170
                            Height =435
                            FontSize =8
                            ForeColor =0
                            Name ="Label115"
                            Caption ="جمع كل فاكتور:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =243
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    BackStyle =1
                    NumeralShapes =2
                    Left =60
                    Top =495
                    Width =1761
                    Height =435
                    FontSize =8
                    FontWeight =700
                    TabIndex =3
                    Name ="HKH"
                    ControlSource ="MABL_HAZ"
                    Format ="#,##0.00;-#,##0.00"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            OverlapFlags =211
                            TextAlign =0
                            TextFontFamily =2
                            Left =1875
                            Top =495
                            Width =1095
                            Height =435
                            FontSize =8
                            ForeColor =0
                            Name ="Label116"
                            Caption ="هزينه خدمات:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =243
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    BackStyle =1
                    NumeralShapes =2
                    Left =60
                    Top =1305
                    Width =1761
                    Height =435
                    FontSize =8
                    FontWeight =700
                    TabIndex =4
                    Name ="TAKHFIF"
                    ControlSource ="TAKHFIF"
                    Format ="#,##0.00;-#,##0.00"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            OverlapFlags =211
                            TextAlign =0
                            TextFontFamily =2
                            Left =1875
                            Top =1305
                            Width =780
                            Height =435
                            FontSize =8
                            ForeColor =0
                            Name ="Label94"
                            Caption ="تخفيفات:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =243
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =60
                    Top =900
                    Width =1761
                    Height =435
                    FontSize =8
                    FontWeight =700
                    TabIndex =5
                    BackColor =16777164
                    Name ="GHABEL"
                    ControlSource ="=[JF]+[MABL_HAZ]"
                    Format ="#,##0.00;-#,##0.00"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            OverlapFlags =211
                            TextAlign =0
                            TextFontFamily =2
                            Left =1875
                            Top =907
                            Width =1515
                            Height =435
                            FontSize =8
                            ForeColor =0
                            Name ="Label117"
                            Caption ="مبلغ قابل پرداخت:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =243
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =60
                    Top =2115
                    Width =1761
                    Height =435
                    FontSize =8
                    FontWeight =700
                    TabIndex =6
                    BackColor =16777164
                    Name ="Text98"
                    ControlSource ="=[GHABEL]-[TAKHFIF]-[NPAR]"
                    Format ="#,##0.00;#,##0.00-"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            OverlapFlags =211
                            TextAlign =0
                            TextFontFamily =2
                            Left =1875
                            Top =2115
                            Width =945
                            Height =435
                            FontSize =8
                            ForeColor =0
                            Name ="Label99"
                            Caption ="مانده فاكتور:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin Line
                    BorderWidth =2
                    OverlapFlags =83
                    Left =3420
                    Top =30
                    Width =0
                    Height =2551
                    BorderColor =0
                    Name ="Line118"
                End
                Begin Line
                    BorderWidth =2
                    OverlapFlags =83
                    Width =0
                    Height =2611
                    BorderColor =0
                    Name ="Line119"
                End
                Begin Line
                    BorderWidth =3
                    OverlapFlags =83
                    Left =15
                    Top =2595
                    Width =10540
                    BorderColor =0
                    Name ="Line120"
                End
                Begin Line
                    OverlapFlags =115
                    Top =480
                    Width =3450
                    BorderColor =0
                    Name ="Line121"
                End
                Begin Line
                    OverlapFlags =115
                    Top =900
                    Width =3450
                    BorderColor =0
                    Name ="Line122"
                End
                Begin Line
                    OverlapFlags =115
                    Top =1320
                    Width =3450
                    BorderColor =0
                    Name ="Line123"
                End
                Begin Line
                    OverlapFlags =115
                    Top =1740
                    Width =3450
                    BorderColor =0
                    Name ="Line124"
                End
                Begin Line
                    OverlapFlags =115
                    Top =2130
                    Width =3450
                    BorderColor =0
                    Name ="Line125"
                End
                Begin Line
                    OverlapFlags =115
                    Left =1845
                    Top =30
                    Width =0
                    Height =2566
                    BorderColor =0
                    Name ="Line126"
                End
                Begin Line
                    BorderWidth =3
                    OverlapFlags =83
                    Left =10575
                    Width =0
                    Height =2626
                    BorderColor =0
                    Name ="Line127"
                End
                Begin Line
                    OverlapFlags =115
                    Left =3405
                    Top =900
                    Width =7143
                    BorderColor =0
                    Name ="Line128"
                End
                Begin Label
                    OverlapFlags =81
                    TextAlign =3
                    TextFontFamily =2
                    Left =9660
                    Top =945
                    Width =840
                    Height =405
                    FontSize =8
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label148"
                    Caption ="توضيحات:"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    OverlapFlags =81
                    TextAlign =3
                    TextFontFamily =2
                    Left =3495
                    Top =945
                    Width =6150
                    Height =1605
                    FontSize =8
                    ForeColor =0
                    Name ="COMM"
                    Caption ="."
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    OverlapFlags =81
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Top =2655
                    Width =10578
                    Height =540
                    FontSize =8
                    TabIndex =7
                    BorderColor =32768
                    Name ="Text153"
                    ControlSource ="=[Forms]![Baseknow]![HIGH_D]"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    TextFontFamily =2
                    Left =4371
                    Top =793
                    FontSize =8
                    TabIndex =8
                    Name ="mab"
                    ControlSource ="mab"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    TextFontFamily =2
                    Left =3918
                    Top =396
                    FontSize =8
                    TabIndex =9
                    Name ="CountOfMABL"
                    ControlSource ="CountOfMABL"
                    FontName ="WeblogmaYekan"
                End
            End
        End
    End
End
CodeBehindForm
' See "LIST_INVOICE_KHARED.cls"
