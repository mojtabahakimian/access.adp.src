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
    Width =10544
    DatasheetFontHeight =10
    ItemSuffix =161
    Left =-1785
    Top =315
    DatasheetGridlinesColor =12632256
    Tag ="@Forms___F_MENU_KHFR___DT1 =[Forms]![F_MENU_KHFR]![DT1] , @Forms___F_MENU_KHFR__"
        "_DT2  =[Forms]![F_MENU_KHFR]![DT2] ,  @Forms___F_MENU_KHFR___MMOIN  nvarchar(20)"
        "= [Forms]![F_MENU_KHFR]![MMOIN]"
    RecSrcDt = Begin
        0xb3ad52fa9046e340
    End
    InputParameters ="@Forms___F_MENU_KHFR___DT1 =[Forms]![F_MENU_KHFR]![DT1] , @Forms___F_MENU_KHFR__"
        "_DT2  =[Forms]![F_MENU_KHFR]![DT2] ,  @Forms___F_MENU_KHFR___MMOIN  nvarchar(20)"
        "= [Forms]![F_MENU_KHFR]![MMOIN]"
    RecordSource ="QSL_LIST_INVOICE_FROOSH2"
    Caption ="فاكتور فروش"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Arial (Arabic)"
    OnError ="[Event Procedure]"
    UnknownProp = {258 ,66 ,12 ,4 ,382 } Begin
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
        0x3d0020005b0046006f0072006d0073005d0021005b0046005f004d0045004e00 ,
        0x55005f004b004800460052005d0021005b004d004d004f0049004e005d00
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
            Height =3075
            Name ="GroupHeader2"
            Begin
                Begin Label
                    TextFontFamily =10
                    Left =3855
                    Width =1380
                    Height =540
                    FontSize =14
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label88"
                    Caption ="فاكتور فروش"
                    FontName ="Titr"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2355
                    Top =495
                    Width =5043
                    Height =540
                    FontSize =14
                    BorderColor =32768
                    Name ="Text90"
                    ControlSource ="=[Forms]![Baseknow]![WIDTH_D]"
                    FontName ="Titr"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =45
                    Top =540
                    Width =1506
                    Height =450
                    FontSize =12
                    TabIndex =1
                    BorderColor =32768
                    Name ="DATE_N"
                    ControlSource ="DATE_N"
                    FontName ="Traffic"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =10
                            Left =1588
                            Top =540
                            Width =660
                            Height =480
                            FontSize =12
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label92"
                            Caption ="تاريخ :"
                            FontName ="Traffic"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =7395
                    Top =510
                    Width =1686
                    Height =435
                    FontSize =12
                    TabIndex =2
                    BorderColor =32768
                    Name ="HEAD_LST.NUMBER"
                    ControlSource ="NUMBER"
                    FontName ="Traffic"
                    EventProcPrefix ="HEAD_LST_NUMBER"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =10
                            Left =9132
                            Top =510
                            Width =1275
                            Height =480
                            FontSize =12
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label93"
                            Caption ="شماره فاكتور:"
                            FontName ="Traffic"
                        End
                    End
                End
                Begin Line
                    BorderWidth =3
                    Top =1050
                    Width =10521
                    BorderColor =0
                    Name ="Line94"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2475
                    Top =1062
                    Width =7146
                    Height =450
                    FontSize =12
                    TabIndex =3
                    BorderColor =32768
                    Name ="CUST_NO"
                    ControlSource ="NAME"
                    FontName ="Traffic"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =10
                            Left =9657
                            Top =1062
                            Width =780
                            Height =480
                            FontSize =12
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label95"
                            Caption ="خريدار:"
                            FontName ="Traffic"
                        End
                    End
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =30
                    Top =1050
                    Width =1431
                    Height =450
                    FontSize =12
                    TabIndex =4
                    BorderColor =32768
                    Name ="BEDN"
                    ControlSource ="BEDN"
                    FontName ="Traffic"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =10
                            Left =1452
                            Top =1050
                            Width =405
                            Height =480
                            FontSize =12
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label96"
                            Caption ="كد:"
                            FontName ="Traffic"
                        End
                    End
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2415
                    Top =1629
                    Width =7191
                    Height =450
                    FontSize =12
                    TabIndex =5
                    BorderColor =32768
                    Name ="ADDRESS"
                    ControlSource ="ADDRESS"
                    FontName ="Traffic"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =10
                            Left =9642
                            Top =1629
                            Width =765
                            Height =480
                            FontSize =12
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label97"
                            Caption ="آدرس:"
                            FontName ="Traffic"
                        End
                    End
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =30
                    Top =1620
                    Width =1401
                    Height =450
                    FontSize =12
                    TabIndex =6
                    BorderColor =32768
                    Name ="TEL"
                    ControlSource ="TEL"
                    FontName ="Traffic"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =10
                            Left =1452
                            Top =1620
                            Width =555
                            Height =480
                            FontSize =12
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label98"
                            Caption ="تلفن:"
                            FontName ="Traffic"
                        End
                    End
                End
                Begin Line
                    BorderWidth =3
                    Top =2550
                    Width =10521
                    BorderColor =0
                    Name ="Line99"
                End
                Begin Line
                    BorderWidth =3
                    Left =10485
                    Top =2565
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line130"
                End
                Begin Line
                    Left =9915
                    Top =2565
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line131"
                End
                Begin Line
                    Left =9345
                    Top =2565
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line132"
                End
                Begin Line
                    Left =5985
                    Top =2565
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line133"
                End
                Begin Line
                    Left =4845
                    Top =2565
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line134"
                End
                Begin Line
                    Left =4020
                    Top =2565
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line135"
                End
                Begin Line
                    Left =3075
                    Top =2565
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line136"
                End
                Begin Line
                    Left =1575
                    Top =2565
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line137"
                End
                Begin Line
                    BorderWidth =3
                    Top =2565
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line138"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =480
                    Top =2610
                    Width =585
                    Height =405
                    FontSize =10
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label139"
                    Caption ="مبــلغ"
                    FontName ="Traffic"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =2049
                    Top =2606
                    Width =585
                    Height =405
                    FontSize =10
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label140"
                    Caption ="فـــي"
                    FontName ="Traffic"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =3165
                    Top =2610
                    Width =810
                    Height =405
                    FontSize =10
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label141"
                    Caption ="مقدار كل"
                    FontName ="Traffic"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =4170
                    Top =2610
                    Width =525
                    Height =405
                    FontSize =10
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label142"
                    Caption ="مقدار"
                    FontName ="Traffic"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =5047
                    Top =2610
                    Width =750
                    Height =405
                    FontSize =10
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label143"
                    Caption ="واحد كالا"
                    FontName ="Traffic"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =6637
                    Top =2610
                    Width =1710
                    Height =405
                    FontSize =10
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label144"
                    Caption ="شــــــــــرح  كالا"
                    FontName ="Traffic"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =9420
                    Top =2610
                    Width =375
                    Height =405
                    FontSize =10
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label145"
                    Caption ="انبار"
                    FontName ="Traffic"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =9921
                    Top =2606
                    Width =540
                    Height =405
                    FontSize =10
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label146"
                    Caption ="رديف"
                    FontName ="Traffic"
                End
                Begin Line
                    BorderWidth =3
                    Top =3075
                    Width =10521
                    BorderColor =0
                    Name ="Line147"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =30
                    Top =2100
                    Width =9366
                    Height =450
                    FontSize =12
                    TabIndex =7
                    BorderColor =32768
                    Name ="MOLAH"
                    ControlSource ="MOLAH"
                    FontName ="Traffic"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =10
                            Left =9420
                            Top =2100
                            Width =1035
                            Height =480
                            FontSize =12
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label155"
                            Caption ="ملاحظات:"
                            FontName ="Traffic"
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
                    TextAlign =2
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =9377
                    Top =16
                    Width =504
                    Height =405
                    FontSize =10
                    FontWeight =700
                    BorderColor =32768
                    Name ="INVO_LST.ANBAR"
                    ControlSource ="IANBAR"
                    FontName ="Traffic"
                    EventProcPrefix ="INVO_LST_ANBAR"
                End
                Begin TextBox
                    DecimalPlaces =0
                    RunningSum =1
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =9931
                    Top =16
                    Width =504
                    Height =405
                    FontSize =10
                    FontWeight =700
                    TabIndex =1
                    BorderColor =32768
                    Name ="RADIF"
                    ControlSource ="=1"
                    FontName ="Traffic"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5997
                    Top =16
                    Width =3324
                    Height =405
                    FontSize =10
                    FontWeight =700
                    TabIndex =2
                    BorderColor =32768
                    Name ="CODE1"
                    ControlSource ="=IIf([Forms]![baseknow]![codeview]=-1,[kala] & \" \" & [code],[kala])"
                    FontName ="Traffic"
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =2
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4023
                    Top =16
                    Width =789
                    Height =405
                    FontSize =10
                    FontWeight =700
                    TabIndex =3
                    BorderColor =32768
                    Name ="MEGH"
                    ControlSource ="MEGH"
                    FontName ="Traffic"
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =2
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3090
                    Top =16
                    Width =894
                    Height =405
                    FontSize =10
                    FontWeight =700
                    TabIndex =4
                    BorderColor =32768
                    Name ="MEGHk"
                    ControlSource ="MEGHk"
                    FontName ="Traffic"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1594
                    Top =16
                    Width =1449
                    Height =405
                    FontSize =10
                    FontWeight =700
                    TabIndex =5
                    BorderColor =32768
                    Name ="MABL"
                    ControlSource ="MABL"
                    Format ="#,###"
                    FontName ="Traffic"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =30
                    Top =16
                    Width =1509
                    Height =405
                    FontSize =10
                    FontWeight =700
                    TabIndex =6
                    BorderColor =32768
                    Name ="MABL_K"
                    ControlSource ="MABL_K"
                    Format ="#,###"
                    FontName ="Traffic"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4845
                    Top =15
                    Width =1110
                    Height =405
                    FontSize =10
                    FontWeight =700
                    TabIndex =7
                    BorderColor =32768
                    Name ="VAHED_K"
                    ControlSource ="TNAMES"
                    FontName ="Traffic"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Top =450
                    Width =10506
                    BorderColor =0
                    Name ="Line100"
                End
                Begin Line
                    BorderWidth =3
                    Left =10485
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line101"
                End
                Begin Line
                    BorderWidth =1
                    Left =9915
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line102"
                End
                Begin Line
                    BorderWidth =1
                    Left =9345
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line103"
                End
                Begin Line
                    BorderWidth =1
                    Left =5985
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line104"
                End
                Begin Line
                    BorderWidth =1
                    Left =4845
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line105"
                End
                Begin Line
                    BorderWidth =1
                    Left =4020
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line106"
                End
                Begin Line
                    BorderWidth =1
                    Left =3075
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line107"
                End
                Begin Line
                    BorderWidth =1
                    Left =1575
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line108"
                End
                Begin Line
                    BorderWidth =2
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
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =340
                    Width =819
                    Height =405
                    FontSize =10
                    FontWeight =700
                    TabIndex =8
                    BorderColor =32768
                    Name ="SMABL_K"
                    ControlSource ="MABL_K"
                    Format ="#,###"
                    FontName ="Traffic"
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
                    Top =30
                    Width =10465
                    BorderColor =0
                    Name ="Line110"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3938
                    Top =95
                    Width =5286
                    Height =750
                    FontSize =10
                    FontWeight =700
                    BorderColor =32768
                    Name ="Text111"
                    ControlSource ="=ALPHANUM([GHABEL]) & \" \" & \"ريال\""
                    FontName ="Koodak"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =10
                            Left =9207
                            Top =83
                            Width =1200
                            Height =405
                            FontSize =10
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label112"
                            Caption ="مبلغ به حروف:"
                            FontName ="Traffic"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    BackStyle =1
                    NumeralShapes =2
                    Left =75
                    Top =1710
                    Width =2151
                    Height =435
                    FontSize =16
                    FontWeight =700
                    TabIndex =1
                    Name ="NPAR"
                    ControlSource ="=[M_NAGHD]+[MABL_VAR]+[MABL_HAV]+[mab]"
                    Format ="#,##0.00;-#,##0.00"
                    FontName ="Kamran"
                    Begin
                        Begin Label
                            TextAlign =0
                            TextFontFamily =10
                            Left =2281
                            Top =1717
                            Width =1695
                            Height =435
                            FontSize =14
                            ForeColor =0
                            Name ="Label86"
                            Caption ="جمع مبالغ پرداختي:"
                            FontName ="Kamran"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    BackStyle =1
                    NumeralShapes =2
                    Left =76
                    Top =90
                    Width =2151
                    Height =435
                    FontSize =16
                    FontWeight =700
                    TabIndex =2
                    Name ="JF"
                    ControlSource ="=[SMABL_K]"
                    Format ="#,##0.00;-#,##0.00"
                    FontName ="Kamran"
                    Begin
                        Begin Label
                            TextAlign =0
                            TextFontFamily =10
                            Left =2286
                            Top =90
                            Width =1425
                            Height =435
                            FontSize =14
                            ForeColor =0
                            Name ="Label115"
                            Caption ="جمع كل فاكتور:"
                            FontName ="Kamran"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    BackStyle =1
                    NumeralShapes =2
                    Left =75
                    Top =495
                    Width =2151
                    Height =435
                    FontSize =16
                    FontWeight =700
                    TabIndex =3
                    Name ="HKH"
                    ControlSource ="MABL_HAZ"
                    Format ="#,##0.00;-#,##0.00"
                    FontName ="Kamran"
                    Begin
                        Begin Label
                            TextAlign =0
                            TextFontFamily =10
                            Left =2280
                            Top =495
                            Width =1410
                            Height =435
                            FontSize =14
                            ForeColor =0
                            Name ="Label116"
                            Caption ="3% تجميع عوارض:"
                            FontName ="Kamran"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    BackStyle =1
                    NumeralShapes =2
                    Left =75
                    Top =1305
                    Width =2151
                    Height =435
                    FontSize =16
                    FontWeight =700
                    TabIndex =4
                    Name ="TAKHFIF"
                    ControlSource ="TAKHFIF"
                    Format ="#,##0.00;-#,##0.00"
                    FontName ="Kamran"
                    Begin
                        Begin Label
                            TextAlign =0
                            TextFontFamily =10
                            Left =2280
                            Top =1305
                            Width =780
                            Height =435
                            FontSize =14
                            ForeColor =0
                            Name ="Label94"
                            Caption ="تخفيفات:"
                            FontName ="Kamran"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =75
                    Top =900
                    Width =2151
                    Height =435
                    FontSize =16
                    FontWeight =700
                    TabIndex =5
                    BackColor =16777164
                    Name ="GHABEL"
                    ControlSource ="=[JF]+[MABL_HAZ]"
                    Format ="#,##0.00;-#,##0.00"
                    FontName ="Kamran"
                    Begin
                        Begin Label
                            TextAlign =0
                            TextFontFamily =10
                            Left =2280
                            Top =907
                            Width =1515
                            Height =435
                            FontSize =14
                            ForeColor =0
                            Name ="Label117"
                            Caption ="مبلغ قابل پرداخت:"
                            FontName ="Kamran"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =75
                    Top =2115
                    Width =2151
                    Height =435
                    FontSize =16
                    FontWeight =700
                    TabIndex =6
                    BackColor =16777164
                    Name ="Text98"
                    ControlSource ="=[GHABEL]-[TAKHFIF]-[NPAR]"
                    Format ="#,##0.00;#,##0.00-"
                    FontName ="Kamran"
                    Begin
                        Begin Label
                            TextAlign =0
                            TextFontFamily =10
                            Left =2280
                            Top =2115
                            Width =945
                            Height =435
                            FontSize =14
                            ForeColor =0
                            Name ="Label99"
                            Caption ="مانده فاكتور:"
                            FontName ="Kamran"
                        End
                    End
                End
                Begin Line
                    BorderWidth =2
                    Left =3825
                    Top =30
                    Width =0
                    Height =2551
                    BorderColor =0
                    Name ="Line118"
                End
                Begin Line
                    BorderWidth =2
                    Width =0
                    Height =2611
                    BorderColor =0
                    Name ="Line119"
                End
                Begin Line
                    BorderWidth =3
                    Left =15
                    Top =2595
                    Width =10450
                    BorderColor =0
                    Name ="Line120"
                End
                Begin Line
                    Left =15
                    Top =480
                    Width =3840
                    BorderColor =0
                    Name ="Line121"
                End
                Begin Line
                    Left =15
                    Top =900
                    Width =3840
                    BorderColor =0
                    Name ="Line122"
                End
                Begin Line
                    Left =15
                    Top =1320
                    Width =3840
                    BorderColor =0
                    Name ="Line123"
                End
                Begin Line
                    Left =15
                    Top =1740
                    Width =3840
                    BorderColor =0
                    Name ="Line124"
                End
                Begin Line
                    Left =15
                    Top =2130
                    Width =3840
                    BorderColor =0
                    Name ="Line125"
                End
                Begin Line
                    Left =2250
                    Top =30
                    Width =0
                    Height =2566
                    BorderColor =0
                    Name ="Line126"
                End
                Begin Line
                    BorderWidth =3
                    Left =10485
                    Width =0
                    Height =2626
                    BorderColor =0
                    Name ="Line127"
                End
                Begin Line
                    Left =3315
                    Top =900
                    Width =7143
                    BorderColor =0
                    Name ="Line128"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =10
                    Left =9570
                    Top =945
                    Width =840
                    Height =405
                    FontSize =10
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label148"
                    Caption ="توضيحات:"
                    FontName ="Traffic"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =10
                    Left =3945
                    Top =945
                    Width =5610
                    Height =1605
                    FontSize =10
                    ForeColor =0
                    Name ="COMM"
                    Caption ="."
                    FontName ="Koodak"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =195
                    Top =2655
                    Width =10293
                    Height =540
                    FontSize =14
                    TabIndex =7
                    BorderColor =32768
                    Name ="Text153"
                    ControlSource ="=[Forms]![Baseknow]![HIGH_D]"
                    FontName ="Titr"
                End
                Begin TextBox
                    Visible = NotDefault
                    Left =4281
                    Top =793
                    TabIndex =8
                    Name ="mab"
                    ControlSource ="mab"
                End
                Begin TextBox
                    Visible = NotDefault
                    Left =3828
                    Top =396
                    TabIndex =9
                    Name ="CountOfMABL"
                    ControlSource ="CountOfMABL"
                End
            End
        End
    End
End
CodeBehindForm
' See "LIST_INVOICE_FROOSH.cls"
