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
    Width =10862
    DatasheetFontHeight =10
    ItemSuffix =203
    Left =270
    Top =210
    DatasheetGridlinesColor =12632256
    Tag ="@Forms___F_MENU_GOZARESH_FROOSH___DT1 bigint= [Forms]![F_MENU_GOZARESH_FROOSH]!["
        "DT1] , @Forms___F_MENU_GOZARESH_FROOSH___DT2 bigint = [Forms]![F_MENU_GOZARESH_F"
        "ROOSH]![DT2],  @Forms___F_MENU_GOZARESH_FROOSH___DEPART=[Forms]![F_MENU_GOZARESH"
        "_FROOSH]![DEPART] , @Forms___F_MENU_GOZARESH_FROOSH___USERR=[Forms]![F_MENU_GOZA"
        "RESH_FROOSH]![USERR] , @Forms___F_MENU_GOZARESH_FROOSH___SSHIFT=[Forms]![F_MENU_"
        "GOZARESH_FROOSH]![SSHIFT]"
    RecSrcDt = Begin
        0x293b73b35164e440
    End
    InputParameters ="@Forms___F_MENU_GOZARESH_FROOSH___DT1 bigint= [Forms]![F_MENU_GOZARESH_FROOSH]!["
        "DT1] , @Forms___F_MENU_GOZARESH_FROOSH___DT2 bigint = [Forms]![F_MENU_GOZARESH_F"
        "ROOSH]![DT2],  @Forms___F_MENU_GOZARESH_FROOSH___DEPART=[Forms]![F_MENU_GOZARESH"
        "_FROOSH]![DEPART] , @Forms___F_MENU_GOZARESH_FROOSH___USERR=[Forms]![F_MENU_GOZA"
        "RESH_FROOSH]![USERR] , @Forms___F_MENU_GOZARESH_FROOSH___SSHIFT=[Forms]![F_MENU_"
        "GOZARESH_FROOSH]![SSHIFT]"
    RecordSource ="Q_FROOSH_DAYLY4"
    Caption ="گزارش فروش كالاها به تفكيك انبارها"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Arial (Arabic)"
    UnknownProp = {258 ,66 ,12 ,4 ,850 } Begin
        0x400046006f0072006d0073005f005f005f0046005f004d0045004e0055005f00 ,
        0x47004f005a00410052004500530048005f00460052004f004f00530048005f00 ,
        0x5f005f00440054003100200062006900670069006e0074003d0020005b004600 ,
        0x6f0072006d0073005d0021005b0046005f004d0045004e0055005f0047004f00 ,
        0x5a00410052004500530048005f00460052004f004f00530048005d0021005b00 ,
        0x4400540031005d0020002c002000400046006f0072006d0073005f005f005f00 ,
        0x46005f004d0045004e0055005f0047004f005a00410052004500530048005f00 ,
        0x460052004f004f00530048005f005f005f004400540032002000620069006700 ,
        0x69006e00740020003d0020005b0046006f0072006d0073005d0021005b004600 ,
        0x5f004d0045004e0055005f0047004f005a00410052004500530048005f004600 ,
        0x52004f004f00530048005d0021005b004400540032005d002c00200020004000 ,
        0x46006f0072006d0073005f005f005f0046005f004d0045004e0055005f004700 ,
        0x4f005a00410052004500530048005f00460052004f004f00530048005f005f00 ,
        0x5f004400450050004100520054003d005b0046006f0072006d0073005d002100 ,
        0x5b0046005f004d0045004e0055005f0047004f005a0041005200450053004800 ,
        0x5f00460052004f004f00530048005d0021005b00440045005000410052005400 ,
        0x5d0020002c002000400046006f0072006d0073005f005f005f0046005f004d00 ,
        0x45004e0055005f0047004f005a00410052004500530048005f00460052004f00 ,
        0x4f00530048005f005f005f00550053004500520052003d005b0046006f007200 ,
        0x6d0073005d0021005b0046005f004d0045004e0055005f0047004f005a004100 ,
        0x52004500530048005f00460052004f004f00530048005d0021005b0055005300 ,
        0x4500520052005d0020002c002000400046006f0072006d0073005f005f005f00 ,
        0x46005f004d0045004e0055005f0047004f005a00410052004500530048005f00 ,
        0x460052004f004f00530048005f005f005f005300530048004900460054003d00 ,
        0x5b0046006f0072006d0073005d0021005b0046005f004d0045004e0055005f00 ,
        0x47004f005a00410052004500530048005f00460052004f004f00530048005d00 ,
        0x21005b005300530048004900460054005d00
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
            ControlSource ="DEPATMAN"
        End
        Begin BreakLevel
            ControlSource ="NUMBER"
        End
        Begin BreakLevel
            ControlSource ="DATE_N"
        End
        Begin PageHeader
            Height =1684
            Name ="PageHeader"
            Begin
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =10260
                    Top =857
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
                    Left =9480
                    Top =863
                    Width =735
                    Height =765
                    FontSize =10
                    ForeColor =0
                    Name ="DATE_S Label"
                    Caption ="كد شخص"
                    FontName ="Traffic"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="DATE_S_Label"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =6811
                    Top =875
                    Width =2656
                    Height =765
                    FontSize =10
                    ForeColor =0
                    Name ="SHARH Label"
                    Caption ="نام شخص"
                    FontName ="Traffic"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="SHARH_Label"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =3904
                    Top =853
                    Width =1221
                    Height =765
                    FontSize =10
                    ForeColor =0
                    Name ="BED Label"
                    Caption ="مبلغ قابل پرداخت"
                    FontName ="Traffic"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="BED_Label"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =2835
                    Top =850
                    Width =1026
                    Height =765
                    FontSize =10
                    ForeColor =0
                    Name ="BES Label"
                    Caption ="نقد"
                    FontName ="Traffic"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="BES_Label"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =1485
                    Top =853
                    Width =1296
                    Height =765
                    FontSize =10
                    ForeColor =0
                    Name ="MAND Label"
                    Caption ="چك "
                    FontName ="Traffic"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="MAND_Label"
                End
                Begin Line
                    BorderWidth =2
                    Top =845
                    Width =10861
                    Name ="Line26"
                End
                Begin Line
                    BorderWidth =2
                    Top =1670
                    Width =10861
                    Name ="Line27"
                End
                Begin Line
                    Left =10245
                    Top =863
                    Width =0
                    Height =794
                    Name ="Line28"
                End
                Begin Line
                    Left =9495
                    Top =866
                    Width =0
                    Height =794
                    Name ="Line29"
                End
                Begin Line
                    Left =6750
                    Top =860
                    Width =0
                    Height =794
                    Name ="Line32"
                End
                Begin Line
                    Left =3885
                    Top =875
                    Width =0
                    Height =794
                    Name ="Line33"
                End
                Begin Line
                    Left =2806
                    Top =860
                    Width =0
                    Height =794
                    Name ="Line34"
                End
                Begin Line
                    BorderWidth =2
                    Top =863
                    Width =0
                    Height =794
                    Name ="Line36"
                End
                Begin Line
                    BorderWidth =2
                    Left =10830
                    Top =863
                    Width =0
                    Height =794
                    Name ="Line37"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =75
                    Top =442
                    Width =1021
                    Height =360
                    FontSize =10
                    Name ="Text52"
                    ControlSource ="=[Forms]![F_MENU_GOZARESH_FROOSH]![DT2]"
                    FontName ="Zar"
                    InputMask ="####/##/##"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1827
                    Top =440
                    Width =1086
                    Height =360
                    FontSize =10
                    TabIndex =1
                    Name ="Text53"
                    ControlSource ="=[Forms]![F_MENU_GOZARESH_FROOSH]![DT1]"
                    FontName ="Zar"
                    InputMask ="####/##/##"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =1144
                    Top =436
                    Width =660
                    Height =360
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="Label54"
                    Caption ="تا تاريخ:"
                    FontName ="Zar"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =2947
                    Top =438
                    Width =690
                    Height =360
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="Label55"
                    Caption ="از تاريخ:"
                    FontName ="Zar"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =6105
                    Top =860
                    Width =630
                    Height =765
                    FontSize =10
                    ForeColor =0
                    Name ="Label72"
                    Caption ="شماره فاكتور"
                    FontName ="Traffic"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    Left =6090
                    Top =866
                    Width =0
                    Height =794
                    Name ="Line73"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =5160
                    Top =860
                    Width =915
                    Height =765
                    FontSize =10
                    ForeColor =0
                    Name ="Label76"
                    Caption ="تاريخ فاكتور"
                    FontName ="Traffic"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    Left =5145
                    Top =866
                    Width =0
                    Height =794
                    Name ="Line77"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =45
                    Top =853
                    Width =1356
                    Height =765
                    FontSize =10
                    ForeColor =0
                    Name ="Label80"
                    Caption ="واريزي/حواله"
                    FontName ="Traffic"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    Left =1456
                    Top =860
                    Width =0
                    Height =794
                    Name ="Line81"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =10
                    Left =3628
                    Width =1581
                    Height =540
                    FontSize =14
                    TabIndex =2
                    Name ="DEPNAME"
                    ControlSource ="DEPNAME"
                    FontName ="Titr"
                    Begin
                        Begin Label
                            TextAlign =0
                            TextFontFamily =10
                            Left =5263
                            Top =7
                            Width =1560
                            Height =540
                            FontSize =14
                            ForeColor =0
                            Name ="Label89"
                            Caption ="گزارش فروش"
                            FontName ="Titr"
                        End
                    End
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =1
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =8220
                    Top =420
                    Height =375
                    FontSize =10
                    TabIndex =3
                    Name ="USS"
                    ControlSource ="USS"
                    FontName ="Zar"
                    Begin
                        Begin Label
                            TextAlign =0
                            TextFontFamily =10
                            Left =9977
                            Top =420
                            Width =495
                            Height =360
                            FontSize =10
                            ForeColor =0
                            Name ="Label87"
                            Caption ="كاربر:"
                            FontName ="Zar"
                        End
                    End
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =1
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =8220
                    Height =375
                    FontSize =10
                    TabIndex =4
                    Name ="SHIFT"
                    ControlSource ="SHNAME"
                    FontName ="Zar"
                    Begin
                        Begin Label
                            TextAlign =0
                            TextFontFamily =10
                            Left =9977
                            Width =540
                            Height =360
                            FontSize =10
                            ForeColor =0
                            Name ="Label88"
                            Caption ="شيفت:"
                            FontName ="Zar"
                        End
                    End
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            CanGrow = NotDefault
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
                    Left =6870
                    Width =2595
                    Height =375
                    FontSize =10
                    Name ="SHARH"
                    ControlSource ="NAME"
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
                    Left =3906
                    Width =1221
                    Height =375
                    FontSize =10
                    TabIndex =1
                    Name ="SumOfMABL_K1"
                    ControlSource ="=[jamf]+[SumOfMABL_HAZ]-[SumOfTAKHFIF]"
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
                    Left =2837
                    Width =1012
                    Height =375
                    FontSize =10
                    TabIndex =2
                    Name ="SumOfM_NAGHD"
                    ControlSource ="SumOfM_NAGHD"
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
                    Left =1487
                    Width =1296
                    Height =375
                    FontSize =10
                    TabIndex =3
                    Name ="JAMCH1"
                    ControlSource ="=[jamch]"
                    Format ="Standard"
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
                    Left =10266
                    Width =546
                    Height =375
                    FontSize =10
                    TabIndex =4
                    Name ="N_S"
                    ControlSource ="=1"
                    FontName ="Zar"
                End
                Begin Line
                    Left =10245
                    Width =0
                    Height =419
                    Name ="Line38"
                End
                Begin Line
                    Left =9495
                    Width =0
                    Height =419
                    Name ="Line39"
                End
                Begin Line
                    Left =6750
                    Width =0
                    Height =419
                    Name ="Line42"
                End
                Begin Line
                    Left =3885
                    Width =0
                    Height =419
                    Name ="Line43"
                End
                Begin Line
                    Left =2806
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
                    Left =10830
                    Width =0
                    Height =419
                    Name ="Line47"
                End
                Begin Line
                    LineSlant = NotDefault
                    Top =419
                    Width =10861
                    BorderColor =4210752
                    Name ="Line48"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =10
                    ReadingOrder =1
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =9480
                    Width =741
                    Height =375
                    FontSize =8
                    TabIndex =5
                    Name ="Text71"
                    ControlSource ="hes"
                    FontName ="Zar"
                End
                Begin Line
                    Left =6090
                    Width =0
                    Height =419
                    Name ="Line74"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =6075
                    Width =651
                    Height =375
                    FontSize =8
                    TabIndex =6
                    Name ="Text75"
                    ControlSource ="NUMBER"
                    Format ="Fixed"
                    FontName ="Zar"
                End
                Begin Line
                    Left =5145
                    Width =0
                    Height =419
                    Name ="Line78"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5175
                    Width =906
                    Height =375
                    FontSize =10
                    TabIndex =7
                    Name ="Text79"
                    ControlSource ="DATE_N"
                    FontName ="Zar"
                    InputMask ="####/##/##"
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
                    TabIndex =8
                    Name ="HAVVAR"
                    ControlSource ="=[SumOfMABL_VAR]+[SumOfMABL_HAV]"
                    Format ="Standard"
                    FontName ="Zar"
                End
                Begin Line
                    Left =1456
                    Width =0
                    Height =419
                    Name ="Line83"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    RunningSum =2
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =7143
                    Width =741
                    Height =360
                    FontSize =10
                    TabIndex =9
                    Name ="SText91"
                    ControlSource ="=[jamf]+[SumOfMABL_HAZ]-[SumOfTAKHFIF]"
                    Format ="Standard"
                    FontName ="Zar"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    RunningSum =2
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3061
                    Width =622
                    Height =375
                    FontSize =10
                    TabIndex =10
                    Name ="SSumOfM_NAGHD"
                    ControlSource ="SumOfM_NAGHD"
                    Format ="Standard"
                    FontName ="Zar"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    RunningSum =2
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1757
                    Width =621
                    Height =375
                    FontSize =10
                    TabIndex =11
                    Name ="SHAVVAR"
                    ControlSource ="=[jamf]+[SumOfMABL_HAZ]-[SumOfTAKHFIF]-[SumOfM_NAGHD]-[jamch]"
                    Format ="Standard"
                    FontName ="Zar"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =6780
                    Width =201
                    Height =375
                    FontSize =10
                    TabIndex =12
                    Name ="khfr"
                    ControlSource ="=IIf([tag2]=13,\"ف\",\"خ\")"
                    FontName ="Zar"
                    Tag ="=IIf([Tag]=2;'ف';'خ')"
                End
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Left =6973
                    Top =56
                    TabIndex =13
                    Name ="TAG2"
                    ControlSource ="TAG"
                End
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Left =7596
                    Top =113
                    TabIndex =14
                    Name ="SumOfMABL_VAR"
                    ControlSource ="SumOfMABL_VAR"
                End
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Left =7035
                    Top =165
                    TabIndex =15
                    Name ="SumOfMABL_HAV"
                    ControlSource ="SumOfMABL_HAV"
                End
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Left =7425
                    Top =60
                    TabIndex =16
                    Name ="jamch"
                    ControlSource ="jamch"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    RunningSum =2
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Width =666
                    Height =375
                    FontSize =10
                    TabIndex =17
                    Name ="SJAMCH"
                    ControlSource ="jamch"
                    Format ="Standard"
                    FontName ="Zar"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    RunningSum =2
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1303
                    Width =666
                    Height =375
                    FontSize =10
                    TabIndex =18
                    Name ="SHAVL"
                    ControlSource ="=[SumOfMABL_VAR]+[SumOfMABL_HAV]"
                    Format ="Standard"
                    FontName ="Zar"
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            CanGrow = NotDefault
            ForceNewPage =2
            Height =5235
            OnFormat ="[Event Procedure]"
            Name ="GroupFooter3"
            Begin
                Begin Line
                    BorderWidth =3
                    Top =450
                    Width =10862
                    BorderColor =0
                    Name ="Line176"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    NumeralShapes =2
                    Left =60
                    Top =2487
                    Width =1531
                    Height =435
                    FontSize =10
                    ForeColor =0
                    Name ="JPAY"
                    Caption ="جمع مبالغ پرداختي:"
                    FontName ="Titr"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    NumeralShapes =2
                    Left =60
                    Top =597
                    Width =1531
                    Height =435
                    FontSize =10
                    ForeColor =0
                    Name ="HKH"
                    Caption ="هزينه خدمات:"
                    FontName ="Titr"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    NumeralShapes =2
                    Left =60
                    Top =1077
                    Width =1531
                    Height =405
                    FontSize =10
                    ForeColor =0
                    Name ="TF"
                    Caption ="تخفيفات:"
                    FontName ="Titr"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    NumeralShapes =2
                    Left =60
                    Top =2007
                    Width =1531
                    Height =435
                    FontSize =10
                    ForeColor =0
                    Name ="GABEL"
                    Caption ="مبلغ قابل پرداخت:"
                    FontName ="Titr"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    NumeralShapes =2
                    Left =68
                    Top =2907
                    Width =1531
                    Height =435
                    FontSize =10
                    ForeColor =0
                    Name ="MAN"
                    Caption ="مانده فاكتور:"
                    FontName ="Titr"
                End
                Begin Label
                    TextFontFamily =2
                    ReadingOrder =2
                    NumeralShapes =2
                    Left =1635
                    Top =2487
                    Width =1486
                    Height =435
                    FontSize =10
                    ForeColor =0
                    Name ="Label215"
                    Caption ="جمع مبالغ دريافتي:"
                    FontName ="Titr"
                End
                Begin Label
                    TextFontFamily =2
                    ReadingOrder =2
                    NumeralShapes =2
                    Left =1635
                    Top =597
                    Width =1486
                    Height =435
                    FontSize =10
                    ForeColor =0
                    Name ="Label216"
                    Caption ="هزينه خدمات:"
                    FontName ="Titr"
                End
                Begin Label
                    TextFontFamily =2
                    ReadingOrder =2
                    NumeralShapes =2
                    Left =1635
                    Top =1077
                    Width =1486
                    Height =405
                    FontSize =10
                    ForeColor =0
                    Name ="Label217"
                    Caption ="تخفيفات:"
                    FontName ="Titr"
                End
                Begin Label
                    TextFontFamily =2
                    ReadingOrder =2
                    NumeralShapes =2
                    Left =1635
                    Top =2007
                    Width =1486
                    Height =435
                    FontSize =10
                    ForeColor =0
                    Name ="Label218"
                    Caption ="مبلغ قابل دريافت:"
                    FontName ="Titr"
                End
                Begin Label
                    TextFontFamily =2
                    ReadingOrder =2
                    NumeralShapes =2
                    Left =1635
                    Top =2907
                    Width =1486
                    Height =435
                    FontSize =10
                    ForeColor =0
                    Name ="Label219"
                    Caption ="مانده بدهكاران:"
                    FontName ="Titr"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    NumeralShapes =2
                    Left =7470
                    Top =2472
                    Width =1531
                    Height =435
                    FontSize =10
                    ForeColor =0
                    Name ="CHK"
                    Caption ="."
                    FontName ="Titr"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    NumeralShapes =2
                    Left =7470
                    Top =597
                    Width =1531
                    Height =435
                    FontSize =10
                    ForeColor =0
                    Name ="NAGHD"
                    Caption ="."
                    FontName ="Titr"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    NumeralShapes =2
                    Left =7485
                    Top =1560
                    Width =1531
                    Height =405
                    FontSize =10
                    ForeColor =0
                    Name ="HAVAL"
                    Caption ="."
                    FontName ="Titr"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    NumeralShapes =2
                    Left =7470
                    Top =1992
                    Width =1531
                    Height =435
                    FontSize =10
                    ForeColor =0
                    Name ="VARIZ"
                    Caption ="."
                    FontName ="Titr"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    NumeralShapes =2
                    Left =7470
                    Top =2937
                    Width =1531
                    Height =435
                    FontSize =10
                    ForeColor =0
                    Name ="PAR"
                    Caption ="."
                    FontName ="Titr"
                End
                Begin Label
                    TextFontFamily =2
                    ReadingOrder =2
                    NumeralShapes =2
                    Left =9045
                    Top =2472
                    Width =1590
                    Height =435
                    FontSize =10
                    ForeColor =0
                    Name ="Label225"
                    Caption ="جمع چگهاي دريافتي:"
                    FontName ="Titr"
                End
                Begin Label
                    TextFontFamily =2
                    ReadingOrder =2
                    NumeralShapes =2
                    Left =9045
                    Top =600
                    Width =1785
                    Height =705
                    FontSize =10
                    ForeColor =0
                    Name ="Label226"
                    Caption ="مبلغ نقد فاكتورها  : منهاي نقد برگشت فروش"
                    FontName ="Titr"
                End
                Begin Label
                    TextFontFamily =2
                    ReadingOrder =2
                    NumeralShapes =2
                    Left =9060
                    Top =1560
                    Width =1486
                    Height =405
                    FontSize =10
                    ForeColor =0
                    Name ="Label227"
                    Caption ="مبلغ حواله:"
                    FontName ="Titr"
                End
                Begin Label
                    TextFontFamily =2
                    ReadingOrder =2
                    NumeralShapes =2
                    Left =9045
                    Top =1992
                    Width =1486
                    Height =435
                    FontSize =10
                    ForeColor =0
                    Name ="Label228"
                    Caption ="مبلغ واريزي:"
                    FontName ="Titr"
                End
                Begin Label
                    TextFontFamily =2
                    ReadingOrder =2
                    NumeralShapes =2
                    Left =9045
                    Top =2937
                    Width =1515
                    Height =435
                    FontSize =10
                    ForeColor =0
                    Name ="Label229"
                    Caption ="جمع مبالغ دريافتي:"
                    FontName ="Titr"
                End
                Begin Rectangle
                    BorderWidth =2
                    Top =510
                    Width =10831
                    Height =2889
                    Name ="Box232"
                End
                Begin Line
                    BorderWidth =1
                    Left =5160
                    Top =525
                    Width =0
                    Height =2880
                    BorderColor =0
                    Name ="Line233"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Left =9015
                    Top =525
                    Width =0
                    Height =2880
                    BorderColor =0
                    Name ="Line234"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Left =1620
                    Top =525
                    Width =0
                    Height =2880
                    BorderColor =0
                    Name ="Line235"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    NumeralShapes =2
                    Left =60
                    Top =1527
                    Width =1531
                    Height =405
                    FontSize =10
                    ForeColor =0
                    Name ="MBAA"
                    Caption ="ارزش افزوده:"
                    FontName ="Titr"
                End
                Begin Label
                    TextFontFamily =2
                    ReadingOrder =2
                    NumeralShapes =2
                    Left =1635
                    Top =1527
                    Width =1486
                    Height =405
                    FontSize =10
                    ForeColor =0
                    Name ="Label270"
                    Caption ="ارزش افزوده:"
                    FontName ="Titr"
                End
                Begin Line
                    BorderWidth =2
                    Left =5175
                    Top =2895
                    Width =5661
                    BorderColor =0
                    Name ="Line190"
                End
                Begin Subform
                    CanShrink = NotDefault
                    Top =3570
                    Width =10852
                    Height =300
                    Name ="LIST_PISH_ANBARS_DTL"
                    SourceObject ="Report.LIST_PISH_ANBARS_DTL"
                End
                Begin Subform
                    CanShrink = NotDefault
                    Top =3915
                    Width =10852
                    Height =360
                    TabIndex =1
                    Name ="CHEKHA"
                    SourceObject ="Report.CHECK_USERS_F"
                End
                Begin Subform
                    CanShrink = NotDefault
                    Top =4395
                    Width =10852
                    Height =360
                    TabIndex =2
                    Name ="BEDEHKAR"
                    SourceObject ="Report.BEDEHKARAN_USERS_F"
                End
                Begin Subform
                    CanShrink = NotDefault
                    Top =4875
                    Width =10852
                    Height =360
                    TabIndex =3
                    Name ="Child195"
                    SourceObject ="Report.BESTANKARAN_USER_F"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    NumeralShapes =2
                    Left =5160
                    Top =1080
                    Width =1531
                    Height =435
                    FontSize =10
                    ForeColor =16711935
                    Name ="LNAGHDDP"
                    Caption ="."
                    FontName ="Titr"
                End
                Begin Label
                    TextFontFamily =2
                    ReadingOrder =2
                    NumeralShapes =2
                    Left =6746
                    Top =1080
                    Width =1486
                    Height =435
                    FontSize =10
                    ForeColor =16711935
                    Name ="Label197"
                    Caption ="مانده صندوق:"
                    FontName ="Titr"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    NumeralShapes =2
                    Left =5159
                    Top =585
                    Width =1531
                    Height =435
                    FontSize =10
                    ForeColor =16711935
                    Name ="LNAGHDD"
                    Caption ="."
                    FontName ="Titr"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    NumeralShapes =2
                    Left =3510
                    Top =592
                    Width =1531
                    Height =435
                    FontSize =10
                    ForeColor =16711935
                    Name ="LNAGHDP"
                    Caption ="."
                    FontName ="Titr"
                End
                Begin Label
                    TextFontFamily =2
                    ReadingOrder =2
                    NumeralShapes =2
                    Left =6750
                    Top =615
                    Width =555
                    Height =390
                    FontSize =10
                    ForeColor =16711935
                    Name ="Label200"
                    Caption ="خزانه:"
                    FontName ="Titr"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    NumeralShapes =2
                    Left =3510
                    Top =1080
                    Width =1530
                    Height =1740
                    FontSize =10
                    ForeColor =16711935
                    Name ="Label201"
                    Caption ="مبلغ دريافت خزانه و  نقد فاكتورها منهاي برداشت خزانه  , برگشت فروش معادل مانده ص"
                        "ندوق است"
                    FontName ="Titr"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3884
                    Width =1237
                    Height =375
                    FontSize =10
                    TabIndex =4
                    Name ="SBED"
                    ControlSource ="=[SText91]"
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
                    Left =2820
                    Width =1026
                    Height =375
                    FontSize =10
                    TabIndex =5
                    Name ="SBES"
                    ControlSource ="=[SSumOfM_NAGHD]"
                    Format ="Standard"
                    FontName ="Zar"
                End
                Begin Line
                    Left =5130
                    Width =0
                    Height =419
                    Name ="Line60"
                End
                Begin Line
                    Left =3870
                    Width =0
                    Height =419
                    Name ="Line61"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =2806
                    Width =0
                    Height =419
                    Name ="Line62"
                End
                Begin Line
                    BorderWidth =2
                    Width =0
                    Height =419
                    Name ="Line64"
                End
                Begin Line
                    BorderWidth =2
                    Left =10830
                    Width =0
                    Height =419
                    Name ="Line67"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =6066
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
                    TabIndex =6
                    Name ="Text84"
                    ControlSource ="=[SHAVL]"
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
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1485
                    Width =1296
                    Height =375
                    FontSize =10
                    TabIndex =7
                    Name ="Text59"
                    ControlSource ="=[SJAMCH]"
                    Format ="Standard"
                    FontName ="Zar"
                End
                Begin Line
                    BorderWidth =2
                    Width =10862
                    Name ="Line86"
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
' See "LIST_FROOSH_ANBARS_DTL.cls"
