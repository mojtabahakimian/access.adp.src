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
    Width =10656
    DatasheetFontHeight =10
    ItemSuffix =162
    Left =930
    Top =1050
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xed4d343750abe240
    End
    RecordSource ="SELECT CHRE_LST.N_SERI, CHRE_LST.BANK, CHRE_LST.DATE_S, CHRE_LST.RADIF, CHRE_LST"
        ".N_MOIN, PAY_GETD.SHOBEH, PAY_GETD.MABL, PAY_GETD.DATE_S AS Expr1, PAY_GETD.N_KO"
        "L3, PAY_GETD.N_MOIN3, CHKREC_H.DATE, CHKREC_H.MOLAH, TCOD_BANKS.NAMES + ' ' + PA"
        "Y_GETD.SHOBEH AS BAN, TDETA_HES.NAME AS HES FROM TCOD_BANKS INNER JOIN PAY_GETD "
        "RIGHT OUTER JOIN CHKREC_H INNER JOIN CHRE_LST ON CHKREC_H.DATE = CHRE_LST.DATE O"
        "N PAY_GETD.DATE_S = CHRE_LST.DATE_S AND PAY_GETD.BANK = CHRE_LST.BANK AND PAY_GE"
        "TD.N_SERI = CHRE_LST.N_SERI ON TCOD_BANKS.CODE = CHRE_LST.BANK INNER JOIN TDETA_"
        "HES ON PAY_GETD.N_KOL = TDETA_HES.N_KOL AND PAY_GETD.N_MOIN = TDETA_HES.NUMBER A"
        "ND PAY_GETD.N_TAF = TDETA_HES.TNUMBER"
    Caption ="برگه اعلام وصول"
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
            GroupHeader = NotDefault
            GroupFooter = NotDefault
            ControlSource ="DATE"
        End
        Begin BreakLevel
            ControlSource ="DATE_S"
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =2295
            Name ="GroupHeader2"
            Begin
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =3253
                    Width =4005
                    Height =540
                    FontSize =14
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label88"
                    Caption ="برگه اعلام وصول چكهاي دريافتي"
                    FontName ="Titr"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =8355
                    Top =1155
                    Width =1506
                    Height =450
                    FontSize =12
                    BorderColor =32768
                    Name ="DATE"
                    ControlSource ="DATE"
                    FontName ="Traffic"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =10
                            Left =9898
                            Top =1155
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
                Begin Line
                    BorderWidth =3
                    Top =1770
                    Width =10626
                    BorderColor =0
                    Name ="Line99"
                End
                Begin Line
                    BorderWidth =3
                    Left =10605
                    Top =1785
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line130"
                End
                Begin Line
                    Left =10020
                    Top =1785
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line131"
                End
                Begin Line
                    Left =9993
                    Top =1785
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line132"
                End
                Begin Line
                    Left =8928
                    Top =1785
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line133"
                End
                Begin Line
                    Left =6200
                    Top =1785
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line134"
                End
                Begin Line
                    Left =4920
                    Top =1785
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line135"
                End
                Begin Line
                    Left =3255
                    Top =1785
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line137"
                End
                Begin Line
                    BorderWidth =3
                    Left =30
                    Top =1785
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line138"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =630
                    Top =1830
                    Width =1350
                    Height =405
                    FontSize =10
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label139"
                    Caption ="وصول به  حساب"
                    FontName ="Traffic"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =3820
                    Top =1835
                    Width =405
                    Height =405
                    FontSize =10
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label141"
                    Caption ="مبلغ"
                    FontName ="Traffic"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =4935
                    Top =1830
                    Width =1215
                    Height =405
                    FontSize =10
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label142"
                    Caption ="تاريخ سررسيد"
                    FontName ="Traffic"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =7221
                    Top =1835
                    Width =450
                    Height =405
                    FontSize =10
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label143"
                    Caption ="بانك"
                    FontName ="Traffic"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =9045
                    Top =1830
                    Width =570
                    Height =405
                    FontSize =10
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label144"
                    Caption ="سريال"
                    FontName ="Traffic"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =10026
                    Top =1826
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
                    Top =2295
                    Width =10626
                    BorderColor =0
                    Name ="Line147"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2402
                    Top =1155
                    Width =5046
                    Height =435
                    FontSize =12
                    TabIndex =1
                    Name ="MOLAH"
                    ControlSource ="MOLAH"
                    FontName ="Traffic"
                    Begin
                        Begin Label
                            TextAlign =0
                            TextFontFamily =10
                            Left =7469
                            Top =1155
                            Width =855
                            Height =435
                            FontSize =12
                            FontWeight =400
                            ForeColor =0
                            Name ="Label155"
                            Caption ="ملاحظات:"
                            FontName ="Traffic"
                        End
                    End
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1763
                    Top =566
                    Width =6753
                    Height =540
                    FontSize =14
                    TabIndex =2
                    BorderColor =32768
                    Name ="Text90"
                    ControlSource ="=[Forms]![Baseknow]![WIDTH_D]"
                    FontName ="Titr"
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =450
            Name ="Detail"
            Begin
                Begin TextBox
                    DecimalPlaces =0
                    RunningSum =1
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =10036
                    Top =16
                    Width =504
                    Height =405
                    FontSize =10
                    FontWeight =700
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
                    Left =8925
                    Top =16
                    Width =1044
                    Height =405
                    FontSize =10
                    FontWeight =700
                    TabIndex =1
                    BorderColor =32768
                    Name ="CODE"
                    ControlSource ="N_SERI"
                    FontName ="Traffic"
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4943
                    Top =16
                    Width =1224
                    Height =405
                    FontSize =10
                    FontWeight =700
                    TabIndex =2
                    BorderColor =32768
                    Name ="DATE_S"
                    ControlSource ="DATE_S"
                    FontName ="Traffic"
                    InputMask ="####/##/##"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3315
                    Top =16
                    Width =1569
                    Height =405
                    FontSize =10
                    FontWeight =700
                    TabIndex =3
                    BorderColor =32768
                    Name ="MEGHk"
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
                    Left =15
                    Top =16
                    Width =3174
                    Height =405
                    FontSize =10
                    FontWeight =700
                    TabIndex =4
                    BorderColor =32768
                    Name ="HES"
                    ControlSource ="HES"
                    Format ="#,###"
                    FontName ="Traffic"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =6213
                    Top =15
                    Width =2685
                    Height =405
                    FontSize =10
                    FontWeight =700
                    TabIndex =5
                    BorderColor =32768
                    Name ="BAN"
                    ControlSource ="BAN"
                    FontName ="Traffic"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Top =450
                    Width =10626
                    BorderColor =0
                    Name ="Line100"
                End
                Begin Line
                    BorderWidth =2
                    Left =10590
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line101"
                End
                Begin Line
                    BorderWidth =1
                    Left =9993
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line103"
                End
                Begin Line
                    BorderWidth =1
                    Left =8928
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line104"
                End
                Begin Line
                    BorderWidth =1
                    Left =6200
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line105"
                End
                Begin Line
                    BorderWidth =1
                    Left =4920
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line106"
                End
                Begin Line
                    BorderWidth =2
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line109"
                End
                Begin Line
                    BorderWidth =1
                    Left =3253
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line161"
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =30
            Name ="GroupFooter3"
            Begin
                Begin Line
                    BorderWidth =3
                    Top =30
                    Width =10630
                    BorderColor =0
                    Name ="Line110"
                End
            End
        End
    End
End
CodeBehindForm
' See "ELVSL_ADA.cls"
