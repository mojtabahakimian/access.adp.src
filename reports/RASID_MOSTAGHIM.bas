Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
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
    Width =10776
    DatasheetFontHeight =10
    ItemSuffix =175
    Left =930
    Top =1050
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x9e6b5e80fdd7e240
    End
    RecordSource ="QSL_RASID_MOSTAGHIM_LARGE"
    Caption ="برگ رسيد كالاي مصرف مستقيم"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial (Arabic)"
    OnError ="[Event Procedure]"
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
            GroupHeader = NotDefault
            GroupFooter = NotDefault
            ControlSource ="NUMBER"
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =2100
            Name ="GroupHeader2"
            Begin
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =3832
                    Width =3270
                    Height =540
                    FontSize =14
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label88"
                    Caption ="برگ رسيد كالاي مصرف مستقيم"
                    FontName ="Titr"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2430
                    Top =495
                    Width =5613
                    Height =540
                    FontSize =14
                    BorderColor =32768
                    Name ="Text90"
                    ControlSource ="=[Forms]![Baseknow]![WIDTH_D]"
                    FontName ="Titr"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =0
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =8050
                    Top =793
                    Height =300
                    FontSize =8
                    TabIndex =1
                    BorderColor =32768
                    Name ="DATE_N"
                    ControlSource ="DATE_N"
                    FontName ="Traffic"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =0
                            Left =9788
                            Top =793
                            Width =465
                            Height =330
                            FontSize =8
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
                    TextFontFamily =0
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =8100
                    Top =450
                    Width =1671
                    Height =300
                    FontSize =8
                    TabIndex =2
                    BorderColor =32768
                    Name ="NUMBER1"
                    ControlSource ="NUMBER"
                    FontName ="Traffic"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =0
                            Left =9765
                            Top =450
                            Width =900
                            Height =330
                            FontSize =8
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label93"
                            Caption ="شماره رسيد:"
                            FontName ="Traffic"
                        End
                    End
                End
                Begin Line
                    BorderWidth =3
                    Left =30
                    Top =1125
                    Width =10746
                    BorderColor =0
                    Name ="Line94"
                End
                Begin Line
                    BorderWidth =3
                    Top =1575
                    Width =10776
                    BorderColor =0
                    Name ="Line99"
                End
                Begin Line
                    BorderWidth =3
                    Left =10755
                    Top =1590
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line130"
                End
                Begin Line
                    Left =10170
                    Top =1590
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line131"
                End
                Begin Line
                    Left =8067
                    Top =1590
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line132"
                End
                Begin Line
                    Left =4329
                    Top =1590
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line133"
                End
                Begin Line
                    Left =3504
                    Top =1590
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line134"
                End
                Begin Line
                    Left =2400
                    Top =1590
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line135"
                End
                Begin Line
                    BorderWidth =3
                    Left =30
                    Top =1590
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line138"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =1792
                    Top =1635
                    Width =300
                    Height =405
                    FontSize =10
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label141"
                    Caption ="في"
                    FontName ="Traffic"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =3532
                    Top =1635
                    Width =810
                    Height =405
                    FontSize =10
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label142"
                    Caption ="مقدار كل"
                    FontName ="Traffic"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =2565
                    Top =1635
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
                    Left =4500
                    Top =1635
                    Width =2625
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
                    Left =8142
                    Top =1635
                    Width =1890
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
                    Left =10176
                    Top =1631
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
                    Top =2100
                    Width =10776
                    BorderColor =0
                    Name ="Line147"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =0
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2445
                    Top =1140
                    Width =1356
                    Height =375
                    FontSize =10
                    TabIndex =3
                    BorderColor =32768
                    Name ="NUMBER"
                    ControlSource ="NUMBER"
                    FontName ="Traffic"
                    Begin
                        Begin Label
                            TextFontFamily =0
                            Left =3862
                            Top =1147
                            Width =1605
                            Height =405
                            FontSize =10
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label161"
                            Caption ="شماره فاكتور خريد:"
                            FontName ="Traffic"
                        End
                    End
                End
                Begin Line
                    Left =1410
                    Top =1590
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line162"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =555
                    Top =1650
                    Width =405
                    Height =405
                    FontSize =10
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label163"
                    Caption ="مبلغ"
                    FontName ="Traffic"
                End
                Begin Label
                    TextFontFamily =0
                    Left =776
                    Top =1133
                    Width =1605
                    Height =405
                    FontSize =10
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label162"
                    Caption ="شماره درخواست:"
                    FontName ="Traffic"
                End
                Begin Line
                    Left =7500
                    Top =1590
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line166"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =7605
                    Top =1635
                    Width =315
                    Height =405
                    FontSize =10
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label167"
                    Caption ="كد"
                    FontName ="Traffic"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =10
                    Left =5490
                    Top =1155
                    Width =4341
                    Height =405
                    FontSize =10
                    FontWeight =700
                    TabIndex =4
                    Name ="CUST_HESAB.NAME"
                    ControlSource ="NAME"
                    FontName ="Traffic"
                    EventProcPrefix ="CUST_HESAB_NAME"
                    Begin
                        Begin Label
                            TextAlign =0
                            TextFontFamily =10
                            Left =9867
                            Top =1155
                            Width =855
                            Height =405
                            FontSize =10
                            BackColor =0
                            ForeColor =0
                            Name ="Label172"
                            Caption ="فروشنده:"
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
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =8099
                    Top =16
                    Width =2034
                    Height =405
                    FontSize =10
                    FontWeight =700
                    BorderColor =32768
                    Name ="INVO_LST.ANBAR"
                    ControlSource ="NAMES"
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
                    Left =10261
                    Top =16
                    Width =429
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
                    Left =4364
                    Top =16
                    Width =3099
                    Height =405
                    FontSize =10
                    FontWeight =700
                    TabIndex =2
                    BorderColor =32768
                    Name ="CODE"
                    ControlSource ="SNAME"
                    FontName ="Traffic"
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =2
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3514
                    Top =16
                    Width =789
                    Height =405
                    FontSize =10
                    FontWeight =700
                    TabIndex =3
                    BorderColor =32768
                    Name ="MEGH"
                    ControlSource ="MEGHk"
                    FontName ="Traffic"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =2
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1455
                    Top =16
                    Width =894
                    Height =405
                    FontSize =10
                    FontWeight =700
                    TabIndex =4
                    BorderColor =32768
                    Name ="MABL"
                    ControlSource ="MABL"
                    Format ="#,###"
                    FontName ="Traffic"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2415
                    Top =15
                    Width =1035
                    Height =405
                    FontSize =10
                    FontWeight =700
                    TabIndex =5
                    BorderColor =32768
                    Name ="VAHED_K"
                    ControlSource ="VNAMES"
                    FontName ="Traffic"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Top =450
                    Width =10776
                    BorderColor =0
                    Name ="Line100"
                End
                Begin Line
                    BorderWidth =3
                    Left =10755
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line101"
                End
                Begin Line
                    BorderWidth =1
                    Left =10170
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line102"
                End
                Begin Line
                    BorderWidth =1
                    Left =8067
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line103"
                End
                Begin Line
                    BorderWidth =1
                    Left =4314
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line104"
                End
                Begin Line
                    BorderWidth =1
                    Left =3489
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line105"
                End
                Begin Line
                    BorderWidth =1
                    Left =2400
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
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =45
                    Width =1314
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
                Begin Line
                    BorderWidth =1
                    Left =1411
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line165"
                End
                Begin TextBox
                    DecimalPlaces =0
                    RunningSum =1
                    TextAlign =2
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =7531
                    Top =16
                    Width =489
                    Height =405
                    FontSize =10
                    FontWeight =700
                    TabIndex =7
                    BorderColor =32768
                    Name ="Text169"
                    ControlSource ="CODE"
                    FontName ="Traffic"
                End
                Begin Line
                    BorderWidth =1
                    Left =7500
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line170"
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =1868
            Name ="GroupFooter3"
            Begin
                Begin Line
                    BorderWidth =3
                    Top =435
                    Width =10765
                    BorderColor =0
                    Name ="Line110"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =9471
                    Top =1414
                    Width =525
                    Height =405
                    FontSize =10
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label154"
                    Caption ="امضاء"
                    FontName ="Traffic"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =5173
                    Top =1409
                    Width =525
                    Height =405
                    FontSize =10
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label157"
                    Caption ="امضاء"
                    FontName ="Traffic"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3900
                    Top =465
                    Width =6021
                    Height =405
                    FontSize =10
                    FontWeight =700
                    BorderColor =32768
                    Name ="Text166"
                    ControlSource ="MOLAH"
                    FontName ="Traffic"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =9467
                    Top =960
                    Width =1170
                    Height =405
                    FontSize =10
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label158"
                    Caption ="تحويل دهنده:"
                    FontName ="Traffic"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =5159
                    Top =1016
                    Width =1215
                    Height =405
                    FontSize =10
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label160"
                    Caption ="تحويل گيرنده:"
                    FontName ="Traffic"
                End
                Begin Line
                    BorderWidth =2
                    Width =0
                    Height =419
                    BorderColor =0
                    Name ="Line205"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =9864
                    Width =435
                    Height =375
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="Label208"
                    Caption ="جمع :"
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
                    Left =1416
                    Width =8406
                    Height =375
                    FontSize =10
                    TabIndex =1
                    Name ="Text211"
                    ControlSource ="=ALPHANUM([jam]) & \" \" & \"ريال\""
                    Format ="Standard"
                    FontName ="Zar"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =1
                    KeyboardLanguage =15
                    ScrollBarAlign =1
                    NumeralShapes =2
                    Left =71
                    Width =1289
                    Height =375
                    FontSize =10
                    TabIndex =2
                    Name ="jam"
                    ControlSource ="=Format(Str(Sum([MABL_K])),\"#,###\") & \"  \""
                    Format ="#,###"
                    FontName ="Traffic"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Left =1411
                    Width =0
                    Height =419
                    BorderColor =0
                    Name ="Line213"
                End
                Begin Line
                    BorderWidth =3
                    Left =10755
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line164"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =9937
                    Top =465
                    Width =825
                    Height =405
                    FontSize =10
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label171"
                    Caption ="ملاحظات:"
                    FontName ="Traffic"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =1605
                    Top =1463
                    Width =525
                    Height =405
                    FontSize =10
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label173"
                    Caption ="امضاء"
                    FontName ="Traffic"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =1849
                    Top =1068
                    Width =705
                    Height =405
                    FontSize =10
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label174"
                    Caption ="انبار دار:"
                    FontName ="Traffic"
                End
            End
        End
    End
End
CodeBehindForm
' See "RASID_MOSTAGHIM.cls"
