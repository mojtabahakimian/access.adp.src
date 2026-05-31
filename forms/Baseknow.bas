Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    ControlBox = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    CloseButton = NotDefault
    DividingLines = NotDefault
    DefaultView =0
    AllowUpdating =4
    ScrollBars =2
    ViewsAllowed =1
    TabularCharSet =178
    TabularFamily =42
    BorderStyle =0
    Orientation =1
    PictureAlignment =5
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =12736
    DatasheetFontHeight =10
    ItemSuffix =458
    Left =8115
    Top =-375
    Right =20865
    Bottom =31290
    HelpContextId =1
    TimerInterval =500000
    DatasheetGridlinesColor =12632256
    MaxRecords =0
    PaintPalette = Begin
        0x000359000000000000ccff0005cdff000aceff000fcfff0014d0ff001ad1ff00 ,
        0x1fd2ff0024d3ff0029d4ff002ed5ff0034d6ff0039d7ff003ed8ff0043d9ff00 ,
        0x48daff004edbff0053dcff0058ddff005ddeff0062dfff0068e0ff006de1ff00 ,
        0x72e2ff0077e3ff007ce4ff0082e6ff0087e7ff008ce8ff0091e9ff0096eaff00 ,
        0x9cebff00a1ecff00a6edff00abeeff00b0efff00b6f0ff00bbf1ff00c0f2ff00 ,
        0xc5f3ff00caf4ff00cccccc00cfcfcf00d0f5ff00d1d1d100d4d4d400d5f6ff00 ,
        0xd6d6d600d9d9d900daf7ff00dbdbdb00dedede00dff8ff00e0e0e000e3e3e300 ,
        0xe4f9ff00e5e5e500e8e8e800eafaff00ebebeb00ededed00effbff00f0f0f000 ,
        0xf2f2f200f4fcff00f5f5f500f7f7f700f9fdff00fafafa00fcfcfc00ffffff00 ,
        0x800000000080000080800000000080008000800000808000c0c0c000c0dcc000 ,
        0xa6caf000fffbf000a0a0a40080808000ff00000000ff0000ffff00000000ff00 ,
        0xff00ff0000ffff0000000000
    End
    RecSrcDt = Begin
        0xd6e8245dd8cae240
    End
    RecordSource ="SAZMAN"
    Caption ="بسم الله الرحمن الرحيم"
    OnOpen ="[Event Procedure]"
    HelpFile ="amin98.hlp>main"
    DatasheetFontName ="Arial (Arabic)"
    OnTimer ="[Event Procedure]"
    OnError ="[Event Procedure]"
    RecordSourceQualifier ="dbo"
    Begin
        Begin Label
            BackStyle =0
            TextFontCharSet =178
            TextFontFamily =42
            FontName ="Arial (Arabic)"
        End
        Begin Rectangle
            BackStyle =0
            BorderWidth =1
            Width =850
            Height =850
        End
        Begin Line
            BorderWidth =1
            Width =1701
        End
        Begin Image
            SpecialEffect =1
            BackStyle =0
            OldBorderStyle =0
            PictureAlignment =2
            Width =1701
            Height =1701
        End
        Begin CommandButton
            TextFontCharSet =178
            TextFontFamily =42
            Width =1701
            Height =283
            FontSize =8
            FontWeight =400
            FontName ="Tahoma"
        End
        Begin OptionButton
            SpecialEffect =1
            LabelX =230
            LabelY =-30
        End
        Begin CheckBox
            SpecialEffect =1
            BorderWidth =3
            LabelX =230
            LabelY =-30
        End
        Begin OptionGroup
            Width =1701
            Height =1701
            BorderColor =12632256
        End
        Begin BoundObjectFrame
            SpecialEffect =1
            OldBorderStyle =0
            BackStyle =0
            Width =4536
            Height =2835
            LabelX =-1701
        End
        Begin TextBox
            TextFontCharSet =178
            TextFontFamily =42
            BackStyle =0
            Width =1701
            LabelX =-1701
            BorderColor =12632256
            FontName ="Arial (Arabic)"
        End
        Begin ListBox
            SpecialEffect =1
            TextFontCharSet =178
            TextFontFamily =42
            Width =1701
            Height =1417
            LabelX =-1701
            FontName ="Arial (Arabic)"
        End
        Begin ComboBox
            TextFontCharSet =178
            TextFontFamily =42
            BackStyle =0
            Width =1701
            LabelX =-1701
            BorderColor =12632256
            FontName ="Arial (Arabic)"
        End
        Begin Subform
            SpecialEffect =1
            Width =1701
            Height =1701
        End
        Begin UnboundObjectFrame
            SpecialEffect =1
            BackStyle =0
            Width =4536
            Height =2835
        End
        Begin ToggleButton
            TextFontCharSet =178
            TextFontFamily =42
            Width =283
            Height =283
            FontSize =8
            FontWeight =400
            FontName ="Arial (Arabic)"
        End
        Begin Tab
            TextFontCharSet =178
            TextFontFamily =42
            BackStyle =0
            Width =5103
            Height =3402
            FontName ="Tahoma"
        End
        Begin FormHeader
            Height =0
            Name ="FormHeader"
        End
        Begin Section
            SpecialEffect =1
            Height =31680
            Name ="Detail"
            Begin
                Begin BoundObjectFrame
                    SizeMode =1
                    OverlapFlags =85
                    Left =1418
                    Top =170
                    Width =456
                    Height =465
                    Name ="EMZA"
                    ControlSource ="EMZA"
                End
                Begin TextBox
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    KeyboardLanguage =15
                    Left =170
                    Width =1140
                    Height =285
                    FontSize =10
                    FontWeight =700
                    TabIndex =1
                    BackColor =16711680
                    BorderColor =3355443
                    ForeColor =16711680
                    Name ="UUSER"
                    FontName ="Tahoma"
                End
                Begin TextBox
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    KeyboardLanguage =15
                    Left =170
                    Top =405
                    Width =1140
                    Height =285
                    FontSize =10
                    FontWeight =700
                    TabIndex =2
                    BackColor =16711680
                    BorderColor =3355443
                    ForeColor =16711680
                    Name ="UGRP"
                    FontName ="Tahoma"
                End
                Begin TextBox
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    KeyboardLanguage =15
                    Left =170
                    Top =810
                    Width =1140
                    Height =285
                    FontSize =10
                    FontWeight =700
                    TabIndex =3
                    BackColor =16711680
                    BorderColor =3355443
                    ForeColor =16711680
                    Name ="USERCOD"
                    FontName ="Tahoma"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =9070
                    ColumnWidth =1155
                    TabIndex =4
                    Name ="UNIVERSITY_CO"
                    ControlSource ="UNIVERSITY_CO"
                    StatusBarText ="شماره شناسايي"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =9070
                    Top =340
                    ColumnWidth =1560
                    TabIndex =5
                    Name ="NAMEE"
                    ControlSource ="NAME"
                    StatusBarText ="نام  سازمان"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =9070
                    Top =680
                    ColumnWidth =1080
                    TabIndex =6
                    Name ="CITY"
                    ControlSource ="CITY"
                    StatusBarText ="نام شهر"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =9070
                    Top =1020
                    TabIndex =7
                    Name ="MANAGER"
                    ControlSource ="MANAGER"
                    StatusBarText ="نام رياست واحد"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =9070
                    Top =1360
                    TabIndex =8
                    Name ="MOAVEN"
                    ControlSource ="MOAVEN"
                    StatusBarText ="معاون اداري مالي"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =9070
                    Top =1701
                    TabIndex =9
                    Name ="ZIHESAB"
                    ControlSource ="ZIHESAB"
                    StatusBarText ="نام و نام خانوادگي مدير مالي"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =9070
                    Top =2041
                    TabIndex =10
                    Name ="AMINAMVAL"
                    ControlSource ="AMINAMVAL"
                    StatusBarText ="نام و نام خانوادگي  امين اموال"
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =9070
                    Top =2381
                    TabIndex =11
                    Name ="YEA"
                    ControlSource ="YEA"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =9070
                    Top =2721
                    TabIndex =12
                    Name ="SANAD"
                    ControlSource ="SANAD"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =9070
                    Top =3061
                    TabIndex =13
                    Name ="GHAYM"
                    ControlSource ="GHAYM"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =9070
                    Top =3401
                    TabIndex =14
                    Name ="KALA"
                    ControlSource ="KALA"
                End
                Begin TextBox
                    OverlapFlags =85
                    BackStyle =1
                    IMESentenceMode =3
                    Left =9070
                    Top =3741
                    TabIndex =15
                    BackColor =52479
                    ForeColor =255
                    Name ="PERSON"
                    ControlSource ="PERSON"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =9070
                    Top =4082
                    TabIndex =16
                    Name ="DIG"
                    ControlSource ="DIG"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =9070
                    Top =4422
                    TabIndex =17
                    Name ="WAR"
                    ControlSource ="WAR"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =9070
                    Top =4762
                    TabIndex =18
                    Name ="LST"
                    ControlSource ="LST"
                End
                Begin TextBox
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =9070
                    Top =5102
                    TabIndex =19
                    Name ="TFTPAGE"
                    ControlSource ="TFTPAGE"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =9070
                    Top =5442
                    TabIndex =20
                    Name ="TFSAZMAN"
                    ControlSource ="TFSAZMAN"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =9070
                    Top =5782
                    TabIndex =21
                    Name ="TFADDRESS"
                    ControlSource ="TFADDRESS"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =9070
                    Top =6123
                    TabIndex =22
                    Name ="TFTEL"
                    ControlSource ="TFTEL"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =9070
                    Top =6463
                    TabIndex =23
                    Name ="TFCODE_E"
                    ControlSource ="TFCODE_E"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =9070
                    Top =6803
                    TabIndex =24
                    Name ="WIDTH_D"
                    ControlSource ="WIDTH_D"
                End
                Begin TextBox
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =9070
                    Top =7143
                    TabIndex =25
                    Name ="HIGH_D"
                    ControlSource ="HIGH_D"
                End
                Begin TextBox
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =9070
                    Top =7483
                    TabIndex =26
                    Name ="CPI"
                    ControlSource ="CPI"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =9070
                    Top =7823
                    TabIndex =27
                    Name ="SANDOGH"
                    ControlSource ="SANDOGH"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =9070
                    Top =8164
                    TabIndex =28
                    Name ="BANKHA"
                    ControlSource ="BANKHA"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =9070
                    Top =8504
                    TabIndex =29
                    Name ="BESTANKAR"
                    ControlSource ="BESTANKAR"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =9070
                    Top =8844
                    TabIndex =30
                    Name ="BEDEHKAR"
                    ControlSource ="BEDEHKAR"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =9070
                    Top =9184
                    TabIndex =31
                    Name ="KHARID"
                    ControlSource ="KHARID"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =9070
                    Top =9524
                    TabIndex =32
                    Name ="MKHARID"
                    ControlSource ="MKHARID"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =9070
                    Top =9864
                    TabIndex =33
                    Name ="TKHARID"
                    ControlSource ="TKHARID"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =9070
                    Top =10204
                    TabIndex =34
                    Name ="HKHARID"
                    ControlSource ="HKHARID"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =9070
                    Top =10545
                    TabIndex =35
                    Name ="FROSH"
                    ControlSource ="FROSH"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =9070
                    Top =10885
                    TabIndex =36
                    Name ="MFROSH"
                    ControlSource ="MFROSH"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =9070
                    Top =11225
                    TabIndex =37
                    Name ="TFROSH"
                    ControlSource ="TFROSH"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =9070
                    Top =11565
                    TabIndex =38
                    Name ="HFROSH"
                    ControlSource ="HFROSH"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =9070
                    Top =11905
                    TabIndex =39
                    Name ="MOGODIA"
                    ControlSource ="MOGODIA"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =9070
                    Top =12245
                    TabIndex =40
                    Name ="MOGODIP"
                    ControlSource ="MOGODIP"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =9070
                    Top =12586
                    TabIndex =41
                    Name ="DARAM"
                    ControlSource ="DARAM"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =9070
                    Top =12926
                    TabIndex =42
                    Name ="HDARAM"
                    ControlSource ="HDARAM"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =9070
                    Top =13266
                    TabIndex =43
                    Name ="HKOL"
                    ControlSource ="HKOL"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =9070
                    Top =13606
                    TabIndex =44
                    Name ="ADA"
                    ControlSource ="ADA"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =9070
                    Top =13946
                    TabIndex =45
                    Name ="APA"
                    ControlSource ="APA"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =9070
                    Top =14286
                    TabIndex =46
                    Name ="ADV"
                    ControlSource ="ADV"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =9070
                    Top =14627
                    TabIndex =47
                    Name ="HAVALAH"
                    ControlSource ="HAVALAH"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =9070
                    Top =14967
                    TabIndex =48
                    Name ="CTRL_TS"
                    ControlSource ="CTRL_TS"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =9070
                    Top =15307
                    TabIndex =49
                    Name ="F_ANBARF"
                    ControlSource ="F_ANBARF"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =9070
                    Top =15647
                    TabIndex =50
                    Name ="GH_PK"
                    ControlSource ="GH_PK"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =9070
                    Top =15987
                    TabIndex =51
                    Name ="L_NUMBER"
                    ControlSource ="L_NUMBER"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =9070
                    Top =16327
                    TabIndex =52
                    Name ="SF_G"
                    ControlSource ="SF_G"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =9070
                    Top =16667
                    TabIndex =53
                    Name ="TAR_KM"
                    ControlSource ="TAR_KM"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =9070
                    Top =17008
                    TabIndex =54
                    Name ="BACKPATH"
                    ControlSource ="BACKPATH"
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =9070
                    Top =17348
                    TabIndex =55
                    Name ="TKHF"
                    ControlSource ="TKHF"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =9070
                    Top =17688
                    TabIndex =56
                    Name ="HAZ_TOL"
                    ControlSource ="HAZ_TOL"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =9070
                    Top =18028
                    TabIndex =57
                    Name ="PJHAZ_TOL1"
                    ControlSource ="PJHAZ_TOL1"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =9070
                    Top =18368
                    TabIndex =58
                    Name ="PHAZ_TOL"
                    ControlSource ="PHAZ_TOL"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =9070
                    Top =18708
                    TabIndex =59
                    Name ="GHEYMAT"
                    ControlSource ="GHEYMAT"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =9070
                    Top =19049
                    TabIndex =60
                    Name ="PPDAST"
                    ControlSource ="PPDAST"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =9070
                    Top =19389
                    TabIndex =61
                    Name ="PPSAR"
                    ControlSource ="PPSAR"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =9070
                    Top =19729
                    TabIndex =62
                    Name ="AMALKARD"
                    ControlSource ="AMALKARD"
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =9070
                    Top =20069
                    TabIndex =63
                    ForeColor =255
                    Name ="PERSONEL"
                    ControlSource ="PERSONEL"
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =9070
                    Top =20409
                    TabIndex =64
                    Name ="PERVAM"
                    ControlSource ="PERVAM"
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =9070
                    Top =20749
                    TabIndex =65
                    Name ="CONKAL"
                    ControlSource ="CONKAL"
                End
                Begin CheckBox
                    OverlapFlags =85
                    ReadingOrder =1
                    Left =10681
                    Top =21233
                    TabIndex =66
                    Name ="HNAH"
                    ControlSource ="HNAH"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =10186
                            Top =21203
                            Width =525
                            Height =255
                            Name ="Label253"
                            Caption ="حق نهار"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    ReadingOrder =1
                    Left =10681
                    Top =21573
                    TabIndex =67
                    Name ="HEZA"
                    ControlSource ="HEZA"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =10096
                            Top =21543
                            Width =615
                            Height =255
                            Name ="Label254"
                            Caption ="اضافه كار"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    ReadingOrder =1
                    Left =10681
                    Top =21913
                    TabIndex =68
                    Name ="HPAD"
                    ControlSource ="HPAD"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =10321
                            Top =21883
                            Width =390
                            Height =255
                            Name ="Label255"
                            Caption ="پاداش"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    ReadingOrder =1
                    Left =10681
                    Top =22253
                    TabIndex =69
                    Name ="HOLA"
                    ControlSource ="HOLA"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =10156
                            Top =22223
                            Width =555
                            Height =255
                            Name ="Label256"
                            Caption ="حق اولاد"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    ReadingOrder =1
                    Left =10681
                    Top =22593
                    TabIndex =70
                    Name ="HKHA"
                    ControlSource ="HKHA"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =10156
                            Top =22563
                            Width =555
                            Height =255
                            Name ="Label257"
                            Caption ="خوار بار"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    ReadingOrder =1
                    Left =10681
                    Top =22934
                    TabIndex =71
                    Name ="HJAZ"
                    ControlSource ="HJAZ"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =10141
                            Top =22904
                            Width =570
                            Height =255
                            Name ="Label258"
                            Caption ="حق جذب"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    ReadingOrder =1
                    Left =10681
                    Top =23274
                    TabIndex =72
                    Name ="HRAN"
                    ControlSource ="HRAN"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =10231
                            Top =23244
                            Width =480
                            Height =255
                            Name ="Label259"
                            Caption ="راندمان"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    ReadingOrder =1
                    Left =10681
                    Top =23614
                    TabIndex =73
                    Name ="HSAY"
                    ControlSource ="HSAY"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =10381
                            Top =23584
                            Width =330
                            Height =255
                            Name ="Label260"
                            Caption ="ساير"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    ReadingOrder =1
                    Left =10681
                    Top =23954
                    TabIndex =74
                    Name ="HCON"
                    ControlSource ="HCON"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =10306
                            Top =23924
                            Width =405
                            Height =255
                            Name ="Label261"
                            Caption ="شرايط"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    ReadingOrder =1
                    Left =10681
                    Top =24294
                    TabIndex =75
                    Name ="HSHI"
                    ControlSource ="HSHI"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =10291
                            Top =24264
                            Width =420
                            Height =255
                            Name ="Label262"
                            Caption ="HSHI"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =9240
                    Top =24634
                    TabIndex =76
                    Name ="HAZEDAR"
                    ControlSource ="HAZEDAR"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =9240
                    Top =24975
                    TabIndex =77
                    Name ="EDABIM"
                    ControlSource ="EDABIM"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =9240
                    Top =25315
                    TabIndex =78
                    Name ="HAZBIM"
                    ControlSource ="HAZBIM"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =9240
                    Top =25655
                    TabIndex =79
                    Name ="BESHO"
                    ControlSource ="BESHO"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =9240
                    Top =25995
                    TabIndex =80
                    Name ="BEDMOS"
                    ControlSource ="BEDMOS"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4251
                    TabIndex =81
                    Name ="PARDAKH"
                    ControlSource ="PARDAKH"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4251
                    Top =340
                    TabIndex =82
                    Name ="HAZMALI"
                    ControlSource ="HAZMALI"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4251
                    Top =681
                    TabIndex =83
                    Name ="SAGHFH"
                    ControlSource ="SAGHFH"
                End
                Begin CheckBox
                    OverlapFlags =85
                    ReadingOrder =1
                    Left =5692
                    Top =1021
                    TabIndex =84
                    Name ="MAND"
                    ControlSource ="MAND"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =4627
                            Top =991
                            Width =1095
                            Height =255
                            Name ="Label349"
                            Caption ="نمايش مانده حساب"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    ReadingOrder =1
                    Left =5692
                    Top =1361
                    TabIndex =85
                    Name ="MOJU"
                    ControlSource ="MOJU"
                    Begin
                        Begin Label
                            OverlapFlags =255
                            Left =4822
                            Top =1331
                            Width =900
                            Height =255
                            Name ="Label350"
                            Caption ="كنترل موجودي"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    ReadingOrder =1
                    Left =5692
                    Top =1701
                    TabIndex =86
                    Name ="SA_HOGH"
                    ControlSource ="SA_HOGH"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =5347
                            Top =1671
                            Width =375
                            Height =255
                            Name ="Label351"
                            Caption ="حقوق"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =93
                    ReadingOrder =1
                    Left =5692
                    Top =2041
                    TabIndex =87
                    Name ="SA_40EZ"
                    ControlSource ="SA_40EZ"
                    Begin
                        Begin Label
                            OverlapFlags =255
                            Left =4777
                            Top =2011
                            Width =945
                            Height =255
                            Name ="Label352"
                            Caption ="40%اضافه كار"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =93
                    ReadingOrder =1
                    Left =5692
                    Top =2381
                    TabIndex =88
                    Name ="SA_EZAF"
                    ControlSource ="SA_EZAF"
                    Begin
                        Begin Label
                            OverlapFlags =255
                            Left =5107
                            Top =2351
                            Width =615
                            Height =255
                            Name ="Label353"
                            Caption ="اضافه كار"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    ReadingOrder =1
                    Left =5692
                    Top =2721
                    TabIndex =89
                    Name ="SA_PADA"
                    ControlSource ="SA_PADA"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =5332
                            Top =2691
                            Width =390
                            Height =255
                            Name ="Label354"
                            Caption ="پاداش"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =93
                    ReadingOrder =1
                    Left =5692
                    Top =3062
                    TabIndex =90
                    Name ="SA_HOLA"
                    ControlSource ="SA_HOLA"
                    Begin
                        Begin Label
                            OverlapFlags =255
                            Left =5167
                            Top =3032
                            Width =555
                            Height =255
                            Name ="Label355"
                            Caption ="حق اولاد"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    ReadingOrder =1
                    Left =5692
                    Top =3402
                    TabIndex =91
                    Name ="SA_KHAR"
                    ControlSource ="SA_KHAR"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =4747
                            Top =3372
                            Width =975
                            Height =255
                            Name ="Label356"
                            Caption ="خواربار و مسكن"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    ReadingOrder =1
                    Left =5692
                    Top =3742
                    TabIndex =92
                    Name ="SA_NAHA"
                    ControlSource ="SA_NAHA"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =5197
                            Top =3712
                            Width =525
                            Height =255
                            Name ="Label357"
                            Caption ="حق نهار"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    ReadingOrder =1
                    Left =5692
                    Top =4082
                    TabIndex =93
                    Name ="SA_JAZB"
                    ControlSource ="SA_JAZB"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =5152
                            Top =4052
                            Width =570
                            Height =255
                            Name ="Label358"
                            Caption ="حق جذب"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    ReadingOrder =1
                    Left =5692
                    Top =4422
                    TabIndex =94
                    Name ="SA_RAND"
                    ControlSource ="SA_RAND"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =5242
                            Top =4392
                            Width =480
                            Height =255
                            Name ="Label359"
                            Caption ="راندمان"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    ReadingOrder =1
                    Left =5692
                    Top =4762
                    TabIndex =95
                    Name ="SA_COND"
                    ControlSource ="SA_COND"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =4987
                            Top =4732
                            Width =735
                            Height =255
                            Name ="Label360"
                            Caption ="شرايط محيط"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    ReadingOrder =1
                    Left =5692
                    Top =5103
                    TabIndex =96
                    Name ="SA_SAYE"
                    ControlSource ="SA_SAYE"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =5392
                            Top =5073
                            Width =330
                            Height =255
                            Name ="Label361"
                            Caption ="ساير"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    ReadingOrder =1
                    Left =5692
                    Top =5443
                    TabIndex =97
                    Name ="SA_23BI"
                    ControlSource ="SA_23BI"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =5107
                            Top =5413
                            Width =615
                            Height =255
                            Name ="Label362"
                            Caption ="23%بيمه"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4251
                    Top =5783
                    TabIndex =98
                    Name ="HAZTOLID"
                    ControlSource ="HAZTOLID"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6768
                            Top =5783
                            Width =885
                            Height =255
                            Name ="Label363"
                            Caption ="HAZTOLID:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4251
                    Top =6123
                    TabIndex =99
                    Name ="HAZFROOSH"
                    ControlSource ="HAZFROOSH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6558
                            Top =6123
                            Width =1095
                            Height =255
                            Name ="Label364"
                            Caption ="HAZFROOSH:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4251
                    Top =6463
                    TabIndex =100
                    Name ="HAZKHADAMAT"
                    ControlSource ="HAZKHADAMAT"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6318
                            Top =6463
                            Width =1335
                            Height =255
                            Name ="Label365"
                            Caption ="HAZKHADAMAT:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4251
                    Top =6803
                    TabIndex =101
                    Name ="PISHDAR"
                    ControlSource ="PISHDAR"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6873
                            Top =6803
                            Width =780
                            Height =255
                            Name ="Label366"
                            Caption ="PISHDAR:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4251
                    Top =7144
                    TabIndex =102
                    Name ="DEFANB"
                    ControlSource ="DEFANB"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =6918
                            Top =7144
                            Width =735
                            Height =255
                            Name ="Label367"
                            Caption ="DEFANB:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4251
                    Top =7484
                    TabIndex =103
                    Name ="DEFTKH"
                    ControlSource ="DEFTKH"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =6948
                            Top =7484
                            Width =705
                            Height =255
                            Name ="Label368"
                            Caption ="DEFTKH:"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    ReadingOrder =1
                    Left =5692
                    Top =7824
                    TabIndex =104
                    Name ="ECONM"
                    ControlSource ="ECONM"
                    StatusBarText ="مواد اوليه در زمان توليد خارج  شود"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =5107
                            Top =7794
                            Width =615
                            Height =255
                            Name ="Label369"
                            Caption ="ECONM"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4251
                    Top =8164
                    TabIndex =105
                    Name ="FRUP"
                    ControlSource ="FRUP"
                    StatusBarText ="ستون مبلغ  قابل تغيير باشد"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4251
                    Top =8504
                    TabIndex =106
                    Name ="UPDDATE"
                    ControlSource ="UPDDATE"
                    StatusBarText ="تاريخ فاكتور"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4251
                    Top =8844
                    TabIndex =107
                    Name ="FINALS"
                    ControlSource ="FINALS"
                    StatusBarText ="نحوه خروج مواداز انبار در روش اجراي حسابداري صنعتي"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4251
                    Top =9184
                    TabIndex =108
                    Name ="PSANDHES"
                    ControlSource ="PSANDHES"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4251
                    Top =9525
                    TabIndex =109
                    Name ="SANAVP"
                    ControlSource ="SANAVP"
                    StatusBarText ="حق سنوات پايه:"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4251
                    Top =9865
                    TabIndex =110
                    Name ="BON"
                    ControlSource ="BON"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4251
                    Top =10205
                    TabIndex =111
                    Name ="ISO_FROOSH"
                    ControlSource ="ISO_FROOSH"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4251
                    Top =10545
                    TabIndex =112
                    Name ="ISO_KHAREED"
                    ControlSource ="ISO_KHAREED"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4251
                    Top =10885
                    TabIndex =113
                    Name ="ISO_MAVAD"
                    ControlSource ="ISO_MAVAD"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4251
                    Top =11225
                    TabIndex =114
                    Name ="ISO_TOLID"
                    ControlSource ="ISO_TOLID"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4251
                    Top =11566
                    TabIndex =115
                    Name ="ISO_MAVADSAYER"
                    ControlSource ="ISO_MAVADSAYER"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4251
                    Top =11906
                    TabIndex =116
                    Name ="SANAT"
                    ControlSource ="SANAT"
                    StatusBarText ="ارتيكل اسناد صنعتي صادر شود"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4251
                    Top =12246
                    TabIndex =117
                    Name ="CODEVIEW"
                    ControlSource ="CODEVIEW"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4251
                    Top =12586
                    TabIndex =118
                    Name ="PKHARID"
                    ControlSource ="PKHARID"
                    StatusBarText ="پاياپاي خريد"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4251
                    Top =12926
                    TabIndex =119
                    Name ="SIGN"
                    ControlSource ="SIGN"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4251
                    Top =13266
                    TabIndex =120
                    Name ="BARCOD"
                    ControlSource ="BARCOD"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4251
                    Top =13607
                    TabIndex =121
                    Name ="SAGHF"
                    ControlSource ="SAGHF"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4251
                    Top =13947
                    TabIndex =122
                    Name ="SERVERNAM"
                    ControlSource ="SERVERNAM"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4251
                    Top =14287
                    TabIndex =123
                    Name ="TENDAR"
                    ControlSource ="TENDAR"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4251
                    Top =14627
                    TabIndex =124
                    Name ="LECOL1"
                    ControlSource ="LECOL1"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4251
                    Top =14967
                    TabIndex =125
                    Name ="LECOL2"
                    ControlSource ="LECOL2"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4251
                    Top =15307
                    TabIndex =126
                    Name ="LECOL3"
                    ControlSource ="LECOL3"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4251
                    Top =15647
                    TabIndex =127
                    Name ="LECOL4"
                    ControlSource ="LECOL4"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4251
                    Top =15988
                    TabIndex =128
                    Name ="LKCOL1"
                    ControlSource ="LKCOL1"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4251
                    Top =16328
                    TabIndex =129
                    Name ="HESMBAA"
                    ControlSource ="HESMBAA"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4251
                    Top =16668
                    TabIndex =130
                    Name ="ECODE"
                    ControlSource ="ECODE"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4251
                    Top =17008
                    TabIndex =131
                    Name ="PCODE"
                    ControlSource ="PCODE"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4251
                    Top =17348
                    TabIndex =132
                    Name ="IYALAT"
                    ControlSource ="IYALAT"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4251
                    Top =17688
                    TabIndex =133
                    Name ="MCODEM"
                    ControlSource ="MCODEM"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4251
                    Top =18029
                    TabIndex =134
                    Name ="HPOR"
                    ControlSource ="HPOR"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4251
                    Top =18369
                    TabIndex =135
                    Name ="SAGHF2"
                    ControlSource ="SAGHF2"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4251
                    Top =18709
                    TabIndex =136
                    Name ="OPTIONSS"
                    ControlSource ="OPTIONSS"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4251
                    Top =19049
                    TabIndex =137
                    Name ="CTL_DT"
                    ControlSource ="CTL_DT"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4251
                    Top =19389
                    TabIndex =138
                    Name ="LOCKFAP"
                    ControlSource ="LOCKFAP"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4251
                    Top =19729
                    TabIndex =139
                    Name ="LOCKFSI"
                    ControlSource ="LOCKFSI"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4251
                    Top =20070
                    TabIndex =140
                    Name ="TRANSF"
                    ControlSource ="TRANSF"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4251
                    Top =20410
                    TabIndex =141
                    Name ="OKF"
                    ControlSource ="OKF"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4251
                    Top =20750
                    TabIndex =142
                    Name ="ARSESH"
                    ControlSource ="ARSESH"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4251
                    Top =21090
                    TabIndex =143
                    Name ="RMOG"
                    ControlSource ="RMOG"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4251
                    Top =21430
                    TabIndex =144
                    Name ="APV"
                    ControlSource ="APV"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4251
                    Top =21770
                    TabIndex =145
                    Name ="HOTCOD"
                    ControlSource ="HOTCOD"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4251
                    Top =22110
                    TabIndex =146
                    Name ="STFR"
                    ControlSource ="STFR"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4251
                    Top =22451
                    TabIndex =147
                    Name ="STKH"
                    ControlSource ="STKH"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4251
                    Top =22791
                    TabIndex =148
                    Name ="STHFR"
                    ControlSource ="STHFR"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4251
                    Top =23131
                    TabIndex =149
                    Name ="STHKH"
                    ControlSource ="STHKH"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4251
                    Top =23471
                    TabIndex =150
                    Name ="STENT"
                    ControlSource ="STENT"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4251
                    Top =23811
                    TabIndex =151
                    Name ="STKHS"
                    ControlSource ="STKHS"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4251
                    Top =24151
                    TabIndex =152
                    Name ="STKHH"
                    ControlSource ="STKHH"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4251
                    Top =24492
                    TabIndex =153
                    Name ="STTOL"
                    ControlSource ="STTOL"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4251
                    Top =24832
                    TabIndex =154
                    Name ="STFRB"
                    ControlSource ="STFRB"
                End
                Begin TextBox
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =4251
                    Top =25172
                    TabIndex =155
                    Name ="STBKH"
                    ControlSource ="STBKH"
                End
                Begin TextBox
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =4251
                    Top =25512
                    TabIndex =156
                    Name ="STMO"
                    ControlSource ="STMO"
                End
                Begin TextBox
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =4251
                    Top =25852
                    TabIndex =157
                    Name ="STKHA"
                    ControlSource ="STKHA"
                End
                Begin TextBox
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =4251
                    Top =26192
                    TabIndex =158
                    Name ="SNDKH"
                    ControlSource ="SNDKH"
                End
                Begin TextBox
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    KeyboardLanguage =15
                    Left =170
                    Top =1215
                    Width =1140
                    Height =285
                    FontSize =10
                    FontWeight =700
                    TabIndex =159
                    BackColor =16711680
                    BorderColor =3355443
                    ForeColor =16711680
                    Name ="Text44"
                    FontName ="Tahoma"
                End
                Begin TextBox
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    KeyboardLanguage =15
                    Left =226
                    Top =1587
                    Width =1140
                    Height =285
                    FontSize =10
                    FontWeight =700
                    TabIndex =160
                    BackColor =16711680
                    BorderColor =3355443
                    ForeColor =16711680
                    Name ="dttmp"
                    DefaultValue ="=farsidate(Date())"
                    FontName ="Tahoma"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =283
                    Top =3344
                    TabIndex =161
                    Name ="SMS_USERNAME"
                    ControlSource ="SMS_USERNAME"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2305
                            Top =3344
                            Width =1380
                            Height =255
                            Name ="Label425"
                            Caption ="SMS_USERNAME:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =283
                    Top =3685
                    TabIndex =162
                    Name ="SMS_PASSWORD"
                    ControlSource ="SMS_PASSWORD"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2245
                            Top =3685
                            Width =1440
                            Height =255
                            Name ="Label426"
                            Caption ="SMS_PASSWORD:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =283
                    Top =4025
                    TabIndex =163
                    Name ="SMS_LIBKEY"
                    ControlSource ="SMS_LIBKEY"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2605
                            Top =4025
                            Width =1080
                            Height =255
                            Name ="Label427"
                            Caption ="SMS_LIBKEY:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =283
                    Top =4365
                    TabIndex =164
                    Name ="SMS_TSMSHOST"
                    ControlSource ="SMS_TSMSHOST"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2305
                            Top =4365
                            Width =1380
                            Height =255
                            Name ="Label428"
                            Caption ="SMS_TSMSHOST:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =87
                    IMESentenceMode =3
                    Left =283
                    Top =4705
                    TabIndex =165
                    Name ="SMS_ProxyUserName"
                    ControlSource ="SMS_ProxyUserName"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =1975
                            Top =4705
                            Width =1710
                            Height =255
                            Name ="Label429"
                            Caption ="SMS_ProxyUserName:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =87
                    IMESentenceMode =3
                    Left =283
                    Top =5045
                    TabIndex =166
                    Name ="SMS_ProxyPassword"
                    ControlSource ="SMS_ProxyPassword"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =1975
                            Top =5045
                            Width =1710
                            Height =255
                            Name ="Label430"
                            Caption ="SMS_ProxyPassword:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =283
                    Top =5385
                    TabIndex =167
                    Name ="SMS_ProxyServer"
                    ControlSource ="SMS_ProxyServer"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2230
                            Top =5385
                            Width =1455
                            Height =255
                            Name ="Label431"
                            Caption ="SMS_ProxyServer:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =283
                    Top =5725
                    TabIndex =168
                    Name ="SMS_ProxyPort"
                    ControlSource ="SMS_ProxyPort"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2440
                            Top =5725
                            Width =1245
                            Height =255
                            Name ="Label432"
                            Caption ="SMS_ProxyPort:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =283
                    Top =6066
                    TabIndex =169
                    Name ="SMS_FirewallUserName"
                    ControlSource ="SMS_FirewallUserName"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =1825
                            Top =6066
                            Width =1860
                            Height =255
                            Name ="Label433"
                            Caption ="SMS_FirewallUserName:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =283
                    Top =6406
                    TabIndex =170
                    Name ="SMS_FirewallPassword"
                    ControlSource ="SMS_FirewallPassword"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =1825
                            Top =6406
                            Width =1860
                            Height =255
                            Name ="Label434"
                            Caption ="SMS_FirewallPassword:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =283
                    Top =6746
                    TabIndex =171
                    Name ="SMS_FirewallHost"
                    ControlSource ="SMS_FirewallHost"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2245
                            Top =6746
                            Width =1440
                            Height =255
                            Name ="Label435"
                            Caption ="SMS_FirewallHost:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =283
                    Top =7086
                    TabIndex =172
                    Name ="SMS_FirewallPort"
                    ControlSource ="SMS_FirewallPort"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2290
                            Top =7086
                            Width =1395
                            Height =255
                            Name ="Label436"
                            Caption ="SMS_FirewallPort:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =283
                    Top =7426
                    TabIndex =173
                    Name ="SMS_FirewallType"
                    ControlSource ="SMS_FirewallType"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2215
                            Top =7426
                            Width =1470
                            Height =255
                            Name ="Label437"
                            Caption ="SMS_FirewallType:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =283
                    Top =7766
                    TabIndex =174
                    Name ="DSMS"
                    ControlSource ="DSMS"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3130
                            Top =7766
                            Width =555
                            Height =255
                            Name ="Label438"
                            Caption ="DSMS:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =283
                    Top =8107
                    TabIndex =175
                    Name ="SMS_OWNER"
                    ControlSource ="SMS_OWNER"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2575
                            Top =8107
                            Width =1110
                            Height =255
                            Name ="Label439"
                            Caption ="SMS_OWNER:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =623
                    Top =2437
                    TabIndex =176
                    Name ="PRMFR"
                    ControlSource ="PRMFR"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3395
                            Top =2437
                            Width =630
                            Height =255
                            Name ="Label440"
                            Caption ="PRMFR:"
                        End
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    KeyboardLanguage =15
                    Left =226
                    Top =1992
                    Width =1140
                    Height =285
                    FontSize =10
                    FontWeight =700
                    TabIndex =177
                    BackColor =16711680
                    BorderColor =3355443
                    ForeColor =16711680
                    Name ="DTT"
                    DefaultValue ="=farsidate(Date())"
                    FontName ="Tahoma"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1634
                    Top =11111
                    TabIndex =178
                    Name ="ISO_DTOLID"
                    ControlSource ="ISO_DTOLID"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1473
                    Top =1530
                    TabIndex =179
                    Name ="HBON"
                    ControlSource ="HBON"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =4320
                            Top =1530
                            Width =555
                            Height =255
                            Name ="Label443"
                            Caption ="HBON:"
                        End
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    BorderWidth =1
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =15
                    BackStyle =1
                    NumeralShapes =2
                    Left =1984
                    Top =453
                    Width =2115
                    Height =465
                    FontSize =10
                    FontWeight =700
                    TabIndex =180
                    BorderColor =3355443
                    Name ="tindata"
                    FontName ="Tahoma"
                End
                Begin TextBox
                    TabStop = NotDefault
                    BorderWidth =1
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    BackStyle =1
                    NumeralShapes =2
                    Left =2494
                    Top =56
                    Width =1140
                    Height =285
                    FontSize =10
                    FontWeight =700
                    TabIndex =181
                    BorderColor =3355443
                    ForeColor =16711680
                    Name ="dt"
                    DefaultValue ="=farsidate2(Date())"
                    FontName ="Tahoma"
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    KeyboardLanguage =15
                    Left =6462
                    Top =566
                    Width =1140
                    Height =285
                    FontSize =10
                    FontWeight =700
                    TabIndex =182
                    BackColor =16711680
                    BorderColor =3355443
                    ForeColor =16711680
                    Name ="MATN"
                    FontName ="Tahoma"
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1303
                    Top =2891
                    ColumnWidth =0
                    TabIndex =183
                    Name ="version"
                    ControlSource ="version"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4045
                            Top =2891
                            Width =660
                            Height =255
                            Name ="Label445"
                            Caption ="version:"
                        End
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    BorderWidth =1
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =11
                    BackStyle =1
                    NumeralShapes =2
                    Left =1984
                    Top =1038
                    Width =2115
                    Height =465
                    FontSize =10
                    FontWeight =700
                    TabIndex =184
                    BorderColor =3355443
                    Name ="hhwin"
                    FontName ="Tahoma"
                End
                Begin TextBox
                    OverlapFlags =247
                    IMESentenceMode =3
                    Left =5045
                    Top =3118
                    TabIndex =185
                    Name ="hesnaghd"
                    ControlSource ="hesnaghd"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =7607
                            Top =3118
                            Width =840
                            Height =255
                            Name ="Label447"
                            Caption ="hesnaghd:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =247
                    IMESentenceMode =3
                    Left =5551
                    Top =25390
                    TabIndex =186
                    Name ="IT1"
                    ControlSource ="IT1"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =8623
                            Top =25390
                            Width =330
                            Height =255
                            Name ="Label97"
                            Caption ="IT1:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =247
                    IMESentenceMode =3
                    Left =5551
                    Top =25730
                    TabIndex =187
                    Name ="IT2"
                    ControlSource ="IT2"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =8623
                            Top =25730
                            Width =330
                            Height =255
                            Name ="Label98"
                            Caption ="IT2:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =247
                    IMESentenceMode =3
                    Left =5551
                    Top =26070
                    TabIndex =188
                    Name ="IT3"
                    ControlSource ="IT3"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =8623
                            Top =26070
                            Width =330
                            Height =255
                            Name ="Label99"
                            Caption ="IT3:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =247
                    IMESentenceMode =3
                    Left =5551
                    Top =26410
                    TabIndex =189
                    Name ="IT4"
                    ControlSource ="IT4"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =8623
                            Top =26410
                            Width =330
                            Height =255
                            Name ="Label100"
                            Caption ="IT4:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =5551
                    Top =26750
                    TabIndex =190
                    Name ="IT5"
                    ControlSource ="IT5"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =8623
                            Top =26750
                            Width =330
                            Height =255
                            Name ="Label101"
                            Caption ="IT5:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =5551
                    Top =27091
                    TabIndex =191
                    Name ="IT6"
                    ControlSource ="IT6"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =8623
                            Top =27091
                            Width =330
                            Height =255
                            Name ="Label102"
                            Caption ="IT6:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =5551
                    Top =27431
                    TabIndex =192
                    Name ="IT7"
                    ControlSource ="IT7"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =8623
                            Top =27431
                            Width =330
                            Height =255
                            Name ="Label103"
                            Caption ="IT7:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =5551
                    Top =27771
                    TabIndex =193
                    Name ="IT8"
                    ControlSource ="IT8"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =8623
                            Top =27771
                            Width =330
                            Height =255
                            Name ="Label104"
                            Caption ="IT8:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =5551
                    Top =28111
                    TabIndex =194
                    Name ="IT9"
                    ControlSource ="IT9"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =8623
                            Top =28111
                            Width =330
                            Height =255
                            Name ="Label105"
                            Caption ="IT9:"
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =2
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =34
                    BackStyle =1
                    IMESentenceMode =3
                    Left =7596
                    Top =28780
                    Width =1791
                    Height =255
                    TabIndex =195
                    BorderColor =0
                    Name ="IS1"
                    ControlSource ="IS1"
                    DefaultValue ="=\"نام شركت\""
                    FontName ="Tahoma"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontFamily =34
                            Left =9600
                            Top =28780
                            Width =765
                            Height =240
                            Name ="Label52"
                            Caption ="نام شركت"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =2
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =34
                    BackStyle =1
                    IMESentenceMode =3
                    Left =7596
                    Top =29110
                    Width =1791
                    Height =255
                    TabIndex =196
                    BorderColor =0
                    Name ="IS2"
                    ControlSource ="IS2"
                    DefaultValue ="\"شهر\""
                    FontName ="Tahoma"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontFamily =34
                            Left =9600
                            Top =29110
                            Width =795
                            Height =240
                            Name ="Label54"
                            Caption ="شهر"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =2
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =34
                    BackStyle =1
                    IMESentenceMode =3
                    Left =7596
                    Top =29441
                    Width =1791
                    Height =255
                    TabIndex =197
                    BorderColor =0
                    Name ="IS3"
                    ControlSource ="IS3"
                    DefaultValue ="\"مدير عامل\""
                    FontName ="Tahoma"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontFamily =34
                            Left =9600
                            Top =29441
                            Width =885
                            Height =240
                            Name ="Label56"
                            Caption ="مدير عامل"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =2
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =34
                    BackStyle =1
                    IMESentenceMode =3
                    Left =7596
                    Top =29772
                    Width =1791
                    Height =255
                    TabIndex =198
                    BorderColor =0
                    Name ="IS4"
                    ControlSource ="IS4"
                    DefaultValue ="\"تلفن كارخانه\""
                    FontName ="Tahoma"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontFamily =34
                            Left =9600
                            Top =29772
                            Width =945
                            Height =240
                            Name ="Label58"
                            Caption ="تلفن كارخانه"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =2
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =34
                    BackStyle =1
                    IMESentenceMode =3
                    Left =7596
                    Top =30102
                    Width =1791
                    Height =255
                    TabIndex =199
                    BorderColor =0
                    Name ="IS5"
                    ControlSource ="IS5"
                    DefaultValue ="\"موبايل مدير عامل\""
                    FontName ="Tahoma"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontFamily =34
                            Left =9600
                            Top =30102
                            Width =1245
                            Height =240
                            Name ="Label60"
                            Caption ="موبايل مدير عامل"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =2
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =34
                    BackStyle =1
                    IMESentenceMode =3
                    Left =7596
                    Top =30433
                    Width =1791
                    Height =255
                    TabIndex =200
                    BorderColor =0
                    Name ="IS6"
                    ControlSource ="IS6"
                    DefaultValue ="\"تعداد پرسنل\""
                    FontName ="Tahoma"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontFamily =34
                            Left =9600
                            Top =30433
                            Width =1005
                            Height =240
                            Name ="Label62"
                            Caption ="تعداد پرسنل"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =2
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =34
                    BackStyle =1
                    IMESentenceMode =3
                    Left =7596
                    Top =30764
                    Width =1791
                    Height =255
                    TabIndex =201
                    BorderColor =0
                    Name ="IS7"
                    ControlSource ="IS7"
                    DefaultValue ="\"وضعيت كارخانه\""
                    FontName ="Tahoma"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontFamily =34
                            Left =9600
                            Top =30764
                            Width =1095
                            Height =240
                            Name ="Label64"
                            Caption ="وضعيت كارخانه"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =2
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =34
                    BackStyle =1
                    IMESentenceMode =3
                    Left =7596
                    Top =31095
                    Width =1791
                    Height =255
                    TabIndex =202
                    BorderColor =0
                    Name ="IS8"
                    ControlSource ="IS8"
                    DefaultValue ="\"محصولات\""
                    FontName ="Tahoma"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontFamily =34
                            Left =9600
                            Top =31095
                            Width =1020
                            Height =240
                            Name ="Label66"
                            Caption ="محصولات"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =2
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =34
                    BackStyle =1
                    IMESentenceMode =3
                    Left =7596
                    Top =31425
                    Width =1791
                    Height =255
                    TabIndex =203
                    BorderColor =0
                    Name ="IS9"
                    ControlSource ="IS9"
                    DefaultValue ="\"آدرس\""
                    FontName ="Tahoma"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontFamily =34
                            Left =9600
                            Top =31425
                            Width =795
                            Height =240
                            Name ="Label68"
                            Caption ="آدرس"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =2
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =34
                    BackStyle =1
                    IMESentenceMode =3
                    Left =4251
                    Top =28735
                    Width =1791
                    Height =255
                    TabIndex =204
                    BorderColor =0
                    Name ="IS10"
                    ControlSource ="IS10"
                    DefaultValue ="\"حسابدار\""
                    FontName ="Tahoma"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontFamily =34
                            Left =6255
                            Top =28735
                            Width =765
                            Height =240
                            Name ="Label72"
                            Caption ="حسابدار"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =2
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =34
                    BackStyle =1
                    IMESentenceMode =3
                    Left =4251
                    Top =29065
                    Width =1791
                    Height =255
                    TabIndex =205
                    BorderColor =0
                    Name ="IS11"
                    ControlSource ="IS11"
                    DefaultValue ="\"نرم افزار\""
                    FontName ="Tahoma"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontFamily =34
                            Left =6255
                            Top =29065
                            Width =795
                            Height =240
                            Name ="Label74"
                            Caption ="نرم افزار"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =2
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =34
                    BackStyle =1
                    IMESentenceMode =3
                    Left =4251
                    Top =29396
                    Width =1791
                    Height =255
                    TabIndex =206
                    BorderColor =0
                    Name ="IS12"
                    ControlSource ="IS12"
                    DefaultValue ="\"اشخاص ذي نفوذ\""
                    FontName ="Tahoma"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontFamily =34
                            Left =6255
                            Top =29396
                            Width =1230
                            Height =240
                            Name ="Label76"
                            Caption ="اشخاص ذي نفوذ"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =2
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =34
                    BackStyle =1
                    IMESentenceMode =3
                    Left =4251
                    Top =29727
                    Width =1791
                    Height =255
                    TabIndex =207
                    BorderColor =0
                    Name ="IS13"
                    ControlSource ="IS13"
                    DefaultValue ="\"تلفن كارخانه\""
                    FontName ="Tahoma"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontFamily =34
                            Left =6255
                            Top =29727
                            Width =945
                            Height =240
                            Name ="Label78"
                            Caption ="تلفن كارخانه"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =2
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =34
                    BackStyle =1
                    IMESentenceMode =3
                    Left =4251
                    Top =30057
                    Width =1791
                    Height =255
                    TabIndex =208
                    BorderColor =0
                    Name ="IS14"
                    ControlSource ="IS14"
                    DefaultValue ="\"معرف\""
                    FontName ="Tahoma"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontFamily =34
                            Left =6255
                            Top =30057
                            Width =1245
                            Height =240
                            Name ="Label80"
                            Caption ="معرف"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =2
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =34
                    BackStyle =1
                    IMESentenceMode =3
                    Left =4251
                    Top =30388
                    Width =1791
                    Height =255
                    TabIndex =209
                    BorderColor =0
                    Name ="IS15"
                    ControlSource ="IS15"
                    DefaultValue ="\"وضعيت\""
                    FontName ="Tahoma"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontFamily =34
                            Left =6255
                            Top =30388
                            Width =1005
                            Height =240
                            Name ="Label82"
                            Caption ="وضعيت"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =2
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =34
                    BackStyle =1
                    IMESentenceMode =3
                    Left =4251
                    Top =30719
                    Width =1791
                    Height =255
                    TabIndex =210
                    BorderColor =0
                    Name ="IS16"
                    ControlSource ="IS16"
                    DefaultValue ="\"توضيحات\""
                    FontName ="Tahoma"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontFamily =34
                            Left =6255
                            Top =30719
                            Width =1095
                            Height =240
                            Name ="Label84"
                            Caption ="توضيحات"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =2
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =34
                    KeyboardLanguage =11
                    BackStyle =1
                    IMESentenceMode =3
                    Left =156
                    Top =26760
                    Width =846
                    TabIndex =211
                    BackColor =13434828
                    BorderColor =0
                    Name ="SSMTRTAKM"
                    ControlSource ="SSMTRTAKM"
                    FontName ="Tahoma"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontFamily =34
                            Left =1065
                            Top =26760
                            Width =3690
                            Height =240
                            Name ="Label379"
                            Caption ="تعداد روزهاي تاخير مجاز :"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =2
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =34
                    KeyboardLanguage =11
                    BackStyle =1
                    IMESentenceMode =3
                    Left =156
                    Top =27100
                    Width =846
                    TabIndex =212
                    BackColor =13434828
                    BorderColor =0
                    Name ="SSMTRTAGM"
                    ControlSource ="SSMTRTAGM"
                    FontName ="Tahoma"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontFamily =34
                            Left =1065
                            Top =27105
                            Width =3690
                            Height =240
                            Name ="Label381"
                            Caption ="\011تعداد روزهاي تعجيل مجاز:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =2
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =34
                    KeyboardLanguage =11
                    BackStyle =1
                    IMESentenceMode =3
                    Left =156
                    Top =27440
                    Width =861
                    TabIndex =213
                    BackColor =13434828
                    BorderColor =0
                    Name ="SSMSNDAUTO"
                    ControlSource ="SSMSNDAUTO"
                    FontName ="Tahoma"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontFamily =34
                            Left =1065
                            Top =27435
                            Width =3690
                            Height =240
                            Name ="Label383"
                            Caption ="\011سندهاي تاخير اتومات باشد :"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =2
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =34
                    KeyboardLanguage =11
                    BackStyle =1
                    IMESentenceMode =3
                    Left =156
                    Top =27780
                    Width =846
                    TabIndex =214
                    BackColor =13434828
                    BorderColor =0
                    Name ="SSMTBMON"
                    ControlSource ="SSMTBMON"
                    FontName ="Tahoma"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontFamily =34
                            Left =1065
                            Top =27780
                            Width =3690
                            Height =240
                            Name ="Label385"
                            Caption ="\011تعداد ماه براي كنترل مستر كاركت مثلا قبل از يك ماه:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =2
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =34
                    KeyboardLanguage =11
                    BackStyle =1
                    IMESentenceMode =3
                    Left =150
                    Top =28095
                    Width =846
                    TabIndex =215
                    BackColor =13434828
                    BorderColor =0
                    Name ="SSMDARSAD"
                    ControlSource ="SSMDARSAD"
                    FontName ="Tahoma"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontFamily =34
                            Left =1059
                            Top =28095
                            Width =3690
                            Height =240
                            Name ="Label410"
                            Caption ="\011درصد  كاهش يا افزايش تخفيف  برحسب روز:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =247
                    IMESentenceMode =3
                    Left =5328
                    Top =2154
                    TabIndex =216
                    Name ="HDARKASRTAKHF"
                    ControlSource ="HDARKASRTAKHF"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =7230
                            Top =2154
                            Width =1500
                            Height =255
                            Name ="Label448"
                            Caption ="HDARKASRTAKHF:"
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    KeyboardLanguage =15
                    Left =6462
                    Top =971
                    Width =1140
                    Height =285
                    FontSize =10
                    FontWeight =700
                    TabIndex =217
                    BackColor =16711680
                    BorderColor =3355443
                    ForeColor =16711680
                    Name ="mrcorrect"
                    DefaultValue ="False"
                    FontName ="Tahoma"
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    KeyboardLanguage =15
                    Left =6684
                    Top =1530
                    Width =2160
                    Height =345
                    FontSize =10
                    FontWeight =700
                    TabIndex =218
                    BackColor =16711680
                    BorderColor =3355443
                    ForeColor =16711680
                    Name ="matnsms"
                    FontName ="Tahoma"
                End
                Begin TextBox
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =6122
                    Top =7086
                    TabIndex =219
                    Name ="PUBLICKEY"
                    ControlSource ="PUBLICKEY"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =8564
                            Top =7086
                            Width =960
                            Height =255
                            Name ="Label453"
                            Caption ="PUBLICKEY:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =6122
                    Top =7426
                    TabIndex =220
                    Name ="PRIVIATEKEY"
                    ControlSource ="PRIVIATEKEY"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =8414
                            Top =7426
                            Width =1110
                            Height =255
                            Name ="Label454"
                            Caption ="PRIVIATEKEY:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =7086
                    Top =5102
                    TabIndex =221
                    Name ="MEMORYID"
                    ControlSource ="MEMORYID"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =9558
                            Top =5102
                            Width =930
                            Height =255
                            Name ="Label455"
                            Caption ="MEMORYID:"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    ReadingOrder =1
                    Left =11370
                    Top =24345
                    TabIndex =222
                    Name ="HTAHOL"
                    ControlSource ="HTAHOL"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =10995
                            Top =24315
                            Width =405
                            Height =255
                            Name ="Label457"
                            Caption ="شرايط"
                        End
                    End
                End
            End
        End
        Begin FormFooter
            Height =0
            Name ="FormFooter"
        End
    End
End
CodeBehindForm
' See "Baseknow.cls"
