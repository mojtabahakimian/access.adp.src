Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    AllowEdits = NotDefault
    AllowUpdating =4
    ScrollBars =0
    ViewsAllowed =1
    TabularCharSet =178
    TabularFamily =42
    Orientation =1
    PictureAlignment =5
    PictureSizeMode =1
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =10431
    DatasheetFontHeight =10
    ItemSuffix =13
    DatasheetGridlinesColor =12632256
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
        0xbdd853fbbdb2e240
    End
    RecordSource ="SUDSURY"
    Caption ="سود وزيان جاري"
    DatasheetFontName ="Arial (Arabic)"
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
            Height =300
            BackColor =12632256
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =2
                    Left =9370
                    Width =795
                    Height =300
                    FontWeight =700
                    Name ="HESS Label"
                    Caption ="حساب"
                    FontName ="Zar"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="HESS_Label"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =2
                    Left =3913
                    Width =1320
                    Height =300
                    FontWeight =700
                    Name ="HAS Label"
                    Caption ="مانده بدهكار"
                    FontName ="Zar"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="HAS_Label"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =2
                    Left =1365
                    Width =1410
                    Height =300
                    FontWeight =700
                    Name ="DAR Label"
                    Caption ="مانده بستانكار"
                    FontName ="Zar"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="DAR_Label"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =2
                    Left =7013
                    Width =1065
                    Height =300
                    FontWeight =700
                    Name ="NAME Label"
                    Caption ="نام حساب"
                    FontName ="Zar"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="NAME_Label"
                End
            End
        End
        Begin Section
            Height =345
            BackColor =12632256
            Name ="Detail"
            Begin
                Begin TextBox
                    SpecialEffect =2
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    BackStyle =1
                    NumeralShapes =2
                    Left =9258
                    Top =15
                    Width =1110
                    Height =315
                    ColumnWidth =1110
                    FontWeight =700
                    Name ="HESS"
                    ControlSource ="HESS"
                    FontName ="Titr"
                End
                Begin TextBox
                    DecimalPlaces =0
                    SpecialEffect =2
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    BackStyle =1
                    NumeralShapes =2
                    Left =3403
                    Top =15
                    Width =2430
                    Height =315
                    ColumnWidth =1035
                    FontWeight =700
                    TabIndex =1
                    Name ="HAS"
                    ControlSource ="HAS"
                    Format ="Standard"
                    FontName ="Titr"
                End
                Begin TextBox
                    DecimalPlaces =0
                    SpecialEffect =2
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    BackStyle =1
                    NumeralShapes =2
                    Left =900
                    Top =15
                    Width =2430
                    Height =315
                    ColumnWidth =1035
                    FontWeight =700
                    TabIndex =2
                    Name ="DAR"
                    ControlSource ="DAR"
                    Format ="Standard"
                    FontName ="Titr"
                End
                Begin TextBox
                    SpecialEffect =2
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    BackStyle =1
                    NumeralShapes =2
                    Left =5956
                    Top =15
                    Width =3270
                    Height =315
                    ColumnWidth =2310
                    FontWeight =700
                    TabIndex =3
                    Name ="NAME"
                    ControlSource ="NAME"
                    FontName ="Titr"
                End
            End
        End
        Begin FormFooter
            Height =737
            BackColor =12632256
            Name ="FormFooter"
            Begin
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3415
                    Top =45
                    Width =2475
                    Height =300
                    FontWeight =700
                    BorderColor =0
                    Name ="SHAS"
                    ControlSource ="=Sum([HAS])"
                    Format ="Standard"
                    FontName ="Titr"
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =870
                    Top =45
                    Width =2475
                    Height =300
                    FontWeight =700
                    TabIndex =1
                    BorderColor =0
                    Name ="SDAR"
                    ControlSource ="=Sum([DAR])"
                    Format ="Standard"
                    FontName ="Titr"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =2
                    Left =6075
                    Top =45
                    Width =1125
                    Height =270
                    FontWeight =700
                    Name ="Label10"
                    Caption ="جمع"
                    FontName ="Titr"
                    Tag ="DetachedLabel"
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3405
                    Top =390
                    Width =2490
                    Height =300
                    FontWeight =700
                    TabIndex =2
                    BorderColor =0
                    Name ="Text11"
                    ControlSource ="=[SDAR]-[SHAS]"
                    Format ="Standard"
                    FontName ="Arabic Transparent"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =2
                    Left =6065
                    Top =390
                    Width =1155
                    Height =270
                    FontWeight =700
                    Name ="Label12"
                    Caption ="سود"
                    FontName ="Arabic Transparent"
                    Tag ="DetachedLabel"
                End
            End
        End
    End
End
