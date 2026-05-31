Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    DefaultView =0
    AllowUpdating =4
    ViewsAllowed =1
    TabularCharSet =178
    TabularFamily =42
    BorderStyle =3
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =64
    GridY =64
    Width =6462
    DatasheetFontHeight =10
    DatasheetFontWeight =700
    ItemSuffix =54
    Left =1635
    Top =2385
    Right =8115
    Bottom =5400
    HelpContextId =161
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x934736f29e1ae240
    End
    Caption ="مشخصــــــــات"
    HelpFile ="AMIN98.HLP>MAIN"
    DatasheetFontName ="Arial (Arabic)"
    OnLoad ="[Event Procedure]"
    OnError ="[Event Procedure]"
    Begin
        Begin Label
            BackStyle =0
            TextFontCharSet =178
        End
        Begin CommandButton
            TextFontCharSet =178
            Width =1701
            Height =283
            FontSize =8
            FontWeight =400
            FontName ="MS Sans Serif"
        End
        Begin OptionGroup
            SpecialEffect =3
            Width =1701
            Height =1701
        End
        Begin TextBox
            SpecialEffect =2
            OldBorderStyle =0
            TextFontCharSet =178
            Width =1701
            LabelX =-1701
        End
        Begin ComboBox
            SpecialEffect =2
            TextFontCharSet =178
            Width =1701
            LabelX =-1701
        End
        Begin ToggleButton
            TextFontCharSet =178
            Width =283
            Height =283
            FontSize =8
            FontWeight =400
            FontName ="MS Sans Serif"
        End
        Begin FormHeader
            SpecialEffect =1
            Height =0
            BackColor =-2147483633
            Name ="FormHeader"
        End
        Begin Section
            SpecialEffect =1
            Height =3034
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =0
                    ReadingOrder =2
                    Left =5446
                    Top =165
                    Width =915
                    Height =285
                    Name ="Label6"
                    Caption ="رنگ فرم:"
                    FontName ="Tahoma"
                End
                Begin TextBox
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =2
                    Left =4831
                    Top =105
                    Width =505
                    Height =404
                    ForeColor =16777215
                    Name ="Text10"
                    FontName ="Tahoma"
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontCharSet =0
                    ReadingOrder =2
                    Left =4410
                    Top =90
                    Width =396
                    Height =441
                    TabIndex =1
                    Name ="Command12"
                    Caption ="Command12"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000019000000180000000100040000000000800100000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff0000ff000000ffff00ff000000ff00ff00 ,
                        0xffff0000ffffff00888888888888888888888888888888888888888888888888 ,
                        0x8888888888888888888888888000000078888888888888888888888800000000 ,
                        0x078888888888884888888880000000000008888888888844888888800fbfbfb0 ,
                        0x00008888888888448888880bfbfbfbfb70000888888444448888880fb00fbfbf ,
                        0xbf000888888444448888880b0ee0fbfbfb700088888444448888880f0ee0b00f ,
                        0xbfb000888888884488888880f00b0990fbf700888888884488800000bfbf0990 ,
                        0xbfbf0088888888488800000bfbfbf00bfbfb008888888888800fbfbfbfbfbfbf ,
                        0xbfbf00888888888880fbfbfbfbfbfb00fbfb00888888888880bfb00fbfbfb077 ,
                        0x0fbf00888888888880fb0550fbfbf0770bf707888888888880bf0550bf00bf00 ,
                        0xbfb008888888888880fbf00bf0dd0bfbfb00788888888888807fbfbfb0dd0fbf ,
                        0xb008888888888888880bfbfbfb00fbfb00888888888888888880bfbfbfbfbfb0 ,
                        0x8888888888888888888807fbfbfbf00888888888888888888888800000000888 ,
                        0x8888888888888888000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000000000000000000000000000000000000000000000000000
                    End
                    FontName ="Tahoma"
                    ObjectPalette = Begin
                        0x0003100000000000800000000080000080800000000080008000800000808000 ,
                        0x80808000c0c0c000ff00000000ff0000ffff00000000ff00ff00ff0000ffff00 ,
                        0xffffff0000000000
                    End
                    ControlTipText ="رنگـــــــها"
                End
                Begin CommandButton
                    Default = NotDefault
                    OverlapFlags =85
                    AccessKey =202
                    TextFontCharSet =0
                    ReadingOrder =2
                    Left =1691
                    Top =2533
                    Width =1248
                    Height =397
                    FontWeight =700
                    TabIndex =2
                    Name ="Command13"
                    Caption ="&تاييد"
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"
                End
                Begin CommandButton
                    Cancel = NotDefault
                    OverlapFlags =85
                    AccessKey =199
                    TextFontCharSet =0
                    ReadingOrder =2
                    Left =3054
                    Top =2533
                    Width =1248
                    Height =397
                    FontWeight =700
                    TabIndex =3
                    Name ="Command14"
                    Caption ="&انصراف"
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"
                End
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =0
                    ReadingOrder =2
                    Left =5452
                    Top =600
                    Width =915
                    Height =285
                    Name ="Label15"
                    Caption ="رنگ زمينه:"
                    FontName ="Tahoma"
                End
                Begin TextBox
                    TabStop = NotDefault
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextFontCharSet =0
                    ReadingOrder =2
                    Left =4837
                    Top =540
                    Width =505
                    Height =404
                    TabIndex =4
                    ForeColor =16777215
                    Name ="Text12"
                    DefaultValue ="1"
                    FontName ="Tahoma"
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontCharSet =0
                    ReadingOrder =2
                    Left =4416
                    Top =525
                    Width =396
                    Height =441
                    TabIndex =5
                    Name ="Command17"
                    Caption ="Command12"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000019000000180000000100040000000000800100000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff0000ff000000ffff00ff000000ff00ff00 ,
                        0xffff0000ffffff00888888888888888888888888888888888888888888888888 ,
                        0x8888888888888888888888888000000078888888888888888888888800000000 ,
                        0x078888888888884888888880000000000008888888888844888888800fbfbfb0 ,
                        0x00008888888888448888880bfbfbfbfb70000888888444448888880fb00fbfbf ,
                        0xbf000888888444448888880b0ee0fbfbfb700088888444448888880f0ee0b00f ,
                        0xbfb000888888884488888880f00b0990fbf700888888884488800000bfbf0990 ,
                        0xbfbf0088888888488800000bfbfbf00bfbfb008888888888800fbfbfbfbfbfbf ,
                        0xbfbf00888888888880fbfbfbfbfbfb00fbfb00888888888880bfb00fbfbfb077 ,
                        0x0fbf00888888888880fb0550fbfbf0770bf707888888888880bf0550bf00bf00 ,
                        0xbfb008888888888880fbf00bf0dd0bfbfb00788888888888807fbfbfb0dd0fbf ,
                        0xb008888888888888880bfbfbfb00fbfb00888888888888888880bfbfbfbfbfb0 ,
                        0x8888888888888888888807fbfbfbf00888888888888888888888800000000888 ,
                        0x8888888888888888000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000000000000000000000000000000000000000000000000000
                    End
                    FontName ="Tahoma"
                    ObjectPalette = Begin
                        0x0003100000000000800000000080000080800000000080008000800000808000 ,
                        0x80808000c0c0c000ff00000000ff0000ffff00000000ff00ff00ff0000ffff00 ,
                        0xffffff0000000000
                    End
                    ControlTipText ="رنگـــــــها"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1010
                    Top =170
                    Width =2121
                    Height =300
                    TabIndex =6
                    Name ="Text18"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT DISTINCTROW font.قلم FROM font GROUP BY font.قلم HAVING (((font.قلم) Is N"
                        "ot Null)) ORDER BY font.قلم;"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="\"Arial (Arabic)\""
                    FontName ="Tahoma"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            ReadingOrder =2
                            Left =3241
                            Top =165
                            Width =1140
                            Height =285
                            Name ="Label19"
                            Caption ="قــلم هاي متن:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    ReadingOrder =1
                    ScrollBarAlign =1
                    NumeralShapes =3
                    Left =172
                    Top =170
                    Width =786
                    Height =300
                    TabIndex =7
                    Name ="Text21"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT DISTINCTROW font.اندازه FROM font GROUP BY font.اندازه ORDER BY font.اندا"
                        "زه;"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="8"
                    FontName ="Tahoma"
                    InputMask ="###"
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    BorderWidth =4
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =2
                    Left =172
                    Top =631
                    Width =2948
                    Height =680
                    Name ="Label24"
                    Caption ="آمريــكا هيچ غلطي نمي تواند  بكند        امام خميني(ره)"
                    FontName ="Tahoma"
                End
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =0
                    ReadingOrder =2
                    Left =3744
                    Top =680
                    Width =540
                    Height =285
                    Name ="Label25"
                    Caption ="نمونه:"
                    FontName ="Tahoma"
                End
                Begin ToggleButton
                    OverlapFlags =85
                    TextFontCharSet =0
                    ReadingOrder =2
                    Left =4725
                    Top =1860
                    Width =627
                    Height =385
                    TabIndex =8
                    Name ="Toggle27"
                    Caption ="پررنگ"
                    FontName ="Tahoma"
                    OnClick ="[Event Procedure]"
                End
                Begin ToggleButton
                    OverlapFlags =85
                    TextFontCharSet =0
                    ReadingOrder =2
                    Left =4050
                    Top =1860
                    Width =672
                    Height =385
                    TabIndex =9
                    Name ="Toggle28"
                    Caption ="مورب"
                    FontName ="Tahoma"
                    OnClick ="[Event Procedure]"
                End
                Begin ToggleButton
                    OverlapFlags =85
                    TextFontCharSet =0
                    ReadingOrder =2
                    Left =3420
                    Top =1860
                    Width =630
                    Height =390
                    TabIndex =10
                    Name ="Toggle29"
                    Caption ="زير خط"
                    FontName ="Tahoma"
                    OnClick ="[Event Procedure]"
                End
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =0
                    ReadingOrder =2
                    Left =5452
                    Top =1050
                    Width =915
                    Height =285
                    Name ="Label31"
                    Caption ="رنگ متن:"
                    FontName ="Tahoma"
                End
                Begin TextBox
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    ReadingOrder =2
                    Left =4837
                    Top =990
                    Width =505
                    Height =404
                    TabIndex =11
                    ForeColor =16777215
                    Name ="Text14"
                    BeforeUpdate ="[Event Procedure]"
                    DefaultValue ="1"
                    FontName ="Tahoma"
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontCharSet =0
                    ReadingOrder =2
                    Left =4416
                    Top =975
                    Width =396
                    Height =441
                    TabIndex =12
                    Name ="Command33"
                    Caption ="Command12"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000019000000180000000100040000000000800100000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff0000ff000000ffff00ff000000ff00ff00 ,
                        0xffff0000ffffff00888888888888888888888888888888888888888888888888 ,
                        0x8888888888888888888888888000000078888888888888888888888800000000 ,
                        0x078888888888884888888880000000000008888888888844888888800fbfbfb0 ,
                        0x00008888888888448888880bfbfbfbfb70000888888444448888880fb00fbfbf ,
                        0xbf000888888444448888880b0ee0fbfbfb700088888444448888880f0ee0b00f ,
                        0xbfb000888888884488888880f00b0990fbf700888888884488800000bfbf0990 ,
                        0xbfbf0088888888488800000bfbfbf00bfbfb008888888888800fbfbfbfbfbfbf ,
                        0xbfbf00888888888880fbfbfbfbfbfb00fbfb00888888888880bfb00fbfbfb077 ,
                        0x0fbf00888888888880fb0550fbfbf0770bf707888888888880bf0550bf00bf00 ,
                        0xbfb008888888888880fbf00bf0dd0bfbfb00788888888888807fbfbfb0dd0fbf ,
                        0xb008888888888888880bfbfbfb00fbfb00888888888888888880bfbfbfbfbfb0 ,
                        0x8888888888888888888807fbfbfbf00888888888888888888888800000000888 ,
                        0x8888888888888888000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000000000000000000000000000000000000000000000000000
                    End
                    FontName ="Tahoma"
                    ObjectPalette = Begin
                        0x0003100000000000800000000080000080800000000080008000800000808000 ,
                        0x80808000c0c0c000ff00000000ff0000ffff00000000ff00ff00ff0000ffff00 ,
                        0xffffff0000000000
                    End
                    ControlTipText ="رنگـــــــها"
                End
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =0
                    ReadingOrder =2
                    Left =5445
                    Top =1710
                    Width =975
                    Height =285
                    Name ="Label39"
                    Caption ="حالت نمايش:"
                    FontName ="Tahoma"
                End
                Begin ToggleButton
                    OverlapFlags =85
                    TextFontCharSet =0
                    ReadingOrder =2
                    Left =3427
                    Top =1470
                    Width =386
                    Height =386
                    TabIndex =13
                    Name ="Toggle40"
                    PictureData = Begin
                        0x2800000019000000180000000100040000000000800100000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff0000ff000000ffff00ff000000ff00ff00 ,
                        0xffff0000ffffff00888888888888888888888888888888888888888888888888 ,
                        0x8888888888888000888888888888888888888888888880088888888888888888 ,
                        0x8888888888888800888888888888888888888888888888808888888888888888 ,
                        0x8888888888888888888888888888888888888888888888888888888888888888 ,
                        0x88888888888888888888888888888888888888888888808888888fffffffffff ,
                        0xffff8888888880888888788888888888888f8888888880888888788888888888 ,
                        0x888f8888888880008888788888888888888f8888888888888888788888888888 ,
                        0x888f8888888800008888788888888888888f8888888887088888777777777777 ,
                        0x7778888888888807888888888888888888888888888888708888888888888888 ,
                        0x8888888888888880888888888888888888888888888888878888888888888888 ,
                        0x8888888888888888888888888888888888888888888888888888888888888888 ,
                        0x8888888888888888888888888888888888888888888888888888888888888888 ,
                        0x8888888888888888000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End
                    FontName ="Tahoma"
                    OnClick ="[Event Procedure]"
                    ObjectPalette = Begin
                        0x0003100000000000800000000080000080800000000080008000800000808000 ,
                        0x80808000c0c0c000ff00000000ff0000ffff00000000ff00ff00ff0000ffff00 ,
                        0xffffff0000000000
                    End
                End
                Begin ToggleButton
                    OverlapFlags =85
                    TextFontCharSet =0
                    ReadingOrder =2
                    Left =3804
                    Top =1474
                    Width =386
                    Height =386
                    TabIndex =14
                    Name ="Toggle41"
                    PictureData = Begin
                        0x2800000019000000180000000100040000000000800100000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff0000ff000000ffff00ff000000ff00ff00 ,
                        0xffff0000ffffff00888888888888888888888888888888888888888888888888 ,
                        0x8888888888888888888888888888888888888888888888888888888888888888 ,
                        0x8888888888888888888888888888888888888888888888888888888888888888 ,
                        0x8888888888888888888888888888888888888888888888888888888888888888 ,
                        0x8888888888888888888888888888888888888888888888888888877777777777 ,
                        0x777788888888888f8888f8888888888888878888888888788888f88888888888 ,
                        0x88878888888888788888f8888888888888878888888888788888f88888888888 ,
                        0x88878888888888788888f8888888888888878888888888788888ffffffffffff ,
                        0xfff8888888888877888888888888888888888888888888888888888888888888 ,
                        0x8888888888888888888888888888888888888888888888888888888888888888 ,
                        0x8888888888888888888888888888888888888888888888888888888888888888 ,
                        0x8888888888888888888888888888888888888888888888888888888888888888 ,
                        0x8888888888888888000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End
                    FontName ="Tahoma"
                    OnClick ="[Event Procedure]"
                    ObjectPalette = Begin
                        0x0003100000000000800000000080000080800000000080008000800000808000 ,
                        0x80808000c0c0c000ff00000000ff0000ffff00000000ff00ff00ff0000ffff00 ,
                        0xffffff0000000000
                    End
                End
                Begin ToggleButton
                    OverlapFlags =85
                    TextFontCharSet =0
                    ReadingOrder =2
                    Left =4192
                    Top =1474
                    Width =386
                    Height =386
                    TabIndex =15
                    Name ="Toggle42"
                    PictureData = Begin
                        0x28000000130000000e0000000100080000000000180100000000000000000000 ,
                        0x0001000000000000000000000000800000800000008080008000000080008000 ,
                        0x80800000c0c0c000c0dcc000f0caa60004040400080808000c0c0c0011111100 ,
                        0x161616001c1c1c002222220029292900555555004d4d4d004242420039393900 ,
                        0x807cff005050ff009300d600ffeccc00c6d6ef00d6e7e70090a9ad0000003300 ,
                        0x00006600000099000000cc00003300000033330000336600003399000033cc00 ,
                        0x0033ff00006600000066330000666600006699000066cc000066ff0000990000 ,
                        0x0099330000996600009999000099cc000099ff0000cc000000cc330000cc6600 ,
                        0x00cc990000cccc0000ccff0000ff660000ff990000ffcc003300000033003300 ,
                        0x33006600330099003300cc003300ff0033330000333333003333660033339900 ,
                        0x3333cc003333ff00336600003366330033666600336699003366cc003366ff00 ,
                        0x339900003399330033996600339999003399cc003399ff0033cc000033cc3300 ,
                        0x33cc660033cc990033cccc0033ccff0033ff330033ff660033ff990033ffcc00 ,
                        0x33ffff00660000006600330066006600660099006600cc006600ff0066330000 ,
                        0x6633330066336600663399006633cc006633ff00666600006666330066666600 ,
                        0x666699006666cc00669900006699330066996600669999006699cc006699ff00 ,
                        0x66cc000066cc330066cc990066cccc0066ccff0066ff000066ff330066ff9900 ,
                        0x66ffcc00cc00ff00ff00cc009999000099339900990099009900cc0099000000 ,
                        0x99333300990066009933cc009900ff0099660000996633009933660099669900 ,
                        0x9966cc009933ff009999330099996600999999009999cc009999ff0099cc0000 ,
                        0x99cc330066cc660099cc990099cccc0099ccff0099ff000099ff330099cc6600 ,
                        0x99ff990099ffcc0099ffff00cc00000099003300cc006600cc009900cc00cc00 ,
                        0x99330000cc333300cc336600cc339900cc33cc00cc33ff00cc660000cc663300 ,
                        0x99666600cc669900cc66cc009966ff00cc990000cc993300cc996600cc999900 ,
                        0xcc99cc00cc99ff00cccc0000cccc3300cccc6600cccc9900cccccc00ccccff00 ,
                        0xccff0000ccff330099ff6600ccff9900ccffcc00ccffff00cc003300ff006600 ,
                        0xff009900cc330000ff333300ff336600ff339900ff33cc00ff33ff00ff660000 ,
                        0xff663300cc666600ff669900ff66cc00cc66ff00ff990000ff993300ff996600 ,
                        0xff999900ff99cc00ff99ff00ffcc0000ffcc3300ffcc6600ffcc9900ffcccc00 ,
                        0xffccff00ffff3300ccff6600ffff9900ffffcc006666ff0066ff660066ffff00 ,
                        0xff666600ff66ff00ffff66002100a5005f5f5f00777777008686860096969600 ,
                        0xcbcbcb00b2b2b200d7d7d700dddddd00e3e3e300eaeaea00f1f1f100f8f8f800 ,
                        0xf0fbff00a4a0a000808080000000ff0000ff000000ffff00ff000000ff00ff00 ,
                        0xffff0000ffffff00070707070707070707070707070707070707070007070707 ,
                        0x0707070707070707070707070707070007070707070707070707070707070707 ,
                        0x0707070007070707070707070707070707070707070707000707070707070707 ,
                        0x0707070707070707070707000707070000000000000000000000000707070700 ,
                        0x0707070007070707070707070707000707070700070707000707070707070707 ,
                        0x0707000707070700070707000707070707070707070700070707070007070700 ,
                        0x0000000000000000000000070707070007070707070707070707070707070707 ,
                        0x0707070007070707070707070707070707070707070707000707070707070707 ,
                        0x0707070707070707070707000707070707070707070707070707070707070700 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000
                    End
                    FontName ="Tahoma"
                    OnClick ="[Event Procedure]"
                    ObjectPalette = Begin
                        0x0003000100000000800000000080000080800000000080008000800000808000 ,
                        0xc0c0c000c0dcc000a6caf00004040400080808000c0c0c001111110016161600 ,
                        0x1c1c1c002222220029292900555555004d4d4d004242420039393900ff7c8000 ,
                        0xff505000d6009300ccecff00efd6c600e7e7d600ada990003300000066000000 ,
                        0x99000000cc00000000330000333300006633000099330000cc330000ff330000 ,
                        0x00660000336600006666000099660000cc660000ff6600000099000033990000 ,
                        0x6699000099990000cc990000ff99000000cc000033cc000066cc000099cc0000 ,
                        0xcccc0000ffcc000066ff000099ff0000ccff0000000033003300330066003300 ,
                        0x99003300cc003300ff00330000333300333333006633330099333300cc333300 ,
                        0xff33330000663300336633006666330099663300cc663300ff66330000993300 ,
                        0x339933006699330099993300cc993300ff99330000cc330033cc330066cc3300 ,
                        0x99cc3300cccc3300ffcc330033ff330066ff330099ff3300ccff3300ffff3300 ,
                        0x00006600330066006600660099006600cc006600ff0066000033660033336600 ,
                        0x6633660099336600cc336600ff33660000666600336666006666660099666600 ,
                        0xcc66660000996600339966006699660099996600cc996600ff99660000cc6600 ,
                        0x33cc660099cc6600cccc6600ffcc660000ff660033ff660099ff6600ccff6600 ,
                        0xff00cc00cc00ff00009999009933990099009900cc0099000000990033339900 ,
                        0x66009900cc339900ff00990000669900336699006633990099669900cc669900 ,
                        0xff339900339999006699990099999900cc999900ff99990000cc990033cc9900 ,
                        0x66cc660099cc9900cccc9900ffcc990000ff990033ff990066cc990099ff9900 ,
                        0xccff9900ffff99000000cc00330099006600cc009900cc00cc00cc0000339900 ,
                        0x3333cc006633cc009933cc00cc33cc00ff33cc000066cc003366cc0066669900 ,
                        0x9966cc00cc66cc00ff6699000099cc003399cc006699cc009999cc00cc99cc00 ,
                        0xff99cc0000cccc0033cccc0066cccc0099cccc00cccccc00ffcccc0000ffcc00 ,
                        0x33ffcc0066ff990099ffcc00ccffcc00ffffcc003300cc006600ff009900ff00 ,
                        0x0033cc003333ff006633ff009933ff00cc33ff00ff33ff000066ff003366ff00 ,
                        0x6666cc009966ff00cc66ff00ff66cc000099ff003399ff006699ff009999ff00 ,
                        0xcc99ff00ff99ff0000ccff0033ccff0066ccff0099ccff00ccccff00ffccff00 ,
                        0x33ffff0066ffcc0099ffff00ccffff00ff66660066ff6600ffff66006666ff00 ,
                        0xff66ff0066ffff00a50021005f5f5f00777777008686860096969600cbcbcb00 ,
                        0xb2b2b200d7d7d700dddddd00e3e3e300eaeaea00f1f1f100f8f8f800fffbf000 ,
                        0xa0a0a40080808000ff00000000ff0000ffff00000000ff00ff00ff0000ffff00 ,
                        0xffffff0000000000
                    End
                    Picture ="C:\\MSOffice\\Access\\Bitmaps\\Dbwiz\\TRANSPAR.BMP"
                End
                Begin ToggleButton
                    OverlapFlags =85
                    TextFontCharSet =0
                    ReadingOrder =2
                    Left =4582
                    Top =1474
                    Width =386
                    Height =386
                    TabIndex =16
                    Name ="Toggle43"
                    PictureData = Begin
                        0x2800000013000000110000000100080000000000540100000000000000000000 ,
                        0x0001000000000000000000000000800000800000008080008000000080008000 ,
                        0x80800000c0c0c000c0dcc000f0caa60004040400080808000c0c0c0011111100 ,
                        0x161616001c1c1c002222220029292900555555004d4d4d004242420039393900 ,
                        0x807cff005050ff009300d600ffeccc00c6d6ef00d6e7e70090a9ad0000003300 ,
                        0x00006600000099000000cc00003300000033330000336600003399000033cc00 ,
                        0x0033ff00006600000066330000666600006699000066cc000066ff0000990000 ,
                        0x0099330000996600009999000099cc000099ff0000cc000000cc330000cc6600 ,
                        0x00cc990000cccc0000ccff0000ff660000ff990000ffcc003300000033003300 ,
                        0x33006600330099003300cc003300ff0033330000333333003333660033339900 ,
                        0x3333cc003333ff00336600003366330033666600336699003366cc003366ff00 ,
                        0x339900003399330033996600339999003399cc003399ff0033cc000033cc3300 ,
                        0x33cc660033cc990033cccc0033ccff0033ff330033ff660033ff990033ffcc00 ,
                        0x33ffff00660000006600330066006600660099006600cc006600ff0066330000 ,
                        0x6633330066336600663399006633cc006633ff00666600006666330066666600 ,
                        0x666699006666cc00669900006699330066996600669999006699cc006699ff00 ,
                        0x66cc000066cc330066cc990066cccc0066ccff0066ff000066ff330066ff9900 ,
                        0x66ffcc00cc00ff00ff00cc009999000099339900990099009900cc0099000000 ,
                        0x99333300990066009933cc009900ff0099660000996633009933660099669900 ,
                        0x9966cc009933ff009999330099996600999999009999cc009999ff0099cc0000 ,
                        0x99cc330066cc660099cc990099cccc0099ccff0099ff000099ff330099cc6600 ,
                        0x99ff990099ffcc0099ffff00cc00000099003300cc006600cc009900cc00cc00 ,
                        0x99330000cc333300cc336600cc339900cc33cc00cc33ff00cc660000cc663300 ,
                        0x99666600cc669900cc66cc009966ff00cc990000cc993300cc996600cc999900 ,
                        0xcc99cc00cc99ff00cccc0000cccc3300cccc6600cccc9900cccccc00ccccff00 ,
                        0xccff0000ccff330099ff6600ccff9900ccffcc00ccffff00cc003300ff006600 ,
                        0xff009900cc330000ff333300ff336600ff339900ff33cc00ff33ff00ff660000 ,
                        0xff663300cc666600ff669900ff66cc00cc66ff00ff990000ff993300ff996600 ,
                        0xff999900ff99cc00ff99ff00ffcc0000ffcc3300ffcc6600ffcc9900ffcccc00 ,
                        0xffccff00ffff3300ccff6600ffff9900ffffcc006666ff0066ff660066ffff00 ,
                        0xff666600ff66ff00ffff66002100a5005f5f5f00777777008686860096969600 ,
                        0xcbcbcb00b2b2b200d7d7d700dddddd00e3e3e300eaeaea00f1f1f100f8f8f800 ,
                        0xf0fbff00a4a0a000808080000000ff0000ff000000ffff00ff000000ff00ff00 ,
                        0xffff0000ffffff00070707070707070707070707070707070707070007070707 ,
                        0x0707070707070707070707070707070007070707070707070707070707070707 ,
                        0x0707070007070707070707070707070707070707070707000707070707000000 ,
                        0x0000000000000000000707000707070707000000000000000000000000070700 ,
                        0x0707070000000000000000000000000000070700070707000707070707070707 ,
                        0x0707000000070700070707000707070707070707070700000007070007070700 ,
                        0x0707070707070707070700070707070007070700000000000000000000000007 ,
                        0x0707070007070707070707070707070707070707070707000707070707070707 ,
                        0x0707070707070707070707000707070707070707070707070707070707070700 ,
                        0x0707070707070707070707070707070707070700070707070707070707070707 ,
                        0x0707070707070700070707070707070707070707070707070707070000000000 ,
                        0x000000000000000000000000000000000000000000000000
                    End
                    FontName ="Tahoma"
                    OnClick ="[Event Procedure]"
                    ObjectPalette = Begin
                        0x0003000100000000800000000080000080800000000080008000800000808000 ,
                        0xc0c0c000c0dcc000a6caf00004040400080808000c0c0c001111110016161600 ,
                        0x1c1c1c002222220029292900555555004d4d4d004242420039393900ff7c8000 ,
                        0xff505000d6009300ccecff00efd6c600e7e7d600ada990003300000066000000 ,
                        0x99000000cc00000000330000333300006633000099330000cc330000ff330000 ,
                        0x00660000336600006666000099660000cc660000ff6600000099000033990000 ,
                        0x6699000099990000cc990000ff99000000cc000033cc000066cc000099cc0000 ,
                        0xcccc0000ffcc000066ff000099ff0000ccff0000000033003300330066003300 ,
                        0x99003300cc003300ff00330000333300333333006633330099333300cc333300 ,
                        0xff33330000663300336633006666330099663300cc663300ff66330000993300 ,
                        0x339933006699330099993300cc993300ff99330000cc330033cc330066cc3300 ,
                        0x99cc3300cccc3300ffcc330033ff330066ff330099ff3300ccff3300ffff3300 ,
                        0x00006600330066006600660099006600cc006600ff0066000033660033336600 ,
                        0x6633660099336600cc336600ff33660000666600336666006666660099666600 ,
                        0xcc66660000996600339966006699660099996600cc996600ff99660000cc6600 ,
                        0x33cc660099cc6600cccc6600ffcc660000ff660033ff660099ff6600ccff6600 ,
                        0xff00cc00cc00ff00009999009933990099009900cc0099000000990033339900 ,
                        0x66009900cc339900ff00990000669900336699006633990099669900cc669900 ,
                        0xff339900339999006699990099999900cc999900ff99990000cc990033cc9900 ,
                        0x66cc660099cc9900cccc9900ffcc990000ff990033ff990066cc990099ff9900 ,
                        0xccff9900ffff99000000cc00330099006600cc009900cc00cc00cc0000339900 ,
                        0x3333cc006633cc009933cc00cc33cc00ff33cc000066cc003366cc0066669900 ,
                        0x9966cc00cc66cc00ff6699000099cc003399cc006699cc009999cc00cc99cc00 ,
                        0xff99cc0000cccc0033cccc0066cccc0099cccc00cccccc00ffcccc0000ffcc00 ,
                        0x33ffcc0066ff990099ffcc00ccffcc00ffffcc003300cc006600ff009900ff00 ,
                        0x0033cc003333ff006633ff009933ff00cc33ff00ff33ff000066ff003366ff00 ,
                        0x6666cc009966ff00cc66ff00ff66cc000099ff003399ff006699ff009999ff00 ,
                        0xcc99ff00ff99ff0000ccff0033ccff0066ccff0099ccff00ccccff00ffccff00 ,
                        0x33ffff0066ffcc0099ffff00ccffff00ff66660066ff6600ffff66006666ff00 ,
                        0xff66ff0066ffff00a50021005f5f5f00777777008686860096969600cbcbcb00 ,
                        0xb2b2b200d7d7d700dddddd00e3e3e300eaeaea00f1f1f100f8f8f800fffbf000 ,
                        0xa0a0a40080808000ff00000000ff0000ffff00000000ff00ff00ff0000ffff00 ,
                        0xffffff0000000000
                    End
                    Picture ="C:\\MSOffice\\Access\\Bitmaps\\Dbwiz\\SHADOW.BMP"
                End
                Begin ToggleButton
                    OverlapFlags =85
                    TextFontCharSet =0
                    ReadingOrder =2
                    Left =4972
                    Top =1470
                    Width =386
                    Height =386
                    TabIndex =17
                    Name ="Toggle44"
                    PictureData = Begin
                        0x28000000110000000f00000001000800000000002c0100000000000000000000 ,
                        0x0001000000000000000000000000800000800000008080008000000080008000 ,
                        0x80800000c0c0c000c0dcc000f0caa60004040400080808000c0c0c0011111100 ,
                        0x161616001c1c1c002222220029292900555555004d4d4d004242420039393900 ,
                        0x807cff005050ff009300d600ffeccc00c6d6ef00d6e7e70090a9ad0000003300 ,
                        0x00006600000099000000cc00003300000033330000336600003399000033cc00 ,
                        0x0033ff00006600000066330000666600006699000066cc000066ff0000990000 ,
                        0x0099330000996600009999000099cc000099ff0000cc000000cc330000cc6600 ,
                        0x00cc990000cccc0000ccff0000ff660000ff990000ffcc003300000033003300 ,
                        0x33006600330099003300cc003300ff0033330000333333003333660033339900 ,
                        0x3333cc003333ff00336600003366330033666600336699003366cc003366ff00 ,
                        0x339900003399330033996600339999003399cc003399ff0033cc000033cc3300 ,
                        0x33cc660033cc990033cccc0033ccff0033ff330033ff660033ff990033ffcc00 ,
                        0x33ffff00660000006600330066006600660099006600cc006600ff0066330000 ,
                        0x6633330066336600663399006633cc006633ff00666600006666330066666600 ,
                        0x666699006666cc00669900006699330066996600669999006699cc006699ff00 ,
                        0x66cc000066cc330066cc990066cccc0066ccff0066ff000066ff330066ff9900 ,
                        0x66ffcc00cc00ff00ff00cc009999000099339900990099009900cc0099000000 ,
                        0x99333300990066009933cc009900ff0099660000996633009933660099669900 ,
                        0x9966cc009933ff009999330099996600999999009999cc009999ff0099cc0000 ,
                        0x99cc330066cc660099cc990099cccc0099ccff0099ff000099ff330099cc6600 ,
                        0x99ff990099ffcc0099ffff00cc00000099003300cc006600cc009900cc00cc00 ,
                        0x99330000cc333300cc336600cc339900cc33cc00cc33ff00cc660000cc663300 ,
                        0x99666600cc669900cc66cc009966ff00cc990000cc993300cc996600cc999900 ,
                        0xcc99cc00cc99ff00cccc0000cccc3300cccc6600cccc9900cccccc00ccccff00 ,
                        0xccff0000ccff330099ff6600ccff9900ccffcc00ccffff00cc003300ff006600 ,
                        0xff009900cc330000ff333300ff336600ff339900ff33cc00ff33ff00ff660000 ,
                        0xff663300cc666600ff669900ff66cc00cc66ff00ff990000ff993300ff996600 ,
                        0xff999900ff99cc00ff99ff00ffcc0000ffcc3300ffcc6600ffcc9900ffcccc00 ,
                        0xffccff00ffff3300ccff6600ffff9900ffffcc006666ff0066ff660066ffff00 ,
                        0xff666600ff66ff00ffff66002100a5005f5f5f00777777008686860096969600 ,
                        0xcbcbcb00b2b2b200d7d7d700dddddd00e3e3e300eaeaea00f1f1f100f8f8f800 ,
                        0xf0fbff00a4a0a000808080000000ff0000ff000000ffff00ff000000ff00ff00 ,
                        0xffff0000ffffff00070707070707070707070707070707070700000007070707 ,
                        0x0707070707070707070707070700000007070707070707070707070707070707 ,
                        0x0700000007070707070707070707070707070707070000000707070707070707 ,
                        0x07070707070707070700000007070707ffffffffffffffffffffff0707000000 ,
                        0x0707f8fff8f8f8f8f8f8f8f8f8f8ff07070000000707f8ff0707070707070707 ,
                        0x07f8ff07070000000707f8ff070707070707070707f8ff07070000000707f8ff ,
                        0xffffffffffffffffffff0707070000000707f8f8f8f8f8f8f8f8f8f8f8f80707 ,
                        0x0700000007070707070707070707070707070707070000000707070707070707 ,
                        0x0707070707070707070000000707070707070707070707070707070707000000 ,
                        0x0707070707070707070707070707070707000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000000000000000000000000000000000000000000000
                    End
                    FontName ="Tahoma"
                    OnClick ="[Event Procedure]"
                    ObjectPalette = Begin
                        0x0003000100000000800000000080000080800000000080008000800000808000 ,
                        0xc0c0c000c0dcc000a6caf00004040400080808000c0c0c001111110016161600 ,
                        0x1c1c1c002222220029292900555555004d4d4d004242420039393900ff7c8000 ,
                        0xff505000d6009300ccecff00efd6c600e7e7d600ada990003300000066000000 ,
                        0x99000000cc00000000330000333300006633000099330000cc330000ff330000 ,
                        0x00660000336600006666000099660000cc660000ff6600000099000033990000 ,
                        0x6699000099990000cc990000ff99000000cc000033cc000066cc000099cc0000 ,
                        0xcccc0000ffcc000066ff000099ff0000ccff0000000033003300330066003300 ,
                        0x99003300cc003300ff00330000333300333333006633330099333300cc333300 ,
                        0xff33330000663300336633006666330099663300cc663300ff66330000993300 ,
                        0x339933006699330099993300cc993300ff99330000cc330033cc330066cc3300 ,
                        0x99cc3300cccc3300ffcc330033ff330066ff330099ff3300ccff3300ffff3300 ,
                        0x00006600330066006600660099006600cc006600ff0066000033660033336600 ,
                        0x6633660099336600cc336600ff33660000666600336666006666660099666600 ,
                        0xcc66660000996600339966006699660099996600cc996600ff99660000cc6600 ,
                        0x33cc660099cc6600cccc6600ffcc660000ff660033ff660099ff6600ccff6600 ,
                        0xff00cc00cc00ff00009999009933990099009900cc0099000000990033339900 ,
                        0x66009900cc339900ff00990000669900336699006633990099669900cc669900 ,
                        0xff339900339999006699990099999900cc999900ff99990000cc990033cc9900 ,
                        0x66cc660099cc9900cccc9900ffcc990000ff990033ff990066cc990099ff9900 ,
                        0xccff9900ffff99000000cc00330099006600cc009900cc00cc00cc0000339900 ,
                        0x3333cc006633cc009933cc00cc33cc00ff33cc000066cc003366cc0066669900 ,
                        0x9966cc00cc66cc00ff6699000099cc003399cc006699cc009999cc00cc99cc00 ,
                        0xff99cc0000cccc0033cccc0066cccc0099cccc00cccccc00ffcccc0000ffcc00 ,
                        0x33ffcc0066ff990099ffcc00ccffcc00ffffcc003300cc006600ff009900ff00 ,
                        0x0033cc003333ff006633ff009933ff00cc33ff00ff33ff000066ff003366ff00 ,
                        0x6666cc009966ff00cc66ff00ff66cc000099ff003399ff006699ff009999ff00 ,
                        0xcc99ff00ff99ff0000ccff0033ccff0066ccff0099ccff00ccccff00ffccff00 ,
                        0x33ffff0066ffcc0099ffff00ccffff00ff66660066ff6600ffff66006666ff00 ,
                        0xff66ff0066ffff00a50021005f5f5f00777777008686860096969600cbcbcb00 ,
                        0xb2b2b200d7d7d700dddddd00e3e3e300eaeaea00f1f1f100f8f8f800fffbf000 ,
                        0xa0a0a40080808000ff00000000ff0000ffff00000000ff00ff00ff0000ffff00 ,
                        0xffffff0000000000
                    End
                    Picture ="C:\\MSOffice\\Access\\Bitmaps\\Dbwiz\\DRAW.BMP"
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2172
                    Top =1470
                    Width =666
                    Height =300
                    TabIndex =18
                    Name ="Text45"
                    Format ="Standard"
                    FontName ="Tahoma"
                    InputMask ="##.###"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            ReadingOrder =2
                            Left =2941
                            Top =1470
                            Width =375
                            Height =285
                            Name ="Label46"
                            Caption ="بالا:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2172
                    Top =1890
                    Width =666
                    Height =300
                    TabIndex =19
                    Name ="Text47"
                    Format ="Standard"
                    FontName ="Tahoma"
                    InputMask ="##.###"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            ReadingOrder =2
                            Left =2937
                            Top =1886
                            Width =405
                            Height =285
                            Name ="Label48"
                            Caption ="چپ:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =765
                    Top =1875
                    Width =666
                    Height =300
                    TabIndex =20
                    Name ="Text49"
                    Format ="Standard"
                    FontName ="Tahoma"
                    InputMask ="##.###"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            ReadingOrder =2
                            Left =1662
                            Top =1888
                            Width =465
                            Height =285
                            Name ="Label50"
                            Caption ="طول:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =762
                    Top =1481
                    Width =666
                    Height =300
                    TabIndex =21
                    Name ="Text51"
                    Format ="Standard"
                    FontName ="Tahoma"
                    InputMask ="##.###"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            ReadingOrder =2
                            Left =1527
                            Top =1481
                            Width =585
                            Height =285
                            Name ="Label52"
                            Caption ="عرض:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin Label
                    FontItalic = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    Left =256
                    Top =1727
                    Width =390
                    Height =270
                    Name ="Label53"
                    Caption ="cm"
                    FontName ="Tahoma"
                End
            End
        End
        Begin FormFooter
            SpecialEffect =1
            Height =0
            BackColor =-2147483633
            Name ="FormFooter"
        End
    End
End
CodeBehindForm
' See "formeleman.cls"
