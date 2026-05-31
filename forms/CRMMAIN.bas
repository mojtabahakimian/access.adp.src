Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    DefaultView =0
    AllowUpdating =4
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =15023
    RowHeight =360
    DatasheetFontHeight =9
    ItemSuffix =112
    Left =1650
    Top =780
    Right =16680
    Bottom =9450
    DatasheetGridlinesColor =12632256
    MaxRecords =10000000
    RecSrcDt = Begin
        0x0ae26492107fe240
    End
    Caption ="C   R   M"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Tahoma"
    OnLoad ="[Event Procedure]"
    OnError ="[Event Procedure]"
    Begin
        Begin Label
            BackStyle =0
            TextFontCharSet =178
            FontName ="Tahoma"
        End
        Begin CommandButton
            Width =1701
            Height =283
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Tahoma"
        End
        Begin OptionButton
            SpecialEffect =2
            LabelX =230
            LabelY =-30
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
            FontName ="Tahoma"
        End
        Begin ListBox
            SpecialEffect =2
            Width =1701
            Height =1417
            LabelX =-1701
            FontName ="Tahoma"
        End
        Begin Subform
            SpecialEffect =2
            Width =1701
            Height =1701
        End
        Begin Section
            CanGrow = NotDefault
            Height =9354
            BackColor =-2147483633
            Name ="Detail"
            OnDblClick ="[Event Procedure]"
            Begin
                Begin Subform
                    OverlapFlags =85
                    SpecialEffect =0
                    Left =116
                    Top =1299
                    Width =14885
                    Height =8055
                    TabIndex =3
                    Name ="COPMANES"
                    SourceObject ="Form.COPMANES_form"
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =11745
                    Top =345
                    Width =2114
                    Height =315
                    FontWeight =700
                    TabIndex =1
                    Name ="Text2"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =13967
                            Top =345
                            Width =1005
                            Height =240
                            FontWeight =700
                            Name ="Label3"
                            Caption ="نام  شركت :"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =12755
                    Width =1094
                    Height =315
                    FontWeight =700
                    TabIndex =2
                    Name ="Text8"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =14030
                            Width =480
                            Height =300
                            FontWeight =700
                            Name ="Label9"
                            Caption ="تلفن:"
                        End
                    End
                End
                Begin OptionGroup
                    SpecialEffect =0
                    BackStyle =1
                    OldBorderStyle =0
                    OverlapFlags =85
                    Top =690
                    Width =14931
                    Height =573
                    TabIndex =4
                    BackColor =16777164
                    Name ="hamah"
                    DefaultValue ="0"
                    OnClick ="[Event Procedure]"
                    Begin
                        Begin OptionButton
                            SpecialEffect =0
                            OverlapFlags =87
                            OldBorderStyle =0
                            ReadingOrder =1
                            Left =1339
                            Top =859
                            Width =269
                            Height =173
                            OptionValue =9
                            Name ="d9"
                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    TextAlign =3
                                    Left =78
                                    Top =806
                                    Width =1298
                                    Height =300
                                    Name ="L9"
                                    Caption ="ارتباط مرتبط"
                                End
                            End
                        End
                        Begin OptionButton
                            SpecialEffect =0
                            OverlapFlags =87
                            OldBorderStyle =0
                            ReadingOrder =1
                            Left =2986
                            Top =863
                            Width =269
                            Height =278
                            OptionValue =8
                            Name ="d8"
                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    TextAlign =3
                                    Left =1650
                                    Top =810
                                    Width =1373
                                    Height =300
                                    Name ="L8"
                                    Caption ="ارتباط مجدد"
                                End
                            End
                        End
                        Begin OptionButton
                            SpecialEffect =0
                            OverlapFlags =95
                            OldBorderStyle =0
                            ReadingOrder =1
                            Left =4539
                            Top =866
                            OptionValue =7
                            Name ="d7"
                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    TextAlign =3
                                    Left =3261
                                    Top =810
                                    Width =1313
                                    Height =240
                                    Name ="L7"
                                    Caption ="قطع ارتباط"
                                End
                            End
                        End
                        Begin OptionButton
                            SpecialEffect =0
                            OverlapFlags =95
                            OldBorderStyle =0
                            ReadingOrder =1
                            Left =6070
                            Top =871
                            OptionValue =6
                            Name ="d6"
                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    TextAlign =3
                                    Left =4791
                                    Top =810
                                    Width =1313
                                    Height =240
                                    Name ="L6"
                                    Caption ="تماس بگير"
                                End
                            End
                        End
                        Begin OptionButton
                            SpecialEffect =0
                            OverlapFlags =87
                            OldBorderStyle =0
                            ReadingOrder =1
                            Left =13853
                            Top =879
                            OptionValue =1
                            Name ="d1"
                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    TextAlign =3
                                    Left =12562
                                    Top =810
                                    Width =1313
                                    Height =240
                                    Name ="L1"
                                    Caption ="سلامتي مديران"
                                End
                            End
                        End
                        Begin OptionButton
                            SpecialEffect =0
                            OverlapFlags =87
                            OldBorderStyle =0
                            ReadingOrder =1
                            Left =14539
                            Top =882
                            OptionValue =0
                            Name ="d10"
                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    TextAlign =3
                                    Left =14175
                                    Top =810
                                    Width =390
                                    Height =240
                                    BackColor =8421504
                                    Name ="l10"
                                    Caption ="همه"
                                End
                            End
                        End
                        Begin OptionButton
                            SpecialEffect =0
                            OverlapFlags =87
                            OldBorderStyle =0
                            ReadingOrder =1
                            Left =7601
                            Top =883
                            OptionValue =5
                            Name ="d5"
                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    TextAlign =3
                                    Left =6321
                                    Top =810
                                    Width =1313
                                    Height =240
                                    Name ="l5"
                                    Caption ="تماس اونها"
                                End
                            End
                        End
                        Begin OptionButton
                            SpecialEffect =0
                            OverlapFlags =87
                            OldBorderStyle =0
                            ReadingOrder =1
                            Left =9164
                            Top =883
                            OptionValue =4
                            Name ="d4"
                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    TextAlign =3
                                    Left =7866
                                    Top =810
                                    Width =1313
                                    Height =240
                                    Name ="L4"
                                    Caption ="تماس مجدد"
                                End
                            End
                        End
                        Begin OptionButton
                            SpecialEffect =0
                            OverlapFlags =87
                            OldBorderStyle =0
                            ReadingOrder =1
                            Left =10751
                            Top =883
                            OptionValue =3
                            Name ="d3"
                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    TextAlign =3
                                    Left =9456
                                    Top =810
                                    Width =1313
                                    Height =240
                                    Name ="L3"
                                    Caption ="سلامتي ما"
                                End
                            End
                        End
                        Begin OptionButton
                            SpecialEffect =0
                            OverlapFlags =87
                            OldBorderStyle =0
                            ReadingOrder =1
                            Left =12296
                            Top =883
                            OptionValue =2
                            Name ="d2"
                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    TextAlign =3
                                    Left =11016
                                    Top =810
                                    Width =1313
                                    Height =240
                                    Name ="L2"
                                    Caption ="سلامتي پرسنل"
                                End
                            End
                        End
                    End
                End
                Begin OptionGroup
                    SpecialEffect =0
                    BackStyle =1
                    OldBorderStyle =0
                    OverlapFlags =93
                    Left =6705
                    Top =60
                    Width =4806
                    Height =548
                    BackColor =52479
                    BorderColor =4210752
                    Name ="gharar"
                    OnDblClick ="[Event Procedure]"
                    DefaultValue ="0"
                    OnClick ="[Event Procedure]"
                    Begin
                        Begin OptionButton
                            OverlapFlags =87
                            ReadingOrder =1
                            Left =11165
                            Top =195
                            OptionValue =1
                            Name ="Option98"
                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    TextAlign =3
                                    Left =10155
                                    Top =195
                                    Width =1035
                                    Height =240
                                    Name ="Label99"
                                    Caption ="قرارهاي امروز"
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =87
                            ReadingOrder =1
                            Left =6981
                            Top =240
                            Width =509
                            Height =368
                            OptionValue =0
                            Name ="Option82"
                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    TextAlign =3
                                    Left =6780
                                    Top =165
                                    Width =420
                                    Height =315
                                    Name ="Label83"
                                    Caption ="همه"
                                End
                            End
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =215
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =7545
                    Top =195
                    Width =1127
                    Height =283
                    TabIndex =5
                    Name ="datekh"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="=farsidate(Date())"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =8734
                            Top =195
                            Width =1350
                            Height =240
                            Name ="Label101"
                            Caption ="قرارها در اين تاريخ:"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =93
                    Left =3310
                    Top =120
                    Width =1025
                    Height =350
                    TabIndex =6
                    Name ="Command104"
                    Caption ="تنظيمات"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =4350
                    Top =120
                    Width =1205
                    Height =350
                    TabIndex =7
                    Name ="Command105"
                    Caption ="ليست فعاليتها"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    OverlapFlags =223
                    Left =2310
                    Top =120
                    Width =1025
                    Height =350
                    TabIndex =8
                    Name ="Command106"
                    Caption ="ارسال پيامک"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =1305
                    Top =120
                    Width =1025
                    Height =350
                    TabIndex =9
                    Name ="Command107"
                    Caption ="ياداشت"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =60
                    Top =120
                    Width =1250
                    Height =350
                    TabIndex =10
                    Name ="Command108"
                    Caption ="تقويم جلسات"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =6090
                    Top =60
                    Width =576
                    Height =576
                    TabIndex =11
                    Name ="Command111"
                    Caption ="آمار"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dadadadadadadada00000000000000000ad0c010d0c010da ,
                        0x0da0c010a0c010ad0ad0c010d0c010da0da0c010a0c010ad0ad0c010d0c010da ,
                        0x0da0c000a0c010ad0ad0c0dad0c010da0da0c0ada00010ad0ad0c0dadad010da ,
                        0x0da000adada010ad0adadadadad010da0dadadadada010ad0adadadadad000da ,
                        0xadadadadadadadad000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End
                    ObjectPalette = Begin
                        0x0003100000000000800000000080000080800000000080008000800000808000 ,
                        0x80808000c0c0c000ff000000c0c0c000ffff00000000ff00c0c0c00000ffff00 ,
                        0xffffff0000000000
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "CRMMAIN.cls"
