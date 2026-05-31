Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    DefaultView =0
    AllowUpdating =4
    ViewsAllowed =1
    TabularCharSet =178
    TabularFamily =42
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =15363
    DatasheetFontHeight =10
    ItemSuffix =104
    Left =2310
    Top =1185
    Right =17715
    Bottom =8040
    DatasheetGridlinesColor =12632256
    MaxRecords =10000000
    RecSrcDt = Begin
        0x0ae26492107fe240
    End
    Caption ="پيگيريها"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial (Arabic)"
    OnError ="[Event Procedure]"
    AllowDatasheetView =0
    Begin
        Begin Label
            BackStyle =0
            TextFontCharSet =178
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
        Begin Subform
            SpecialEffect =2
            Width =1701
            Height =1701
        End
        Begin Section
            CanGrow = NotDefault
            Height =9694
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin Subform
                    OverlapFlags =85
                    SpecialEffect =0
                    Top =1258
                    Width =15360
                    Height =7575
                    TabIndex =3
                    Name ="COPMANES"
                    SourceObject ="Form.COPMANES"
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =10950
                    Top =165
                    Width =3239
                    Height =315
                    FontWeight =700
                    TabIndex =1
                    Name ="Text2"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =14297
                            Top =165
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
                    Left =8727
                    Top =195
                    Width =1499
                    Height =315
                    FontWeight =700
                    TabIndex =2
                    Name ="Text8"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =10407
                            Top =195
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
                    Left =150
                    Top =630
                    Width =15156
                    Height =513
                    TabIndex =4
                    BackColor =16777164
                    BorderColor =4210752
                    Name ="hamah"
                    DefaultValue ="0"
                    OnClick ="[Event Procedure]"
                    Begin
                        Begin OptionButton
                            OverlapFlags =87
                            ReadingOrder =1
                            Left =1714
                            Top =799
                            Width =269
                            Height =173
                            OptionValue =9
                            Name ="Option102"
                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    TextAlign =3
                                    Left =453
                                    Top =746
                                    Width =1298
                                    Height =300
                                    Name ="L9"
                                    Caption ="روي ميز"
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =95
                            ReadingOrder =1
                            Left =3361
                            Top =803
                            Width =269
                            Height =173
                            OptionValue =8
                            Name ="Option79"
                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    TextAlign =3
                                    Left =2100
                                    Top =750
                                    Width =1298
                                    Height =300
                                    Name ="L8"
                                    Caption ="روي ميز"
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =95
                            ReadingOrder =1
                            Left =4914
                            Top =806
                            OptionValue =7
                            Name ="Option60"
                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    TextAlign =3
                                    Left =3636
                                    Top =750
                                    Width =1313
                                    Height =240
                                    Name ="L7"
                                    Caption ="قرارداد بسته شد"
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =95
                            ReadingOrder =1
                            Left =6445
                            Top =811
                            OptionValue =6
                            Name ="Option27"
                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    TextAlign =3
                                    Left =5166
                                    Top =750
                                    Width =1313
                                    Height =240
                                    Name ="L6"
                                    Caption ="راكد"
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =87
                            ReadingOrder =1
                            Left =14228
                            Top =819
                            OptionValue =1
                            Name ="Option15"
                            Begin
                                Begin Label
                                    OverlapFlags =255
                                    TextAlign =3
                                    Left =12937
                                    Top =750
                                    Width =1313
                                    Height =240
                                    Name ="L1"
                                    Caption ="شروع ارتباط"
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =87
                            ReadingOrder =1
                            Left =14914
                            Top =822
                            OptionValue =0
                            Name ="Option25"
                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    TextAlign =3
                                    Left =14550
                                    Top =750
                                    Width =390
                                    Height =240
                                    Name ="Label26"
                                    Caption ="همه"
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =95
                            ReadingOrder =1
                            Left =7976
                            Top =823
                            OptionValue =5
                            Name ="Option23"
                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    TextAlign =3
                                    Left =6696
                                    Top =750
                                    Width =1313
                                    Height =240
                                    Name ="L5"
                                    Caption ="در شرف قرارداد"
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =87
                            ReadingOrder =1
                            Left =9539
                            Top =823
                            OptionValue =4
                            Name ="Option21"
                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    TextAlign =3
                                    Left =8241
                                    Top =750
                                    Width =1313
                                    Height =240
                                    Name ="L4"
                                    Caption ="ملاقات شده"
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =95
                            ReadingOrder =1
                            Left =11126
                            Top =823
                            OptionValue =3
                            Name ="Option19"
                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    TextAlign =3
                                    Left =9831
                                    Top =750
                                    Width =1313
                                    Height =240
                                    Name ="L3"
                                    Caption ="قرار ملاقات"
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =119
                            ReadingOrder =1
                            Left =12671
                            Top =823
                            OptionValue =2
                            Name ="Option17"
                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    TextAlign =3
                                    Left =11391
                                    Top =750
                                    Width =1313
                                    Height =240
                                    Name ="L2"
                                    Caption ="درحال مذاكره"
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
                    Left =1200
                    Top =60
                    Width =5961
                    Height =533
                    BackColor =52479
                    BorderColor =4210752
                    Name ="gharar"
                    DefaultValue ="0"
                    OnClick ="[Event Procedure]"
                    Begin
                        Begin OptionButton
                            OverlapFlags =87
                            ReadingOrder =1
                            Left =2361
                            Top =195
                            Width =509
                            Height =368
                            OptionValue =0
                            Name ="Option82"
                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    TextAlign =3
                                    Left =1680
                                    Top =195
                                    Width =900
                                    Height =240
                                    Name ="Label83"
                                    Caption ="همه"
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =87
                            ReadingOrder =1
                            Left =6815
                            Top =195
                            OptionValue =1
                            Name ="Option98"
                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    TextAlign =3
                                    Left =5805
                                    Top =195
                                    Width =1035
                                    Height =240
                                    Name ="Label99"
                                    Caption ="قرارهاي امروز"
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
                    Left =3195
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
                            Left =4384
                            Top =195
                            Width =1350
                            Height =240
                            Name ="Label101"
                            Caption ="قرارها در اين تاريخ:"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "Form12.cls"
