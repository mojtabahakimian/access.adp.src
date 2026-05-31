Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    DefaultView =0
    AllowUpdating =4
    ScrollBars =2
    ViewsAllowed =1
    TabularCharSet =178
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =20
    GridY =20
    Width =11593
    DatasheetFontHeight =10
    ItemSuffix =118
    Left =5475
    Top =390
    Right =17070
    Bottom =8370
    DatasheetGridlinesColor =12632256
    MaxRecords =100000
    RecSrcDt = Begin
        0x2133dbde7221e340
    End
    Caption ="صورت مالي"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial"
    OnError ="[Event Procedure]"
    AllowDatasheetView =0
    Begin
        Begin Label
            BackStyle =0
            TextFontCharSet =178
            FontName ="Tahoma"
        End
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
            Width =850
            Height =850
        End
        Begin Line
            Width =1701
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
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            OldBorderStyle =0
            TextFontCharSet =178
            Width =1701
            LabelX =-1701
            FontName ="Tahoma"
            AsianLineBreak =255
        End
        Begin Section
            Height =9070
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =6596
                    Top =744
                    Width =1828
                    Height =283
                    Name ="MOMAEB"
                    Format ="#,###"
                    DefaultValue ="0"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =8512
                            Top =749
                            Width =2895
                            Height =240
                            Name ="Label1"
                            Caption ="موجودي مواد اوليه در ابتداي دوره:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =6596
                    Top =1082
                    Width =1828
                    Height =283
                    TabIndex =1
                    Name ="KHTEDO"
                    Format ="#,###"
                    DefaultValue ="0"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =8512
                            Top =1087
                            Width =2895
                            Height =240
                            Name ="Label3"
                            Caption ="خريد مواد اوليه طي دوره:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =6596
                    Top =1404
                    Width =1828
                    Height =283
                    TabIndex =2
                    BackColor =13434828
                    Name ="MAVAMMA"
                    ControlSource ="=[MOMAEB]+[KHTEDO]"
                    Format ="#,###"
                    DefaultValue ="0"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =8512
                            Top =1409
                            Width =2895
                            Height =240
                            Name ="Label5"
                            Caption ="مواد اوليه آماده براي مصرف:"
                        End
                    End
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =1
                    Left =8512
                    Top =1704
                    Width =2895
                    Height =240
                    Name ="Label6"
                    Caption ="كسر مي شود:"
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =6596
                    Top =1989
                    Width =1828
                    Height =283
                    TabIndex =3
                    Name ="MOAVPA"
                    Format ="#,###"
                    DefaultValue ="0"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =8512
                            Top =1994
                            Width =2895
                            Height =240
                            Name ="Label8"
                            Caption ="موجودي مواد اوليه در پايان دوره:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =6596
                    Top =2334
                    Width =1828
                    Height =283
                    TabIndex =4
                    BackColor =13434828
                    Name ="MAMATO"
                    ControlSource ="=[MOMAEB]+[KHTEDO]-[MOAVPA]"
                    Format ="#,###"
                    DefaultValue ="0"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =8512
                            Top =2339
                            Width =2895
                            Height =240
                            Name ="Label10"
                            Caption ="مواداوليه مصرف شده در توليد:"
                        End
                    End
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =1
                    Left =8512
                    Top =2649
                    Width =2895
                    Height =240
                    Name ="Label11"
                    Caption ="اضافه مي شود:"
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =6596
                    Top =2949
                    Width =1828
                    Height =283
                    TabIndex =5
                    Name ="DASTO"
                    Format ="#,###"
                    DefaultValue ="0"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =8512
                            Top =2954
                            Width =2895
                            Height =240
                            Name ="Label13"
                            Caption ="دستمرد مستقيم:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =6596
                    Top =3279
                    Width =1828
                    Height =283
                    TabIndex =6
                    Name ="SAHAZTO"
                    Format ="#,###"
                    DefaultValue ="0"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =8512
                            Top =3281
                            Width =2895
                            Height =240
                            Name ="Label15"
                            Caption ="ساير هزينه هاي توليد(سربار توليد):"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =6596
                    Top =3624
                    Width =1828
                    Height =283
                    TabIndex =7
                    Name ="MOKADJAVDO"
                    Format ="#,###"
                    DefaultValue ="0"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =8512
                            Top =3628
                            Width =2895
                            Height =240
                            Name ="Label17"
                            Caption ="موجودي كالاي درجريان ساخت اول دوره:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =6596
                    Top =3969
                    Width =1828
                    Height =255
                    TabIndex =8
                    BackColor =13434828
                    Name ="JAM1"
                    ControlSource ="=[DASTO]+[SAHAZTO]+[MOKADJAVDO]+[MAMATO]"
                    Format ="#,###"
                    DefaultValue ="0"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =8512
                            Top =3973
                            Width =2895
                            Height =240
                            Name ="Label19"
                            Caption ="جمع:"
                        End
                    End
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =1
                    Left =8512
                    Top =4254
                    Width =2895
                    Height =240
                    Name ="Label21"
                    Caption ="كسر مي شود:"
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =6596
                    Top =4537
                    Width =1828
                    Height =283
                    TabIndex =9
                    Name ="MODJSAPA"
                    Format ="#,###"
                    DefaultValue ="0"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =8512
                            Top =4537
                            Width =2895
                            Height =240
                            Name ="Label23"
                            Caption ="موجودي كالاي در جريان ساخت پايان دوره:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =6596
                    Top =4869
                    Width =1828
                    Height =283
                    TabIndex =10
                    BackColor =16776960
                    Name ="BAHATSAKHT"
                    ControlSource ="=[JAM1]-[MODJSAPA]"
                    Format ="#,###"
                    DefaultValue ="0"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =8512
                            Top =4869
                            Width =2895
                            Height =240
                            Name ="Label25"
                            Caption ="بهاي تمام شده كالاي ساخته شده:"
                        End
                    End
                End
                Begin Rectangle
                    OverlapFlags =255
                    Left =6435
                    Top =504
                    Width =5014
                    Height =4794
                    Name ="Box26"
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =247
                    TextFontFamily =2
                    Left =7680
                    Top =195
                    Width =2790
                    Height =480
                    FontSize =12
                    FontWeight =700
                    BackColor =-2147483633
                    Name ="Label27"
                    Caption ="بهاي تمام شده كالاي ساخته شده"
                    FontName ="Titr"
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =6617
                    Top =5958
                    Width =1828
                    Height =283
                    TabIndex =11
                    Name ="MOKALAV"
                    Format ="#,###"
                    DefaultValue ="0"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =8638
                            Top =5963
                            Width =2790
                            Height =240
                            Name ="Label45"
                            Caption ="موجودي كالا در ابتداي دوره:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =6617
                    Top =6296
                    Width =1828
                    Height =283
                    TabIndex =12
                    Name ="KHTYDO"
                    Format ="#,###"
                    DefaultValue ="0"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =8638
                            Top =6301
                            Width =2790
                            Height =240
                            Name ="Label47"
                            Caption ="خريد  كالا طي دوره:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =6617
                    Top =6618
                    Width =1828
                    Height =283
                    TabIndex =13
                    BackColor =13434828
                    Name ="BAHASAKHT"
                    ControlSource ="=[JAM1]-[MODJSAPA]"
                    Format ="#,###"
                    DefaultValue ="0"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =8638
                            Top =6623
                            Width =2790
                            Height =240
                            Name ="Label49"
                            Caption ="بهاي تمام شده كالاي ساخته شده :"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =6621
                    Top =6969
                    Width =1828
                    Height =283
                    TabIndex =14
                    Name ="SAYER"
                    Format ="#,###"
                    DefaultValue ="0"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =8642
                            Top =6974
                            Width =2790
                            Height =240
                            Name ="Label52"
                            Caption ="ساير:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =6636
                    Top =7299
                    Width =1828
                    Height =283
                    TabIndex =15
                    BackColor =13434828
                    Name ="JAM2"
                    ControlSource ="=[MOKALAV]+[KHTYDO]+[BAHASAKHT]+[SAYER]"
                    Format ="#,###"
                    DefaultValue ="0"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =8657
                            Top =7304
                            Width =2790
                            Height =240
                            Name ="Label54"
                            Caption ="جمع:"
                        End
                    End
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =1
                    Left =8646
                    Top =7614
                    Width =2790
                    Height =240
                    Name ="Label55"
                    Caption ="كسر مي شود:"
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =6636
                    Top =8229
                    Width =1828
                    Height =283
                    TabIndex =17
                    BackColor =16777164
                    Name ="BAHATAFR"
                    ControlSource ="=[MOKALAV]+[KHTYDO]+[BAHASAKHT]+[SAYER]-[MOKALPA]"
                    Format ="#,###"
                    DefaultValue ="0"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =8657
                            Top =8234
                            Width =2790
                            Height =240
                            Name ="Label57"
                            Caption ="بهاي تمام شده كالاي فروش رفته:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =6636
                    Top =7899
                    Width =1828
                    Height =283
                    TabIndex =16
                    Name ="MOKALPA"
                    Format ="#,###"
                    DefaultValue ="0"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =8650
                            Top =7905
                            Width =2850
                            Height =240
                            Name ="Label61"
                            Caption ="موجودي كالا در پايان دوره:"
                        End
                    End
                End
                Begin Rectangle
                    OverlapFlags =255
                    Left =6441
                    Top =5718
                    Width =5029
                    Height =2964
                    Name ="Box69"
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =247
                    TextFontFamily =2
                    Left =7716
                    Top =5349
                    Width =2820
                    Height =480
                    FontSize =12
                    FontWeight =700
                    BackColor =-2147483633
                    Name ="Label70"
                    Caption ="بهاي تمام شده كالاي فروش رفته"
                    FontName ="Titr"
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =296
                    Top =735
                    Width =1828
                    Height =283
                    TabIndex =18
                    Name ="FRNAKH"
                    Format ="#,###"
                    DefaultValue ="0"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =2212
                            Top =740
                            Width =2895
                            Height =240
                            Name ="Label74"
                            Caption ="فروش ناخالص:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =296
                    Top =1073
                    Width =1828
                    Height =283
                    TabIndex =19
                    Name ="BRFRT"
                    Format ="#,###"
                    DefaultValue ="0"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =2205
                            Top =1071
                            Width =3045
                            Height =240
                            Name ="Label76"
                            Caption ="كسر مي شود برگشت از فروش و تخفيفات:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =296
                    Top =1395
                    Width =1828
                    Height =283
                    TabIndex =20
                    Name ="FRKH"
                    ControlSource ="=[FRNAKH]-[BRFRT]"
                    Format ="#,###"
                    DefaultValue ="0"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =2212
                            Top =1400
                            Width =2895
                            Height =240
                            Name ="Label78"
                            Caption ="فروش خالص:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =296
                    Top =1755
                    Width =1828
                    Height =283
                    TabIndex =21
                    Name ="BAHATAKAFR"
                    ControlSource ="=[MOKALAV]+[KHTYDO]+[BAHASAKHT]+[SAYER]-[MOKALPA]"
                    Format ="#,###"
                    DefaultValue ="0"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =2205
                            Top =1761
                            Width =3435
                            Height =240
                            Name ="Label81"
                            Caption ="كسر مي شود بهاي تمام شده كالاي فروش رفته:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =296
                    Top =2100
                    Width =1828
                    Height =283
                    TabIndex =22
                    BackColor =13434828
                    Name ="SUDNAV"
                    ControlSource ="=[FRKH]-[BAHATAKAFR]"
                    Format ="#,###"
                    DefaultValue ="0"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =2212
                            Top =2105
                            Width =2895
                            Height =240
                            Name ="Label83"
                            Caption ="سود (زيان) ناويژه:"
                        End
                    End
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =1
                    Left =2212
                    Top =2415
                    Width =2895
                    Height =240
                    Name ="Label84"
                    Caption ="اضافه مي شود ساير درآمدها:"
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =296
                    Top =2715
                    Width =1828
                    Height =283
                    TabIndex =23
                    Name ="DRHAKHA"
                    Format ="#,###"
                    DefaultValue ="0"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =2212
                            Top =2720
                            Width =2895
                            Height =240
                            Name ="Label86"
                            Caption ="درآمد حاصل از خدمات:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =296
                    Top =3045
                    Width =1828
                    Height =283
                    TabIndex =24
                    Name ="DRMOBA"
                    Format ="#,###"
                    DefaultValue ="0"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =2205
                            Top =3051
                            Width =3285
                            Height =240
                            Name ="Label88"
                            Caption ="درآمد حاصل از معاملات بانكي و سرمايه گذاري:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =296
                    Top =3390
                    Width =1828
                    Height =283
                    TabIndex =25
                    Name ="DREJA"
                    Format ="#,###"
                    DefaultValue ="0"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =2212
                            Top =3394
                            Width =2895
                            Height =240
                            Name ="Label90"
                            Caption ="درآمد حاصل از اجاره:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =296
                    Top =3735
                    Width =1828
                    Height =283
                    TabIndex =26
                    Name ="DRGMOTAR"
                    Format ="#,###"
                    DefaultValue ="0"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =2212
                            Top =3739
                            Width =2895
                            Height =240
                            Name ="Label92"
                            Caption ="درآمدهاي غير مترقبه:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =296
                    Top =4095
                    Width =1828
                    Height =283
                    TabIndex =27
                    Name ="SAYERDA"
                    Format ="#,###"
                    DefaultValue ="0"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =2212
                            Top =4095
                            Width =2895
                            Height =240
                            Name ="Label95"
                            Caption ="ساير دريافتي ها و درآمدها:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =296
                    Top =4427
                    Width =1828
                    Height =283
                    TabIndex =28
                    BackColor =16777164
                    Name ="JAMDRSUNV"
                    ControlSource ="=[DRHAKHA]+[DRMOBA]+[DREJA]+[DRGMOTAR]+[SAYERDA]+[SUDNAV]"
                    Format ="#,###"
                    DefaultValue ="0"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =2212
                            Top =4427
                            Width =2895
                            Height =240
                            Name ="Label97"
                            Caption ="جمع درآمدها و سود (ناويژه):"
                        End
                    End
                End
                Begin Rectangle
                    OverlapFlags =255
                    Left =195
                    Top =504
                    Width =5539
                    Height =7464
                    Name ="Box98"
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =247
                    TextFontFamily =2
                    Left =1560
                    Top =195
                    Width =2295
                    Height =480
                    FontSize =12
                    FontWeight =700
                    BackColor =-2147483633
                    Name ="Label99"
                    Caption ="حساب عملكرد سود و زيان "
                    FontName ="Titr"
                End
                Begin TextBox
                    OverlapFlags =247
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =285
                    Top =5108
                    Width =1828
                    Height =283
                    TabIndex =29
                    Name ="HOGH"
                    Format ="#,###"
                    DefaultValue ="0"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextAlign =1
                            Left =2201
                            Top =5108
                            Width =2895
                            Height =240
                            Name ="Label101"
                            Caption ="حقوق و دستمزد:"
                        End
                    End
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =1
                    Left =2220
                    Top =4793
                    Width =2895
                    Height =240
                    Name ="Label102"
                    Caption ="كسر مي شود هزينه ها:"
                End
                Begin TextBox
                    OverlapFlags =247
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =285
                    Top =5445
                    Width =1828
                    Height =283
                    TabIndex =30
                    Name ="HAZED"
                    Format ="#,###"
                    DefaultValue ="0"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextAlign =1
                            Left =2201
                            Top =5445
                            Width =2895
                            Height =240
                            Name ="Label104"
                            Caption ="هزينه هاي اداري:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =247
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =285
                    Top =5790
                    Width =1828
                    Height =283
                    TabIndex =31
                    Name ="HAZMAL"
                    Format ="#,###"
                    DefaultValue ="0"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextAlign =1
                            Left =2201
                            Top =5790
                            Width =2895
                            Height =240
                            Name ="Label106"
                            Caption ="هزينه هاي مالي:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =247
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =285
                    Top =6150
                    Width =1828
                    Height =283
                    TabIndex =32
                    Name ="HAZTOZI"
                    Format ="#,###"
                    DefaultValue ="0"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextAlign =1
                            Left =2201
                            Top =6150
                            Width =2895
                            Height =240
                            Name ="Label108"
                            Caption ="هزينه هاي توزيع و فروش:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =247
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =285
                    Top =6495
                    Width =1828
                    Height =283
                    TabIndex =33
                    Name ="HAZEST"
                    Format ="#,###"
                    DefaultValue ="0"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextAlign =1
                            Left =2201
                            Top =6495
                            Width =2895
                            Height =240
                            Name ="Label110"
                            Caption ="هزينه هاي استهلاك(غير توليدي):"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =247
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =285
                    Top =6840
                    Width =1828
                    Height =283
                    TabIndex =34
                    Name ="SAYERHAZ"
                    Format ="#,###"
                    DefaultValue ="0"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextAlign =1
                            Left =2201
                            Top =6840
                            Width =2895
                            Height =240
                            Name ="Label112"
                            Caption ="ساير هزينه ها:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =247
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =285
                    Top =7185
                    Width =1828
                    Height =283
                    TabIndex =35
                    BackColor =16777164
                    Name ="JHAZ"
                    ControlSource ="=[HOGH]+[HAZED]+[HAZMAL]+[HAZTOZI]+[HAZEST]+[SAYERHAZ]"
                    Format ="#,###"
                    DefaultValue ="0"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextAlign =1
                            Left =2201
                            Top =7185
                            Width =2895
                            Height =240
                            Name ="Label114"
                            Caption ="جمع هزينه ها:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =247
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =285
                    Top =7530
                    Width =1828
                    Height =283
                    TabIndex =36
                    BackColor =13434828
                    Name ="SUD"
                    ControlSource ="=[JAMDRSUNV]-[JHAZ]"
                    Format ="#,###"
                    DefaultValue ="0"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextAlign =1
                            Left =2201
                            Top =7530
                            Width =2895
                            Height =240
                            Name ="Label116"
                            Caption ="سود و زيان ويژه دوره عملكرد:"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =170
                    Top =8192
                    Height =536
                    TabIndex =37
                    Name ="Command117"
                    Caption ="چاپ صورت مالي"
                    OnClick ="[Event Procedure]"
                End
            End
        End
    End
End
CodeBehindForm
' See "SURAT_MALI_MAIN.cls"
