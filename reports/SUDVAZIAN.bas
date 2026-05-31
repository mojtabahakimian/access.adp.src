Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    AllowUpdating =4
    TabularCharSet =178
    TabularFamily =0
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =10714
    DatasheetFontHeight =10
    ItemSuffix =1
    Left =270
    Top =1380
    DatasheetGridlinesColor =12632256
    MaxRecords =100000
    RecSrcDt = Begin
        0x2eaa31915123e340
    End
    Caption ="صورتهاي مالي"
    DatasheetFontName ="Arial"
    Begin
        Begin Label
            BackStyle =0
            TextFontFamily =2
            FontName ="Arial"
        End
        Begin Rectangle
            Width =850
            Height =850
        End
        Begin TextBox
            FELineBreak = NotDefault
            OldBorderStyle =0
            TextFontFamily =2
            Width =1701
            LabelX =-1701
            FontName ="Arial"
            AsianLineBreak =255
        End
        Begin PageHeader
            Height =540
            Name ="PageHeaderSection"
            Begin
                Begin Label
                    TextFontCharSet =178
                    TextAlign =2
                    Left =4734
                    Width =1515
                    Height =540
                    FontSize =14
                    BackColor =-2147483633
                    Name ="Label0"
                    Caption ="صورتهاي مالي"
                    FontName ="WeblogmaYekan"
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =14626
            Name ="Detail"
            Begin
                Begin TextBox
                    OldBorderStyle =1
                    TextFontCharSet =178
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =5775
                    Top =630
                    Width =1828
                    Height =345
                    FontSize =10
                    Name ="MOMAEB"
                    ControlSource ="=[Forms]![SURAT_MALI_MAIN]![MOMAEB]"
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextFontCharSet =178
                            TextAlign =1
                            Left =7691
                            Top =635
                            Width =2895
                            Height =240
                            Name ="Label1"
                            Caption ="موجودي مواد اوليه در ابتداي دوره:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =1
                    TextFontCharSet =178
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =5790
                    Top =1020
                    Width =1828
                    Height =345
                    FontSize =10
                    TabIndex =1
                    Name ="KHTEDO"
                    ControlSource ="=[Forms]![SURAT_MALI_MAIN]![KHTEDO]"
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextFontCharSet =178
                            TextAlign =1
                            Left =7706
                            Top =1025
                            Width =2895
                            Height =240
                            Name ="Label3"
                            Caption ="خريد مواد اوليه طي دوره:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =1
                    TextFontCharSet =178
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    BackStyle =0
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =5790
                    Top =1425
                    Width =1828
                    Height =345
                    FontSize =10
                    TabIndex =2
                    BackColor =13434828
                    Name ="MAVAMMA"
                    ControlSource ="=[Forms]![SURAT_MALI_MAIN]![MAVAMMA]"
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextFontCharSet =178
                            TextAlign =1
                            Left =7706
                            Top =1430
                            Width =2895
                            Height =240
                            Name ="Label5"
                            Caption ="مواد اوليه آماده براي مصرف:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin Label
                    TextFontCharSet =178
                    TextAlign =1
                    Left =7706
                    Top =1725
                    Width =2895
                    Height =240
                    Name ="Label6"
                    Caption ="كسر مي شود:"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    OldBorderStyle =1
                    TextFontCharSet =178
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =5790
                    Top =2010
                    Width =1828
                    Height =345
                    FontSize =10
                    TabIndex =3
                    Name ="MOAVPA"
                    ControlSource ="=[Forms]![SURAT_MALI_MAIN]![MOAVPA]"
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextFontCharSet =178
                            TextAlign =1
                            Left =7706
                            Top =2015
                            Width =2895
                            Height =240
                            Name ="Label8"
                            Caption ="موجودي مواد اوليه در پايان دوره:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =1
                    TextFontCharSet =178
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    BackStyle =0
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =5790
                    Top =2430
                    Width =1828
                    Height =345
                    FontSize =10
                    TabIndex =4
                    BackColor =13434828
                    Name ="MAMATO"
                    ControlSource ="=[Forms]![SURAT_MALI_MAIN]![MAMATO]"
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextFontCharSet =178
                            TextAlign =1
                            Left =7706
                            Top =2435
                            Width =2895
                            Height =240
                            Name ="Label10"
                            Caption ="مواداوليه مصرف شده در توليد:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin Label
                    TextFontCharSet =178
                    TextAlign =1
                    Left =7706
                    Top =2745
                    Width =2895
                    Height =240
                    Name ="Label11"
                    Caption ="اضافه مي شود:"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    OldBorderStyle =1
                    TextFontCharSet =178
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =5790
                    Top =3045
                    Width =1828
                    Height =345
                    FontSize =10
                    TabIndex =5
                    Name ="DASTO"
                    ControlSource ="=[Forms]![SURAT_MALI_MAIN]![DASTO]"
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextFontCharSet =178
                            TextAlign =1
                            Left =7706
                            Top =3050
                            Width =2895
                            Height =240
                            Name ="Label13"
                            Caption ="دستمرد مستقيم:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =1
                    TextFontCharSet =178
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =5790
                    Top =3465
                    Width =1828
                    Height =345
                    FontSize =10
                    TabIndex =6
                    Name ="SAHAZTO"
                    ControlSource ="=[Forms]![SURAT_MALI_MAIN]![SAHAZTO]"
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextFontCharSet =178
                            TextAlign =1
                            Left =7706
                            Top =3467
                            Width =2895
                            Height =240
                            Name ="Label15"
                            Caption ="ساير هزينه هاي توليد(سربار توليد):"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =1
                    TextFontCharSet =178
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =5790
                    Top =3855
                    Width =1828
                    Height =345
                    FontSize =10
                    TabIndex =7
                    Name ="MOKADJAVDO"
                    ControlSource ="=[Forms]![SURAT_MALI_MAIN]![MOKADJAVDO]"
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextFontCharSet =178
                            TextAlign =1
                            Left =7706
                            Top =3859
                            Width =2895
                            Height =240
                            Name ="Label17"
                            Caption ="موجودي كالاي درجريان ساخت اول دوره:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =1
                    TextFontCharSet =178
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    BackStyle =0
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =5790
                    Top =4260
                    Width =1828
                    Height =345
                    FontSize =10
                    TabIndex =8
                    BackColor =13434828
                    Name ="JAM1"
                    ControlSource ="=[Forms]![SURAT_MALI_MAIN]![JAM1]"
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextFontCharSet =178
                            TextAlign =1
                            Left =7706
                            Top =4264
                            Width =2895
                            Height =240
                            Name ="Label19"
                            Caption ="جمع:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin Label
                    TextFontCharSet =178
                    TextAlign =1
                    Left =7706
                    Top =4545
                    Width =2895
                    Height =240
                    Name ="Label21"
                    Caption ="كسر مي شود:"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    OldBorderStyle =1
                    TextFontCharSet =178
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =5790
                    Top =4828
                    Width =1828
                    Height =345
                    FontSize =10
                    TabIndex =9
                    Name ="MODJSAPA"
                    ControlSource ="=[Forms]![SURAT_MALI_MAIN]![MODJSAPA]"
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextFontCharSet =178
                            TextAlign =1
                            Left =7706
                            Top =4828
                            Width =2895
                            Height =240
                            Name ="Label23"
                            Caption ="موجودي كالاي در جريان ساخت پايان دوره:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =1
                    TextFontCharSet =178
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =5790
                    Top =5220
                    Width =1828
                    Height =345
                    FontSize =10
                    TabIndex =10
                    BackColor =16776960
                    Name ="BAHATSAKHT"
                    ControlSource ="=[Forms]![SURAT_MALI_MAIN]![BAHATSAKHT]"
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextFontCharSet =178
                            TextAlign =1
                            Left =7706
                            Top =5220
                            Width =2895
                            Height =240
                            Name ="Label25"
                            Caption ="بهاي تمام شده كالاي ساخته شده:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin Rectangle
                    BackStyle =0
                    Left =5595
                    Top =426
                    Width =4999
                    Height =5304
                    Name ="Box26"
                End
                Begin Label
                    BackStyle =1
                    TextFontCharSet =178
                    Left =6570
                    Top =139
                    Width =2625
                    Height =405
                    FontSize =12
                    Name ="Label27"
                    Caption ="بهاي تمام شده كالاي ساخته شده"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    OldBorderStyle =1
                    TextFontCharSet =178
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =252
                    Top =665
                    Width =1828
                    Height =345
                    FontSize =10
                    TabIndex =11
                    Name ="MOKALAV"
                    ControlSource ="=[Forms]![SURAT_MALI_MAIN]![MOKALAV]"
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextFontCharSet =178
                            TextAlign =1
                            Left =2272
                            Top =670
                            Width =2790
                            Height =240
                            Name ="Label45"
                            Caption ="موجودي كالا در ابتداي دوره:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =1
                    TextFontCharSet =178
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =255
                    Top =1080
                    Width =1828
                    Height =345
                    FontSize =10
                    TabIndex =12
                    Name ="KHTYDO"
                    ControlSource ="=[Forms]![SURAT_MALI_MAIN]![KHTYDO]"
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextFontCharSet =178
                            TextAlign =1
                            Left =2275
                            Top =1085
                            Width =2790
                            Height =240
                            Name ="Label47"
                            Caption ="خريد  كالا طي دوره:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =1
                    TextFontCharSet =178
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    BackStyle =0
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =270
                    Top =1485
                    Width =1828
                    Height =345
                    FontSize =10
                    TabIndex =13
                    BackColor =13434828
                    Name ="BAHASAKHT"
                    ControlSource ="=[Forms]![SURAT_MALI_MAIN]![BAHASAKHT]"
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextFontCharSet =178
                            TextAlign =1
                            Left =2290
                            Top =1490
                            Width =2790
                            Height =240
                            Name ="Label49"
                            Caption ="بهاي تمام شده كالاي ساخته شده :"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =1
                    TextFontCharSet =178
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =285
                    Top =1920
                    Width =1828
                    Height =345
                    FontSize =10
                    TabIndex =14
                    Name ="SAYER"
                    ControlSource ="=[Forms]![SURAT_MALI_MAIN]![SAYER]"
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextFontCharSet =178
                            TextAlign =1
                            Left =2309
                            Top =1925
                            Width =2790
                            Height =240
                            Name ="Label52"
                            Caption ="ساير:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =1
                    TextFontCharSet =178
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    BackStyle =0
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =285
                    Top =2325
                    Width =1828
                    Height =345
                    FontSize =10
                    TabIndex =15
                    BackColor =13434828
                    Name ="JAM2"
                    ControlSource ="=[Forms]![SURAT_MALI_MAIN]![JAM2]"
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextFontCharSet =178
                            TextAlign =1
                            Left =2324
                            Top =2330
                            Width =2790
                            Height =240
                            Name ="Label54"
                            Caption ="جمع:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin Label
                    TextFontCharSet =178
                    TextAlign =1
                    Left =2313
                    Top =2640
                    Width =2790
                    Height =240
                    Name ="Label55"
                    Caption ="كسر مي شود:"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    OldBorderStyle =1
                    TextFontCharSet =178
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =285
                    Top =3345
                    Width =1828
                    Height =345
                    FontSize =10
                    TabIndex =16
                    BackColor =16777164
                    Name ="BAHATAFR"
                    ControlSource ="=[Forms]![SURAT_MALI_MAIN]![BAHATAFR]"
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextFontCharSet =178
                            TextAlign =1
                            Left =2324
                            Top =3350
                            Width =2790
                            Height =240
                            Name ="Label57"
                            Caption ="بهاي تمام شده كالاي فروش رفته:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =1
                    TextFontCharSet =178
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =285
                    Top =2925
                    Width =1828
                    Height =345
                    FontSize =10
                    TabIndex =17
                    Name ="MOKALPA"
                    ControlSource ="=[Forms]![SURAT_MALI_MAIN]![MOKALPA]"
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextFontCharSet =178
                            TextAlign =1
                            Left =2317
                            Top =2931
                            Width =2850
                            Height =240
                            Name ="Label61"
                            Caption ="موجودي كالا در پايان دوره:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin Rectangle
                    BackStyle =0
                    Left =75
                    Top =425
                    Width =5029
                    Height =3489
                    Name ="Box69"
                End
                Begin Label
                    BackStyle =1
                    TextFontCharSet =178
                    Left =1080
                    Top =150
                    Width =3240
                    Height =405
                    FontSize =12
                    Name ="Label70"
                    Caption ="بهاي تمام شده كالاي فروش رفته"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    OldBorderStyle =1
                    TextFontCharSet =178
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2895
                    Top =6360
                    Width =1828
                    Height =345
                    FontSize =10
                    TabIndex =18
                    Name ="FRNAKH"
                    ControlSource ="=[Forms]![SURAT_MALI_MAIN]![FRNAKH]"
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextFontCharSet =178
                            TextAlign =1
                            Left =4811
                            Top =6365
                            Width =2895
                            Height =240
                            Name ="Label74"
                            Caption ="فروش ناخالص:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =1
                    TextFontCharSet =178
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2891
                    Top =6752
                    Width =1828
                    Height =345
                    FontSize =10
                    TabIndex =19
                    Name ="BRFRT"
                    ControlSource ="=[Forms]![SURAT_MALI_MAIN]![BRFRT]"
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextFontCharSet =178
                            TextAlign =1
                            Left =4800
                            Top =6750
                            Width =3045
                            Height =240
                            Name ="Label76"
                            Caption ="كسر مي شود برگشت از فروش و تخفيفات:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =1
                    TextFontCharSet =178
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2891
                    Top =7155
                    Width =1828
                    Height =345
                    FontSize =10
                    TabIndex =20
                    Name ="FRKH"
                    ControlSource ="=[Forms]![SURAT_MALI_MAIN]![FRKH]"
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextFontCharSet =178
                            TextAlign =1
                            Left =4807
                            Top =7160
                            Width =2895
                            Height =240
                            Name ="Label78"
                            Caption ="فروش خالص:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =1
                    TextFontCharSet =178
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2891
                    Top =7590
                    Width =1828
                    Height =345
                    FontSize =10
                    TabIndex =21
                    Name ="BAHATAKAFR"
                    ControlSource ="=[Forms]![SURAT_MALI_MAIN]![BAHATAKAFR]"
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextFontCharSet =178
                            TextAlign =1
                            Left =4800
                            Top =7596
                            Width =3435
                            Height =240
                            Name ="Label81"
                            Caption ="كسر مي شود بهاي تمام شده كالاي فروش رفته:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =1
                    TextFontCharSet =178
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2891
                    Top =7995
                    Width =1828
                    Height =345
                    FontSize =10
                    TabIndex =22
                    BackColor =13434828
                    Name ="SUDNAV"
                    ControlSource ="=[Forms]![SURAT_MALI_MAIN]![SUDNAV]"
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextFontCharSet =178
                            TextAlign =1
                            Left =4807
                            Top =8000
                            Width =2895
                            Height =240
                            Name ="Label83"
                            Caption ="سود (زيان) ناويژه:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin Label
                    TextFontCharSet =178
                    TextAlign =1
                    Left =4807
                    Top =8310
                    Width =2895
                    Height =240
                    Name ="Label84"
                    Caption ="اضافه مي شود ساير درآمدها:"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    OldBorderStyle =1
                    TextFontCharSet =178
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2891
                    Top =8610
                    Width =1828
                    Height =345
                    FontSize =10
                    TabIndex =23
                    Name ="DRHAKHA"
                    ControlSource ="=[Forms]![SURAT_MALI_MAIN]![DRHAKHA]"
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextFontCharSet =178
                            TextAlign =1
                            Left =4807
                            Top =8615
                            Width =2895
                            Height =240
                            Name ="Label86"
                            Caption ="درآمد حاصل از خدمات:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =1
                    TextFontCharSet =178
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2891
                    Top =9015
                    Width =1828
                    Height =345
                    FontSize =10
                    TabIndex =24
                    Name ="DRMOBA"
                    ControlSource ="=[Forms]![SURAT_MALI_MAIN]![DRMOBA]"
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextFontCharSet =178
                            TextAlign =1
                            Left =4800
                            Top =9021
                            Width =3285
                            Height =240
                            Name ="Label88"
                            Caption ="درآمد حاصل از معاملات بانكي و سرمايه گذاري:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =1
                    TextFontCharSet =178
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2891
                    Top =9420
                    Width =1828
                    Height =345
                    FontSize =10
                    TabIndex =25
                    Name ="DREJA"
                    ControlSource ="=[Forms]![SURAT_MALI_MAIN]![DREJA]"
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextFontCharSet =178
                            TextAlign =1
                            Left =4807
                            Top =9424
                            Width =2895
                            Height =240
                            Name ="Label90"
                            Caption ="درآمد حاصل از اجاره:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =1
                    TextFontCharSet =178
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2891
                    Top =9810
                    Width =1828
                    Height =345
                    FontSize =10
                    TabIndex =26
                    Name ="DRGMOTAR"
                    ControlSource ="=[Forms]![SURAT_MALI_MAIN]![DRGMOTAR]"
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextFontCharSet =178
                            TextAlign =1
                            Left =4807
                            Top =9814
                            Width =2895
                            Height =240
                            Name ="Label92"
                            Caption ="درآمدهاي غير مترقبه:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =1
                    TextFontCharSet =178
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2891
                    Top =10266
                    Width =1828
                    Height =345
                    FontSize =10
                    TabIndex =27
                    Name ="SAYERDA"
                    ControlSource ="=[Forms]![SURAT_MALI_MAIN]![SAYERDA]"
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextFontCharSet =178
                            TextAlign =1
                            Left =4807
                            Top =10266
                            Width =2895
                            Height =240
                            Name ="Label95"
                            Caption ="ساير دريافتي ها و درآمدها:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =1
                    TextFontCharSet =178
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2891
                    Top =10650
                    Width =1828
                    Height =345
                    FontSize =10
                    TabIndex =28
                    BackColor =16777164
                    Name ="JAMDRSUNV"
                    ControlSource ="=[Forms]![SURAT_MALI_MAIN]![JAMDRSUNV]"
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextFontCharSet =178
                            TextAlign =1
                            Left =4807
                            Top =10650
                            Width =2895
                            Height =240
                            Name ="Label97"
                            Caption ="جمع درآمدها و سود (ناويژه):"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin Rectangle
                    BackStyle =0
                    Left =1080
                    Top =5955
                    Width =8794
                    Height =8664
                    Name ="Box98"
                End
                Begin Label
                    BackStyle =1
                    TextFontCharSet =178
                    Left =4200
                    Top =5790
                    Width =2280
                    Height =405
                    FontSize =12
                    Name ="Label99"
                    Caption ="حساب عملكرد سود و زيان "
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    OldBorderStyle =1
                    TextFontCharSet =178
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2880
                    Top =11331
                    Width =1828
                    Height =345
                    FontSize =10
                    TabIndex =29
                    Name ="HOGH"
                    ControlSource ="=[Forms]![SURAT_MALI_MAIN]![HOGH]"
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextFontCharSet =178
                            TextAlign =1
                            Left =4796
                            Top =11331
                            Width =2895
                            Height =240
                            Name ="Label101"
                            Caption ="حقوق و دستمزد:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin Label
                    TextFontCharSet =178
                    TextAlign =1
                    Left =4815
                    Top =11016
                    Width =2895
                    Height =240
                    Name ="Label102"
                    Caption ="كسر مي شود هزينه ها:"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    OldBorderStyle =1
                    TextFontCharSet =178
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2880
                    Top =11668
                    Width =1828
                    Height =345
                    FontSize =10
                    TabIndex =30
                    Name ="HAZED"
                    ControlSource ="=[Forms]![SURAT_MALI_MAIN]![HAZED]"
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextFontCharSet =178
                            TextAlign =1
                            Left =4796
                            Top =11668
                            Width =2895
                            Height =240
                            Name ="Label104"
                            Caption ="هزينه هاي اداري:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =1
                    TextFontCharSet =178
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2880
                    Top =12058
                    Width =1828
                    Height =345
                    FontSize =10
                    TabIndex =31
                    Name ="HAZMAL"
                    ControlSource ="=[Forms]![SURAT_MALI_MAIN]![HAZMAL]"
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextFontCharSet =178
                            TextAlign =1
                            Left =4796
                            Top =12058
                            Width =2895
                            Height =240
                            Name ="Label106"
                            Caption ="هزينه هاي مالي:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =1
                    TextFontCharSet =178
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2880
                    Top =12478
                    Width =1828
                    Height =345
                    FontSize =10
                    TabIndex =32
                    Name ="HAZTOZI"
                    ControlSource ="=[Forms]![SURAT_MALI_MAIN]![HAZTOZI]"
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextFontCharSet =178
                            TextAlign =1
                            Left =4796
                            Top =12478
                            Width =2895
                            Height =240
                            Name ="Label108"
                            Caption ="هزينه هاي توزيع و فروش:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =1
                    TextFontCharSet =178
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2880
                    Top =12883
                    Width =1828
                    Height =345
                    FontSize =10
                    TabIndex =33
                    Name ="HAZEST"
                    ControlSource ="=[Forms]![SURAT_MALI_MAIN]![HAZEST]"
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextFontCharSet =178
                            TextAlign =1
                            Left =4796
                            Top =12883
                            Width =2895
                            Height =240
                            Name ="Label110"
                            Caption ="هزينه هاي استهلاك(غير توليدي):"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =1
                    TextFontCharSet =178
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2880
                    Top =13288
                    Width =1828
                    Height =345
                    FontSize =10
                    TabIndex =34
                    Name ="SAYERHAZ"
                    ControlSource ="=[Forms]![SURAT_MALI_MAIN]![SAYERHAZ]"
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextFontCharSet =178
                            TextAlign =1
                            Left =4796
                            Top =13288
                            Width =2895
                            Height =240
                            Name ="Label112"
                            Caption ="ساير هزينه ها:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =1
                    TextFontCharSet =178
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2895
                    Top =13678
                    Width =1828
                    Height =345
                    FontSize =10
                    TabIndex =35
                    BackColor =16777164
                    Name ="JHAZ"
                    ControlSource ="=[Forms]![SURAT_MALI_MAIN]![JHAZ]"
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextFontCharSet =178
                            TextAlign =1
                            Left =4811
                            Top =13678
                            Width =2895
                            Height =240
                            Name ="Label114"
                            Caption ="جمع هزينه ها:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =1
                    TextFontCharSet =178
                    TextAlign =2
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2655
                    Top =14055
                    Width =2323
                    Height =510
                    FontSize =12
                    TabIndex =36
                    BackColor =13434828
                    Name ="SUD"
                    ControlSource ="=[Forms]![SURAT_MALI_MAIN]![SUD]"
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextFontCharSet =178
                            TextAlign =1
                            Left =5051
                            Top =14160
                            Width =2895
                            Height =360
                            Name ="Label116"
                            Caption ="سود و زيان ويژه دوره عملكرد:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
            End
        End
        Begin PageFooter
            Height =0
            Name ="PageFooterSection"
        End
    End
End
CodeBehindForm
' See "SUDVAZIAN.cls"
