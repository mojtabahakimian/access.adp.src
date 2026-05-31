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
    GridX =64
    GridY =64
    Width =15440
    DatasheetFontHeight =10
    ItemSuffix =301
    Left =600
    Top =1110
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x90ac73651a9fe240
    End
    Caption ="گزارش ورود و خروج ميمانان"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
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
        Begin Tab
            TextFontCharSet =178
            TextFontFamily =2
            Width =5103
            Height =3402
            FontName ="Arial (Arabic)"
        End
        Begin BreakLevel
            ControlSource ="ROOMN"
        End
        Begin BreakLevel
            ControlSource ="IDH"
        End
        Begin PageHeader
            Height =1537
            Name ="PageHeader"
            Begin
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =14936
                    Top =975
                    Width =450
                    Height =510
                    FontSize =8
                    ForeColor =0
                    Name ="N_S Label"
                    Caption ="رديف"
                    FontName ="Zar"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="N_S_Label"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =13832
                    Top =975
                    Width =1081
                    Height =510
                    FontSize =8
                    ForeColor =0
                    Name ="SHARH Label"
                    Caption ="نام "
                    FontName ="Zar"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="SHARH_Label"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =9063
                    Top =975
                    Width =600
                    Height =510
                    FontSize =8
                    ForeColor =0
                    Name ="BES Label"
                    Caption ="شماره شناسنامه"
                    FontName ="Zar"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="BES_Label"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =2
                    Left =15
                    Top =915
                    Width =15409
                    BorderColor =0
                    Name ="Line26"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =2
                    Left =15
                    Top =1515
                    Width =15410
                    BorderColor =0
                    Name ="Line27"
                End
                Begin Line
                    Left =14939
                    Top =919
                    Width =0
                    Height =618
                    Name ="Line29"
                End
                Begin Line
                    BorderWidth =2
                    Left =15392
                    Top =919
                    Width =0
                    Height =618
                    BorderColor =0
                    Name ="Line37"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =6333
                    Top =513
                    Width =2685
                    Height =390
                    FontSize =10
                    ForeColor =0
                    Name ="Label22"
                    Caption ="گزارش روزانه ورود و خروج ميهمانان"
                    FontName ="Titr"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =10382
                    Top =975
                    Width =600
                    Height =510
                    FontSize =8
                    ForeColor =0
                    Name ="Label73"
                    Caption ="تاريخ تولد"
                    FontName ="Zar"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    Left =11504
                    Top =915
                    Width =0
                    Height =599
                    Name ="Line75"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =11252
                    Top =975
                    Width =240
                    Height =510
                    FontSize =7
                    ForeColor =0
                    Name ="Label82"
                    Caption ="مرد"
                    FontName ="Zar"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =11042
                    Top =975
                    Width =210
                    Height =510
                    FontSize =7
                    ForeColor =0
                    Name ="Label83"
                    Caption ="زن"
                    FontName ="Zar"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    Left =11270
                    Top =915
                    Width =0
                    Height =603
                    Name ="Line84"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =9797
                    Top =975
                    Width =495
                    Height =510
                    FontSize =8
                    ForeColor =0
                    Name ="Label118"
                    Caption ="تابعيت"
                    FontName ="Zar"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =8391
                    Top =975
                    Width =600
                    Height =510
                    FontSize =8
                    ForeColor =0
                    Name ="Label122"
                    Caption ="محل صدور"
                    FontName ="Zar"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    Left =9019
                    Top =919
                    Width =0
                    Height =618
                    Name ="Line123"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =7521
                    Top =975
                    Width =810
                    Height =510
                    FontSize =8
                    ForeColor =0
                    Name ="Label126"
                    Caption ="شغل"
                    FontName ="Zar"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    Left =8359
                    Top =919
                    Width =0
                    Height =618
                    Name ="Line127"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =6605
                    Top =975
                    Width =825
                    Height =510
                    FontSize =8
                    ForeColor =0
                    Name ="Label130"
                    Caption ="قصد از مسافرت"
                    FontName ="Zar"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    Left =7468
                    Top =915
                    Width =0
                    Height =618
                    Name ="Line131"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =5744
                    Top =975
                    Width =765
                    Height =510
                    FontSize =8
                    ForeColor =0
                    Name ="Label134"
                    Caption ="شهر مبدا"
                    FontName ="Zar"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    Left =6537
                    Top =919
                    Width =0
                    Height =618
                    Name ="Line135"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =4911
                    Top =975
                    Width =750
                    Height =510
                    FontSize =8
                    ForeColor =0
                    Name ="Label138"
                    Caption ="شهر مقصد"
                    FontName ="Zar"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    Left =5697
                    Top =919
                    Width =0
                    Height =618
                    Name ="Line139"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =4070
                    Top =975
                    Width =765
                    Height =510
                    FontSize =8
                    ForeColor =0
                    Name ="Label142"
                    Caption ="وسيله سفر"
                    FontName ="Zar"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    Left =4872
                    Top =919
                    Width =0
                    Height =618
                    Name ="Line143"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =3080
                    Top =975
                    Width =915
                    Height =510
                    FontSize =8
                    ForeColor =0
                    Name ="Label146"
                    Caption ="تاريخ ورود"
                    FontName ="Zar"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    Left =4023
                    Top =919
                    Width =0
                    Height =618
                    Name ="Line147"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =1868
                    Top =975
                    Width =555
                    Height =510
                    FontSize =8
                    ForeColor =0
                    Name ="Label150"
                    Caption ="تخت اضافه"
                    FontName ="Zar"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    Left =2436
                    Top =919
                    Width =0
                    Height =618
                    Name ="Line151"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =1215
                    Top =975
                    Width =585
                    Height =510
                    FontSize =8
                    ForeColor =0
                    Name ="Label154"
                    Caption ="شماره اتاق"
                    FontName ="Zar"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    Left =1834
                    Top =915
                    Width =0
                    Height =618
                    Name ="Line155"
                End
                Begin Line
                    Left =12341
                    Top =919
                    Width =0
                    Height =618
                    Name ="Line158"
                End
                Begin Line
                    Left =11035
                    Top =915
                    Width =0
                    Height =603
                    Name ="Line178"
                End
                Begin Line
                    Left =10324
                    Top =915
                    Width =0
                    Height =618
                    Name ="Line179"
                End
                Begin Line
                    BorderWidth =2
                    Left =15
                    Top =915
                    Width =0
                    Height =618
                    BorderColor =0
                    Name ="Line206"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =2468
                    Top =975
                    Width =570
                    Height =510
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="Label214"
                    Caption ="تعداد همراهان"
                    FontName ="Zar"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    Left =3051
                    Top =919
                    Width =0
                    Height =618
                    Name ="Line215"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =11582
                    Top =975
                    Width =720
                    Height =510
                    FontSize =8
                    ForeColor =0
                    Name ="Label218"
                    Caption ="نام پدر"
                    FontName ="Zar"
                    Tag ="DetachedLabel"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4905
                    Top =15
                    Width =5433
                    Height =540
                    FontSize =14
                    BorderColor =32768
                    Name ="Text90"
                    ControlSource ="=[Forms]![Baseknow]![WIDTH_D]"
                    FontName ="Titr"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =12392
                    Top =975
                    Width =1365
                    Height =510
                    FontSize =8
                    ForeColor =0
                    Name ="Label236"
                    Caption ="نام خانوادگي"
                    FontName ="Zar"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    Left =13784
                    Top =919
                    Width =0
                    Height =618
                    Name ="Line237"
                End
                Begin Line
                    Left =9739
                    Top =919
                    Width =0
                    Height =618
                    Name ="Line240"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =60
                    Top =975
                    Width =1080
                    Height =510
                    FontSize =8
                    ForeColor =0
                    Name ="Label242"
                    Caption ="ملاحظات"
                    FontName ="Zar"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    Left =1174
                    Top =915
                    Width =0
                    Height =618
                    Name ="Line243"
                End
                Begin TextBox
                    TextFontFamily =10
                    ReadingOrder =1
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =13570
                    Top =433
                    Width =1231
                    Height =360
                    FontWeight =700
                    TabIndex =1
                    Name ="Text259"
                    ControlSource ="=[Forms]![F_MENU_HOTEL]![DT1]"
                    FontName ="Titr"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            TextAlign =0
                            TextFontFamily =10
                            Left =14822
                            Top =433
                            Width =465
                            Height =360
                            ForeColor =0
                            Name ="Label260"
                            Caption ="تاريخ:"
                            FontName ="Titr"
                        End
                    End
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            CanGrow = NotDefault
            Height =9478
            Name ="Detail"
            Begin
                Begin Subform
                    OldBorderStyle =1
                    Left =25
                    Width =15397
                    Height =4740
                    Name ="HOTEL_SUB1"
                    SourceObject ="Report.HOTEL_SUB1"
                End
                Begin Line
                    Left =14939
                    Width =0
                    Height =4756
                    BorderColor =0
                    Name ="Line39"
                End
                Begin Line
                    Left =12341
                    Width =0
                    Height =4756
                    BorderColor =0
                    Name ="Line43"
                End
                Begin Line
                    BorderWidth =2
                    Left =15392
                    Width =0
                    Height =4756
                    BorderColor =0
                    Name ="Line47"
                End
                Begin Line
                    Left =11524
                    Width =0
                    Height =4756
                    BorderColor =0
                    Name ="Line79"
                End
                Begin Line
                    Left =11029
                    Width =0
                    Height =4756
                    BorderColor =0
                    Name ="Line81"
                End
                Begin Line
                    Left =11264
                    Width =0
                    Height =4756
                    BorderColor =0
                    Name ="Line92"
                End
                Begin Line
                    Left =10317
                    Width =0
                    Height =4756
                    BorderColor =0
                    Name ="Line121"
                End
                Begin Line
                    Left =9034
                    Width =0
                    Height =4756
                    BorderColor =0
                    Name ="Line125"
                End
                Begin Line
                    Left =8375
                    Width =0
                    Height =4756
                    BorderColor =0
                    Name ="Line129"
                End
                Begin Line
                    Left =7483
                    Width =0
                    Height =4756
                    BorderColor =0
                    Name ="Line133"
                End
                Begin Line
                    Left =6552
                    Width =0
                    Height =4756
                    BorderColor =0
                    Name ="Line137"
                End
                Begin Line
                    Left =5712
                    Width =0
                    Height =4756
                    BorderColor =0
                    Name ="Line141"
                End
                Begin Line
                    Left =4871
                    Width =0
                    Height =4756
                    BorderColor =0
                    Name ="Line145"
                End
                Begin Line
                    Left =4022
                    Width =0
                    Height =4756
                    BorderColor =0
                    Name ="Line149"
                End
                Begin Line
                    Left =2434
                    Width =0
                    Height =4756
                    BorderColor =0
                    Name ="Line153"
                End
                Begin Line
                    Left =1834
                    Width =0
                    Height =4756
                    BorderColor =0
                    Name ="Line157"
                End
                Begin Line
                    BorderWidth =2
                    Left =15
                    Width =0
                    Height =4756
                    BorderColor =0
                    Name ="Line207"
                End
                Begin Line
                    Left =3049
                    Width =0
                    Height =4756
                    BorderColor =0
                    Name ="Line217"
                End
                Begin Line
                    Left =13784
                    Width =0
                    Height =4756
                    BorderColor =0
                    Name ="Line239"
                End
                Begin Line
                    Left =9754
                    Width =0
                    Height =4756
                    BorderColor =0
                    Name ="Line241"
                End
                Begin Line
                    Left =1174
                    Width =0
                    Height =4756
                    BorderColor =0
                    Name ="Line244"
                End
                Begin Subform
                    OldBorderStyle =1
                    Left =31
                    Top =5102
                    Width =6862
                    Height =4320
                    TabIndex =1
                    Name ="Child248"
                    SourceObject ="Report.HOTEL_SUB2"
                    Begin
                        Begin Label
                            TextAlign =2
                            TextFontFamily =10
                            Left =4635
                            Top =4783
                            Width =2265
                            Height =375
                            FontSize =10
                            FontWeight =400
                            ForeColor =0
                            Name ="Label258"
                            Caption ="ب:ميهمانهائيكه از هتل خارج شده اند"
                            FontName ="Zar"
                        End
                    End
                End
                Begin Line
                    Left =6420
                    Top =5115
                    Width =0
                    Height =4321
                    BorderColor =0
                    Name ="Line249"
                End
                Begin Line
                    Left =5265
                    Top =5115
                    Width =0
                    Height =4321
                    BorderColor =0
                    Name ="Line251"
                End
                Begin Line
                    Left =3825
                    Top =5115
                    Width =0
                    Height =4321
                    BorderColor =0
                    Name ="Line252"
                End
                Begin Line
                    Left =1995
                    Top =5115
                    Width =0
                    Height =4321
                    BorderColor =0
                    Name ="Line253"
                End
                Begin Line
                    Left =1440
                    Top =5115
                    Width =0
                    Height =4321
                    BorderColor =0
                    Name ="Line254"
                End
                Begin Line
                    BorderWidth =2
                    Top =5115
                    Width =0
                    Height =4321
                    BorderColor =0
                    Name ="Line255"
                End
                Begin Line
                    Left =2985
                    Top =5115
                    Width =0
                    Height =4321
                    BorderColor =0
                    Name ="Line256"
                End
                Begin Line
                    Left =780
                    Top =5115
                    Width =0
                    Height =4321
                    BorderColor =0
                    Name ="Line257"
                End
                Begin Rectangle
                    Left =7529
                    Top =5299
                    Width =7396
                    Height =3266
                    Name ="Box261"
                End
                Begin Line
                    Left =7530
                    Top =5715
                    Width =7396
                    BorderColor =0
                    Name ="Line262"
                End
                Begin Line
                    Left =7531
                    Top =6110
                    Width =7396
                    BorderColor =0
                    Name ="Line263"
                End
                Begin Line
                    Left =7532
                    Top =6505
                    Width =7396
                    BorderColor =0
                    Name ="Line264"
                End
                Begin Line
                    Left =7533
                    Top =6900
                    Width =7396
                    BorderColor =0
                    Name ="Line265"
                End
                Begin Line
                    Left =7534
                    Top =7295
                    Width =7396
                    BorderColor =0
                    Name ="Line266"
                End
                Begin Line
                    Left =7535
                    Top =7690
                    Width =7396
                    BorderColor =0
                    Name ="Line267"
                End
                Begin Line
                    Left =10725
                    Top =8085
                    Width =4201
                    BorderColor =0
                    Name ="Line268"
                End
                Begin Line
                    Left =8503
                    Top =5297
                    Width =0
                    Height =2411
                    BorderColor =0
                    Name ="Line270"
                End
                Begin Line
                    Left =10725
                    Top =4905
                    Width =0
                    Height =3671
                    BorderColor =0
                    Name ="Line271"
                End
                Begin Line
                    Left =11340
                    Top =4905
                    Width =0
                    Height =2786
                    BorderColor =0
                    Name ="Line272"
                End
                Begin Line
                    Left =11970
                    Top =4905
                    Width =0
                    Height =2786
                    BorderColor =0
                    Name ="Line274"
                End
                Begin Line
                    Left =12614
                    Top =4907
                    Width =0
                    Height =3671
                    BorderColor =0
                    Name ="Line275"
                End
                Begin Label
                    TextFontFamily =10
                    Left =13624
                    Top =4889
                    Width =1140
                    Height =390
                    FontSize =10
                    ForeColor =0
                    Name ="Label276"
                    Caption ="پ:جدول نتيجه"
                    FontName ="Titr"
                End
                Begin Label
                    TextFontFamily =10
                    Left =11987
                    Top =4935
                    Width =615
                    Height =330
                    ForeColor =0
                    Name ="Label277"
                    Caption ="يك تخته"
                    FontName ="Zar"
                End
                Begin Label
                    TextFontFamily =10
                    Left =11372
                    Top =4935
                    Width =570
                    Height =330
                    ForeColor =0
                    Name ="Label278"
                    Caption ="دو تخته"
                    FontName ="Zar"
                End
                Begin Label
                    TextFontFamily =10
                    Left =10770
                    Top =4936
                    Width =510
                    Height =330
                    ForeColor =0
                    Name ="Label279"
                    Caption ="سوئيت"
                    FontName ="Zar"
                End
                Begin Line
                    Left =10727
                    Top =4889
                    Width =1899
                    BorderColor =0
                    Name ="Line280"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =10
                    Left =12671
                    Top =5341
                    Width =2175
                    Height =330
                    ForeColor =0
                    Name ="Label281"
                    Caption ="تعداد كل اتاقهاي هتل"
                    FontName ="Zar"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =10
                    Left =12675
                    Top =5730
                    Width =2175
                    Height =330
                    ForeColor =0
                    Name ="Label282"
                    Caption ="تعدا اتاقهاي اشغال شده"
                    FontName ="Zar"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =10
                    Left =12690
                    Top =6150
                    Width =2175
                    Height =330
                    ForeColor =0
                    Name ="Label283"
                    Caption ="تعدادكل تختهاي ثابت هتل"
                    FontName ="Zar"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =10
                    Left =12675
                    Top =6540
                    Width =2175
                    Height =330
                    ForeColor =0
                    Name ="Label284"
                    Caption ="تعداد تختهاي ثابت اشغال شده"
                    FontName ="Zar"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =10
                    Left =12675
                    Top =6930
                    Width =2175
                    Height =330
                    ForeColor =0
                    Name ="Label285"
                    Caption ="تعدا تختهاي اضافه روزهاي قبل"
                    FontName ="Zar"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =10
                    Left =12675
                    Top =7320
                    Width =2175
                    Height =330
                    ForeColor =0
                    Name ="Label286"
                    Caption ="تعداد تختهاي اضافه روز جاري"
                    FontName ="Zar"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =10
                    Left =12675
                    Top =7710
                    Width =2175
                    Height =330
                    ForeColor =0
                    Name ="Label287"
                    Caption ="تعداد مسافران ورودي ايراني"
                    FontName ="Zar"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =10
                    Left =12675
                    Top =8100
                    Width =2175
                    Height =330
                    ForeColor =0
                    Name ="Label288"
                    Caption ="تعدادمسافران ورودي خارجي"
                    FontName ="Zar"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =10
                    Left =8535
                    Top =5355
                    Width =2175
                    Height =330
                    ForeColor =0
                    Name ="Label290"
                    Caption ="تعداد مسافران خروجي ايراني"
                    FontName ="Zar"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =10
                    Left =8535
                    Top =5745
                    Width =2175
                    Height =330
                    ForeColor =0
                    Name ="Label291"
                    Caption ="تعداد مسافران خروجي خارجي"
                    FontName ="Zar"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =10
                    Left =8535
                    Top =6135
                    Width =2175
                    Height =330
                    ForeColor =0
                    Name ="Label292"
                    Caption ="تعداد همراهان روزهاي قبل"
                    FontName ="Zar"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =10
                    Left =8505
                    Top =6525
                    Width =2220
                    Height =330
                    ForeColor =0
                    Name ="Label293"
                    Caption ="تعداد همراهان ورودي روز جاري"
                    FontName ="Zar"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =10
                    Left =8535
                    Top =6915
                    Width =2175
                    Height =330
                    ForeColor =0
                    Name ="Label294"
                    Caption ="تعداد مسافران مقيم شب قبل"
                    FontName ="Zar"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =10
                    Left =8535
                    Top =7305
                    Width =2175
                    Height =330
                    ForeColor =0
                    Name ="Label295"
                    Caption ="تعداد مسافران مقيم شب جاري"
                    FontName ="Zar"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =10
                    Left =8805
                    Top =7740
                    Width =1875
                    Height =315
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="Label298"
                    Caption ="نام و نام خانودگي و امضاء مدير هتل:"
                    FontName ="Zar"
                End
                Begin Line
                    Top =4740
                    Width =15440
                    BorderColor =0
                    Name ="Line299"
                End
                Begin Rectangle
                    Top =5100
                    Width =6897
                    Height =4322
                    Name ="Box300"
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
' See "R_LIST_MEHMAN.cls"
