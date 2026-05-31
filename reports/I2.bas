Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    FilterOn = NotDefault
    DefaultView =0
    AllowUpdating =4
    TabularCharSet =178
    TabularFamily =42
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =60
    GridY =60
    Width =10696
    DatasheetFontHeight =10
    ItemSuffix =321
    Left =600
    Top =540
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x81614820f429e440
    End
    RecordSource ="QSL_HAVLAH_ANBAR"
    Caption ="فاكتور فروش"
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
        Begin PageBreak
            Width =283
        End
        Begin BreakLevel
            GroupFooter = NotDefault
            KeepTogether =1
            ControlSource ="NUMBER"
        End
        Begin BreakLevel
            ControlSource ="RADIF"
        End
        Begin PageHeader
            Height =2267
            OnFormat ="[Event Procedure]"
            Name ="PageHeader"
            Begin
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =6777
                    Width =3873
                    Height =300
                    FontSize =7
                    TabIndex =1
                    BorderColor =32768
                    Name ="Text90"
                    ControlSource ="=[Forms]![Baseknow]![WIDTH_D]"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5577
                    Top =525
                    Width =696
                    Height =285
                    FontSize =6
                    TabIndex =2
                    BorderColor =32768
                    Name ="DATE_N"
                    ControlSource ="DATE_N"
                    FontName ="WeblogmaYekan"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            TextFontFamily =2
                            Left =6300
                            Top =525
                            Width =420
                            Height =300
                            FontSize =7
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label92"
                            Caption ="تاريخ :"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5592
                    Top =240
                    Width =681
                    Height =270
                    FontSize =7
                    TabIndex =3
                    BorderColor =32768
                    Name ="NUMBER1"
                    ControlSource ="NUMBER1"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextFontFamily =2
                            Left =6298
                            Top =240
                            Width =495
                            Height =300
                            FontSize =7
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label93"
                            Caption ="شماره :"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =7140
                    Top =810
                    Width =3051
                    Height =300
                    FontSize =7
                    TabIndex =4
                    BorderColor =32768
                    Name ="CUST_NO2"
                    ControlSource ="=[CUST_NO] & \"   /    \" & [HESAB]"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextFontFamily =2
                            Left =10204
                            Top =810
                            Width =465
                            Height =300
                            FontSize =7
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label95"
                            Caption ="مشتري:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =7032
                    Top =1395
                    Width =3036
                    Height =300
                    FontSize =7
                    BorderColor =32768
                    Name ="MOLAH"
                    ControlSource ="MOLAH"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextFontFamily =2
                            Left =10077
                            Top =1395
                            Width =600
                            Height =300
                            FontSize =7
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label155"
                            Caption ="توضيحات:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin Line
                    BorderWidth =1
                    Left =5562
                    Top =810
                    Width =5133
                    BorderColor =0
                    Name ="Line164"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5562
                    Top =1395
                    Width =1011
                    Height =300
                    FontSize =7
                    TabIndex =5
                    BorderColor =32768
                    Name ="TEL"
                    ControlSource ="TEL"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextFontFamily =2
                            Left =6612
                            Top =1395
                            Width =375
                            Height =300
                            FontSize =7
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label194"
                            Caption ="تلفن:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =6795
                    Top =240
                    Width =996
                    Height =285
                    FontSize =7
                    TabIndex =6
                    BorderColor =32768
                    Name ="ECODE"
                    ControlSource ="=[Forms]![Baseknow]![ECODE]"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextFontFamily =2
                            Left =7801
                            Top =240
                            Width =750
                            Height =300
                            FontSize =7
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label225"
                            Caption ="ش. اقتصادي:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =8982
                    Top =240
                    Width =1161
                    Height =285
                    FontSize =7
                    TabIndex =7
                    BorderColor =32768
                    Name ="Text208"
                    ControlSource ="=[Forms]![Baseknow]![MCODEM]"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextFontFamily =2
                            Left =10159
                            Top =240
                            Width =480
                            Height =300
                            FontSize =7
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label209"
                            Caption ="ش. ملي:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin Line
                    BorderWidth =1
                    Left =6747
                    Top =15
                    Width =0
                    Height =805
                    BorderColor =0
                    Name ="Line219"
                End
                Begin Line
                    BorderWidth =1
                    Left =6747
                    Top =540
                    Width =3948
                    BorderColor =0
                    Name ="Line220"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =8007
                    Top =525
                    Width =1185
                    Height =300
                    FontSize =7
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label221"
                    Caption ="صورتحساب فروش كالا"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5562
                    Top =768
                    Width =1026
                    Height =285
                    FontSize =7
                    TabIndex =8
                    BorderColor =32768
                    Name ="MCODEM"
                    ControlSource ="MCODEM"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextFontFamily =2
                            Left =6611
                            Top =768
                            Width =510
                            Height =300
                            FontSize =7
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label227"
                            Caption ="ش.ملي:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5562
                    Top =1066
                    Width =1026
                    Height =285
                    FontSize =7
                    TabIndex =9
                    BorderColor =32768
                    Name ="PCODE"
                    ControlSource ="PCODE"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextFontFamily =2
                            Left =6616
                            Top =1066
                            Width =405
                            Height =300
                            FontSize =7
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label229"
                            Caption ="ك.پ:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =7062
                    Top =1066
                    Width =3186
                    Height =285
                    FontSize =7
                    TabIndex =10
                    BorderColor =32768
                    Name ="ADDRESS"
                    ControlSource ="ADDRESS"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextFontFamily =2
                            Left =10269
                            Top =1066
                            Width =390
                            Height =300
                            FontSize =7
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label97"
                            Caption ="نشاني:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin Rectangle
                    BackStyle =1
                    Left =5562
                    Top =1965
                    Width =5129
                    Height =302
                    BackColor =10092543
                    Name ="Box224"
                End
                Begin Line
                    Left =10227
                    Top =1965
                    Width =0
                    Height =300
                    BorderColor =0
                    Name ="Line131"
                End
                Begin Line
                    Left =7992
                    Top =1965
                    Width =0
                    Height =300
                    BorderColor =0
                    Name ="Line133"
                End
                Begin Line
                    Left =7362
                    Top =1965
                    Width =0
                    Height =300
                    BorderColor =0
                    Name ="Line136"
                End
                Begin Line
                    Left =6642
                    Top =1965
                    Width =0
                    Height =300
                    BorderColor =0
                    Name ="Line137"
                End
                Begin Line
                    BorderWidth =1
                    Left =5562
                    Top =1965
                    Width =0
                    Height =300
                    BorderColor =0
                    Name ="Line138"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =5862
                    Top =1965
                    Width =390
                    Height =285
                    FontSize =7
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label139"
                    Caption ="مبــلغ"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =6777
                    Top =1965
                    Width =375
                    Height =285
                    FontSize =7
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label140"
                    Caption ="فـــي"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =7437
                    Top =1980
                    Width =510
                    Height =285
                    FontSize =7
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label141"
                    Caption ="وزن كالا"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =8675
                    Top =1980
                    Width =1035
                    Height =285
                    FontSize =7
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label144"
                    Caption ="شــــــــــرح  كالا"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =10272
                    Top =1980
                    Width =375
                    Height =285
                    FontSize =7
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label146"
                    Caption ="رديف"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =8832
                    Top =1650
                    Width =1011
                    Height =300
                    FontSize =7
                    TabIndex =11
                    BorderColor =32768
                    Name ="CAMIUN_NUM"
                    ControlSource ="CAMIUN_NUM"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextFontFamily =2
                            Left =9881
                            Top =1654
                            Width =795
                            Height =300
                            FontSize =7
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label232"
                            Caption ="شماره ماشين:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =6972
                    Top =1665
                    Width =1176
                    Height =300
                    FontSize =7
                    TabIndex =12
                    BorderColor =32768
                    Name ="DRIVER"
                    ControlSource ="DRIVER"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextFontFamily =2
                            Left =8171
                            Top =1669
                            Width =645
                            Height =300
                            FontSize =7
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label234"
                            Caption ="نام راننده:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5562
                    Top =1650
                    Width =576
                    Height =270
                    FontSize =7
                    TabIndex =13
                    BorderColor =32768
                    Name ="Text235"
                    ControlSource ="NUMBER"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextFontFamily =2
                            Left =6162
                            Top =1650
                            Width =870
                            Height =300
                            FontSize =7
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label236"
                            Caption ="ش. قبض انبار:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin Rectangle
                    Left =5562
                    Width =5133
                    Height =1941
                    Name ="Box237"
                End
                Begin Label
                    TextFontFamily =2
                    Left =5562
                    Width =795
                    Height =270
                    FontSize =7
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label239"
                    Caption ="(نسخه مشتري)"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1217
                    Width =3873
                    Height =300
                    FontSize =7
                    TabIndex =14
                    BorderColor =32768
                    Name ="Text240"
                    ControlSource ="=[Forms]![Baseknow]![WIDTH_D]"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =17
                    Top =525
                    Width =696
                    Height =285
                    FontSize =6
                    TabIndex =15
                    BorderColor =32768
                    Name ="Text241"
                    ControlSource ="DATE_N"
                    FontName ="WeblogmaYekan"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            TextFontFamily =2
                            Left =740
                            Top =525
                            Width =420
                            Height =300
                            FontSize =7
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label242"
                            Caption ="تاريخ :"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =32
                    Top =240
                    Width =681
                    Height =270
                    FontSize =7
                    TabIndex =16
                    BorderColor =32768
                    Name ="Text243"
                    ControlSource ="NUMBER1"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextFontFamily =2
                            Left =738
                            Top =240
                            Width =495
                            Height =300
                            FontSize =7
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label244"
                            Caption ="شماره :"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1590
                    Top =795
                    Width =3051
                    Height =300
                    FontSize =7
                    TabIndex =17
                    BorderColor =32768
                    Name ="CUST_NO3"
                    ControlSource ="=[CUST_NO] & \"   /    \" & [HESAB]"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextFontFamily =2
                            Left =4654
                            Top =795
                            Width =465
                            Height =300
                            FontSize =7
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label246"
                            Caption ="مشتري:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1472
                    Top =1395
                    Width =3036
                    Height =300
                    FontSize =7
                    TabIndex =18
                    BorderColor =32768
                    Name ="Text247"
                    ControlSource ="MOLAH"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextFontFamily =2
                            Left =4517
                            Top =1395
                            Width =600
                            Height =300
                            FontSize =7
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label248"
                            Caption ="توضيحات:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin Line
                    BorderWidth =1
                    Left =2
                    Top =810
                    Width =5118
                    BorderColor =0
                    Name ="Line249"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2
                    Top =1395
                    Width =1011
                    Height =300
                    FontSize =7
                    TabIndex =19
                    BorderColor =32768
                    Name ="Text250"
                    ControlSource ="TEL"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextFontFamily =2
                            Left =1052
                            Top =1395
                            Width =375
                            Height =300
                            FontSize =7
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label251"
                            Caption ="تلفن:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1217
                    Top =240
                    Width =997
                    Height =285
                    FontSize =7
                    TabIndex =20
                    BorderColor =32768
                    Name ="Text252"
                    ControlSource ="=[Forms]![Baseknow]![ECODE]"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextFontFamily =2
                            Left =2238
                            Top =240
                            Width =825
                            Height =300
                            FontSize =7
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label253"
                            Caption ="ش. اقتصادي:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3422
                    Top =240
                    Width =1161
                    Height =285
                    FontSize =7
                    TabIndex =21
                    BorderColor =32768
                    Name ="Text254"
                    ControlSource ="=[Forms]![Baseknow]![MCODEM]"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextFontFamily =2
                            Left =4599
                            Top =240
                            Width =480
                            Height =300
                            FontSize =7
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label255"
                            Caption ="ش. ملي:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin Line
                    BorderWidth =1
                    Left =1187
                    Top =15
                    Width =0
                    Height =805
                    BorderColor =0
                    Name ="Line256"
                End
                Begin Line
                    BorderWidth =1
                    Left =1187
                    Top =540
                    Width =3933
                    BorderColor =0
                    Name ="Line257"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =2447
                    Top =525
                    Width =1185
                    Height =300
                    FontSize =7
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label258"
                    Caption ="صورتحساب فروش كالا"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2
                    Top =768
                    Width =1026
                    Height =285
                    FontSize =7
                    TabIndex =22
                    BorderColor =32768
                    Name ="Text259"
                    ControlSource ="MCODEM"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextFontFamily =2
                            Left =1051
                            Top =768
                            Width =510
                            Height =300
                            FontSize =7
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label260"
                            Caption ="ش.ملي:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2
                    Top =1066
                    Width =1026
                    Height =285
                    FontSize =7
                    TabIndex =23
                    BorderColor =32768
                    Name ="Text261"
                    ControlSource ="PCODE"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextFontFamily =2
                            Left =1056
                            Top =1066
                            Width =405
                            Height =300
                            FontSize =7
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label262"
                            Caption ="ك.پ:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1502
                    Top =1066
                    Width =3186
                    Height =285
                    FontSize =7
                    TabIndex =24
                    BorderColor =32768
                    Name ="Text263"
                    ControlSource ="ADDRESS"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextFontFamily =2
                            Left =4709
                            Top =1066
                            Width =390
                            Height =300
                            FontSize =7
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label264"
                            Caption ="نشاني:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin Rectangle
                    BackStyle =1
                    Left =2
                    Top =1965
                    Width =5129
                    Height =302
                    BackColor =10092543
                    Name ="Box265"
                End
                Begin Line
                    Left =4667
                    Top =1965
                    Width =0
                    Height =300
                    BorderColor =0
                    Name ="Line266"
                End
                Begin Line
                    Left =2432
                    Top =1965
                    Width =0
                    Height =300
                    BorderColor =0
                    Name ="Line267"
                End
                Begin Line
                    Left =1802
                    Top =1965
                    Width =0
                    Height =300
                    BorderColor =0
                    Name ="Line268"
                End
                Begin Line
                    Left =1082
                    Top =1965
                    Width =0
                    Height =300
                    BorderColor =0
                    Name ="Line269"
                End
                Begin Line
                    BorderWidth =1
                    Left =2
                    Top =1965
                    Width =0
                    Height =300
                    BorderColor =0
                    Name ="Line270"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =302
                    Top =1965
                    Width =390
                    Height =285
                    FontSize =7
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label271"
                    Caption ="مبــلغ"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =1217
                    Top =1965
                    Width =375
                    Height =285
                    FontSize =7
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label272"
                    Caption ="فـــي"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =1877
                    Top =1980
                    Width =510
                    Height =285
                    FontSize =7
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label273"
                    Caption ="وزن كالا"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =3115
                    Top =1980
                    Width =1035
                    Height =285
                    FontSize =7
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label274"
                    Caption ="شــــــــــرح  كالا"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =4712
                    Top =1980
                    Width =375
                    Height =285
                    FontSize =7
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label275"
                    Caption ="رديف"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3272
                    Top =1650
                    Width =1011
                    Height =300
                    FontSize =7
                    TabIndex =25
                    BorderColor =32768
                    Name ="Text276"
                    ControlSource ="CAMIUN_NUM"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextFontFamily =2
                            Left =4321
                            Top =1654
                            Width =795
                            Height =300
                            FontSize =7
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label277"
                            Caption ="شماره ماشين:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1412
                    Top =1665
                    Width =1176
                    Height =300
                    FontSize =7
                    TabIndex =26
                    BorderColor =32768
                    Name ="Text278"
                    ControlSource ="DRIVER"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextFontFamily =2
                            Left =2611
                            Top =1669
                            Width =645
                            Height =300
                            FontSize =7
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label279"
                            Caption ="نام راننده:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2
                    Top =1650
                    Width =576
                    Height =270
                    FontSize =7
                    TabIndex =27
                    BorderColor =32768
                    Name ="Text280"
                    ControlSource ="NUMBER"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextFontFamily =2
                            Left =602
                            Top =1650
                            Width =870
                            Height =300
                            FontSize =7
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label281"
                            Caption ="ش. قبض انبار:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin Rectangle
                    Left =2
                    Width =5133
                    Height =1941
                    Name ="Box282"
                End
                Begin Label
                    TextFontFamily =2
                    Top =18
                    Width =990
                    Height =270
                    FontSize =7
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label283"
                    Caption ="(نسخه حسابداري)"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    BorderWidth =1
                    Top =1665
                    Width =5130
                    BorderColor =0
                    Name ="Line314"
                End
                Begin Line
                    BorderWidth =1
                    Left =5565
                    Top =1665
                    Width =5130
                    BorderColor =0
                    Name ="Line317"
                End
                Begin TextBox
                    Visible = NotDefault
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =5149
                    Top =784
                    Width =366
                    TabIndex =28
                    Name ="HESAB"
                    ControlSource ="HESAB"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =5168
                    Top =1398
                    Width =501
                    TabIndex =29
                    Name ="CUST_NO"
                    ControlSource ="CUST_NO"
                    FontName ="WeblogmaYekan"
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =375
            Name ="Detail"
            Begin
                Begin TextBox
                    DecimalPlaces =0
                    RunningSum =1
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =10261
                    Top =15
                    Width =414
                    Height =300
                    FontSize =7
                    BorderColor =32768
                    Name ="RADIF123"
                    ControlSource ="=1"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =8022
                    Top =16
                    Width =2184
                    Height =300
                    FontSize =7
                    TabIndex =1
                    BorderColor =32768
                    Name ="CODE1"
                    ControlSource ="=IIf([Forms]![baseknow]![codeview]=-1,[kala] & \" \" & [code],[kala])"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =7422
                    Top =15
                    Width =549
                    Height =300
                    FontSize =7
                    TabIndex =2
                    BorderColor =32768
                    Name ="MEGHk"
                    ControlSource ="MEGHk"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =6667
                    Top =15
                    Width =669
                    Height =300
                    FontSize =7
                    TabIndex =3
                    BorderColor =32768
                    Name ="MABL"
                    ControlSource ="MABL"
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5577
                    Top =15
                    Width =1029
                    Height =300
                    FontSize =7
                    TabIndex =4
                    BorderColor =32768
                    Name ="MABL_K"
                    ControlSource ="MABL_K"
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Left =5562
                    Top =360
                    Width =5133
                    BorderColor =0
                    Name ="Line100"
                End
                Begin Line
                    BorderWidth =1
                    Left =10695
                    Width =0
                    Height =375
                    BorderColor =0
                    Name ="Line101"
                End
                Begin Line
                    BorderWidth =1
                    Left =10227
                    Width =0
                    Height =375
                    BorderColor =0
                    Name ="Line102"
                End
                Begin Line
                    BorderWidth =1
                    Left =7992
                    Width =0
                    Height =375
                    BorderColor =0
                    Name ="Line104"
                End
                Begin Line
                    BorderWidth =1
                    Left =7362
                    Width =0
                    Height =375
                    BorderColor =0
                    Name ="Line107"
                End
                Begin Line
                    BorderWidth =1
                    Left =6642
                    Width =0
                    Height =375
                    BorderColor =0
                    Name ="Line108"
                End
                Begin Line
                    BorderWidth =1
                    Left =5562
                    Width =0
                    Height =375
                    BorderColor =0
                    Name ="Line109"
                End
                Begin TextBox
                    Visible = NotDefault
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =8095
                    Top =47
                    Height =195
                    FontSize =7
                    TabIndex =5
                    Name ="CODE"
                    ControlSource ="CODE"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =0
                    RunningSum =1
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4701
                    Top =15
                    Width =414
                    Height =300
                    FontSize =7
                    TabIndex =6
                    BorderColor =32768
                    Name ="Text284"
                    ControlSource ="=1"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2462
                    Top =16
                    Width =2184
                    Height =300
                    FontSize =7
                    TabIndex =7
                    BorderColor =32768
                    Name ="Text285"
                    ControlSource ="=IIf([Forms]![baseknow]![codeview]=-1,[kala] & \" \" & [code],[kala])"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1862
                    Top =15
                    Width =549
                    Height =300
                    FontSize =7
                    TabIndex =8
                    BorderColor =32768
                    Name ="Text286"
                    ControlSource ="MEGHk"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1107
                    Top =15
                    Width =669
                    Height =300
                    FontSize =7
                    TabIndex =9
                    BorderColor =32768
                    Name ="Text287"
                    ControlSource ="MABL"
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =17
                    Top =15
                    Width =1029
                    Height =300
                    FontSize =7
                    TabIndex =10
                    BorderColor =32768
                    Name ="Text288"
                    ControlSource ="MABL_K"
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Left =2
                    Top =360
                    Width =5133
                    BorderColor =0
                    Name ="Line289"
                End
                Begin Line
                    BorderWidth =1
                    Left =5130
                    Width =0
                    Height =375
                    BorderColor =0
                    Name ="Line290"
                End
                Begin Line
                    BorderWidth =1
                    Left =4667
                    Width =0
                    Height =375
                    BorderColor =0
                    Name ="Line291"
                End
                Begin Line
                    BorderWidth =1
                    Left =2432
                    Width =0
                    Height =375
                    BorderColor =0
                    Name ="Line292"
                End
                Begin Line
                    BorderWidth =1
                    Left =1802
                    Width =0
                    Height =375
                    BorderColor =0
                    Name ="Line293"
                End
                Begin Line
                    BorderWidth =1
                    Left =1082
                    Width =0
                    Height =375
                    BorderColor =0
                    Name ="Line294"
                End
                Begin Line
                    BorderWidth =1
                    Left =2
                    Width =0
                    Height =375
                    BorderColor =0
                    Name ="Line295"
                End
                Begin TextBox
                    Visible = NotDefault
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =2535
                    Top =47
                    Height =195
                    FontSize =7
                    TabIndex =11
                    Name ="Text296"
                    ControlSource ="CODE"
                    FontName ="WeblogmaYekan"
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =2406
            Name ="GroupFooter3"
            Begin
                Begin Line
                    Left =5562
                    Top =2085
                    Width =5130
                    BorderColor =0
                    Name ="Line125"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =2
                    Left =5592
                    Top =1125
                    Width =5085
                    Height =240
                    FontSize =7
                    FontWeight =400
                    ForeColor =0
                    Name ="COMM"
                    Caption ="."
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    NumeralShapes =2
                    Left =5577
                    Top =1
                    Width =2235
                    Height =390
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="JF"
                    Caption ="جمع كل فاكتور:"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextFontFamily =2
                    Left =8726
                    Top =375
                    Width =1035
                    Height =300
                    FontSize =10
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label189"
                    Caption ="مبلغ به حروف:"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    OldBorderStyle =1
                    BorderWidth =1
                    TextAlign =3
                    TextFontFamily =2
                    Left =5577
                    Top =394
                    Width =3130
                    Height =270
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="HR"
                    Caption ="مبلغ به حروف:"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =8141
                    Top =1833
                    Width =2529
                    FontSize =7
                    BorderColor =32768
                    Name ="Text211"
                    ControlSource ="=Format(Now(),\"hh:nn:ss ampm\")"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =1
                    BorderWidth =1
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =7811
                    Top =1
                    Width =894
                    Height =391
                    FontSize =10
                    TabIndex =1
                    Name ="Text200"
                    ControlSource ="=Sum([MEGHk])"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextFontFamily =2
                    Left =8726
                    Width =735
                    Height =390
                    FontSize =10
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label226"
                    Caption ="جمع كل:"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =2
                    Left =5577
                    Top =810
                    Width =5085
                    Height =330
                    FontSize =10
                    FontWeight =400
                    BackColor =128
                    ForeColor =0
                    Name ="MANDAH"
                    Caption ="مانده حساب:"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5605
                    Top =1662
                    Width =1446
                    Height =300
                    FontSize =7
                    TabIndex =2
                    BorderColor =32768
                    Name ="Text209"
                    ControlSource ="=UCURRENTUSER()"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =5954
                            Top =1436
                            Width =675
                            Height =270
                            FontSize =7
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label210"
                            Caption ="صادر كننده:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =2
                    Left =9210
                    Top =1410
                    Width =1125
                    Height =270
                    FontSize =7
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label228"
                    Caption ="مهر و امضاء فروشنده:"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =7107
                    Top =2100
                    Width =3006
                    Height =300
                    FontSize =8
                    TabIndex =3
                    BorderColor =32768
                    Name ="Text199"
                    ControlSource ="=[Forms]![Baseknow]![TFADDRESS]"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =10152
                            Top =2100
                            Width =495
                            Height =300
                            FontSize =7
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label200"
                            Caption ="آدرس:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5607
                    Top =2100
                    Width =1086
                    Height =270
                    FontSize =7
                    TabIndex =4
                    BorderColor =32768
                    Name ="Text197"
                    ControlSource ="=[Forms]![Baseknow]![TFTEL]"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =6717
                            Top =2100
                            Width =375
                            Height =300
                            FontSize =7
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label198"
                            Caption ="تلفن:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin Rectangle
                    Left =5562
                    Width =5133
                    Height =2406
                    Name ="Box238"
                End
                Begin Line
                    Left =2
                    Top =2085
                    Width =5130
                    BorderColor =0
                    Name ="Line297"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =2
                    Left =32
                    Top =1125
                    Width =5085
                    Height =240
                    FontSize =7
                    FontWeight =400
                    ForeColor =0
                    Name ="COMM2"
                    Caption ="."
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    NumeralShapes =2
                    Left =17
                    Top =1
                    Width =2235
                    Height =390
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="jf2"
                    Caption ="جمع كل فاكتور:"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextFontFamily =2
                    Left =3166
                    Top =375
                    Width =1035
                    Height =300
                    FontSize =10
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label300"
                    Caption ="مبلغ به حروف:"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    OldBorderStyle =1
                    BorderWidth =1
                    TextAlign =3
                    TextFontFamily =2
                    Left =17
                    Top =394
                    Width =3130
                    Height =270
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="hr2"
                    Caption ="مبلغ به حروف:"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2581
                    Top =1833
                    Width =2529
                    FontSize =7
                    TabIndex =5
                    BorderColor =32768
                    Name ="Text302"
                    ControlSource ="=Format(Now(),\"hh:nn:ss ampm\")"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =1
                    BorderWidth =1
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2251
                    Top =1
                    Width =894
                    Height =391
                    FontSize =10
                    TabIndex =6
                    Name ="Text303"
                    ControlSource ="=Sum([MEGHk])"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextFontFamily =2
                    Left =3166
                    Width =735
                    Height =390
                    FontSize =10
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label304"
                    Caption ="جمع كل:"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =2
                    Left =17
                    Top =810
                    Width =5085
                    Height =330
                    FontSize =10
                    FontWeight =400
                    BackColor =128
                    ForeColor =0
                    Name ="MANDAH2"
                    Caption ="مانده حساب:"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =45
                    Top =1662
                    Width =1446
                    Height =300
                    FontSize =7
                    TabIndex =7
                    BorderColor =32768
                    Name ="Text306"
                    ControlSource ="=UCURRENTUSER()"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =394
                            Top =1436
                            Width =675
                            Height =270
                            FontSize =7
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label307"
                            Caption ="صادر كننده:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =2
                    Left =3540
                    Top =1428
                    Width =1125
                    Height =270
                    FontSize =7
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label308"
                    Caption ="مهر و امضاء فروشنده:"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1547
                    Top =2100
                    Width =3006
                    Height =300
                    FontSize =8
                    TabIndex =8
                    BorderColor =32768
                    Name ="Text309"
                    ControlSource ="=[Forms]![Baseknow]![TFADDRESS]"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =4592
                            Top =2100
                            Width =495
                            Height =300
                            FontSize =7
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label310"
                            Caption ="آدرس:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =47
                    Top =2100
                    Width =1086
                    Height =270
                    FontSize =7
                    TabIndex =9
                    BorderColor =32768
                    Name ="Text311"
                    ControlSource ="=[Forms]![Baseknow]![TFTEL]"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =1157
                            Top =2100
                            Width =375
                            Height =300
                            FontSize =7
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label312"
                            Caption ="تلفن:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin Rectangle
                    Left =2
                    Width =5133
                    Height =2406
                    Name ="Box313"
                End
                Begin TextBox
                    Visible = NotDefault
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =2872
                    Top =963
                    TabIndex =10
                    Name ="TAG"
                    ControlSource ="TAG"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            Visible = NotDefault
                            TextAlign =0
                            TextFontFamily =2
                            Left =5749
                            Top =963
                            Width =525
                            Height =270
                            FontWeight =400
                            Name ="Label320"
                            Caption ="TAG:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
            End
        End
        Begin PageFooter
            Visible = NotDefault
            Height =0
            Name ="PageFooter"
        End
    End
End
CodeBehindForm
' See "I2.cls"
