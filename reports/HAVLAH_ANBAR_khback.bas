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
    Width =10550
    DatasheetFontHeight =10
    ItemSuffix =199
    Left =930
    Top =660
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xa04a1b1b6fa7e340
    End
    RecordSource ="QSL_INVOICE_KH_BACK"
    Caption ="حواله انبار"
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
            GroupFooter = NotDefault
            ControlSource ="NUMBER"
        End
        Begin BreakLevel
            ControlSource ="RADIF"
        End
        Begin PageHeader
            Height =2025
            Name ="PageHeaderSection"
            Begin
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =3510
                    Width =2505
                    Height =540
                    FontSize =10
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label88"
                    Caption ="حواله انبار برگشت خريد"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2307
                    Top =435
                    Width =5103
                    Height =540
                    FontSize =10
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
                    Left =15
                    Top =540
                    Width =1146
                    Height =375
                    FontSize =10
                    TabIndex =1
                    BorderColor =32768
                    Name ="DATE_N"
                    ControlSource ="DATE_N"
                    FontName ="WeblogmaYekan"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =1213
                            Top =540
                            Width =570
                            Height =390
                            FontSize =10
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
                    Left =8130
                    Top =540
                    Width =1311
                    Height =375
                    FontSize =10
                    TabIndex =2
                    BorderColor =32768
                    Name ="NUMBER"
                    ControlSource ="NUMBER"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =9480
                            Top =540
                            Width =1020
                            Height =390
                            FontSize =10
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label93"
                            Caption ="شماره حواله:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin Line
                    BorderWidth =3
                    Top =975
                    Width =10521
                    BorderColor =0
                    Name ="Line94"
                End
                Begin Line
                    BorderWidth =3
                    Top =1470
                    Width =10506
                    BorderColor =0
                    Name ="Line99"
                End
                Begin Line
                    BorderWidth =3
                    Left =10485
                    Top =1440
                    Width =0
                    Height =585
                    BorderColor =0
                    Name ="Line130"
                End
                Begin Line
                    BorderWidth =1
                    Left =9927
                    Top =1485
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line131"
                End
                Begin Line
                    BorderWidth =1
                    Left =7824
                    Top =1485
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line132"
                End
                Begin Line
                    BorderWidth =1
                    Left =2607
                    Top =1485
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line133"
                End
                Begin Line
                    BorderWidth =1
                    Left =1155
                    Top =1485
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line135"
                End
                Begin Line
                    BorderWidth =2
                    Top =1485
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line138"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =1242
                    Top =1530
                    Width =1275
                    Height =405
                    FontSize =10
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label143"
                    Caption ="واحد كالا"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =2823
                    Top =1530
                    Width =4110
                    Height =405
                    FontSize =10
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
                    Left =9405
                    Top =1530
                    Width =435
                    Height =405
                    FontSize =10
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label145"
                    Caption ="انبار"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =9933
                    Top =1526
                    Width =540
                    Height =405
                    FontSize =10
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label146"
                    Caption ="رديف"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    BorderWidth =3
                    Top =2010
                    Width =10506
                    BorderColor =0
                    Name ="Line147"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =15
                    Top =1500
                    Width =1185
                    Height =480
                    FontSize =10
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label163"
                    Caption ="مقدار تحويلي"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =6240
                    Top =1020
                    Width =3561
                    Height =375
                    FontSize =10
                    TabIndex =3
                    BorderColor =32768
                    Name ="NAME"
                    ControlSource ="NAME"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextFontFamily =2
                            Left =9855
                            Top =1020
                            Width =645
                            Height =390
                            FontSize =10
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label171"
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
                    Left =65
                    Top =1020
                    Width =5556
                    Height =375
                    FontSize =10
                    TabIndex =4
                    BorderColor =32768
                    Name ="ADDRESS"
                    ControlSource ="ADDRESS"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =5585
                            Top =1020
                            Width =600
                            Height =390
                            FontSize =10
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label173"
                            Caption ="آدرس:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =7275
                    Top =1530
                    Width =510
                    Height =390
                    FontSize =10
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label176"
                    Caption ="كد"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    BorderWidth =1
                    Left =7215
                    Top =1485
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line178"
                End
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Top =226
                    Width =2336
                    Height =300
                    FontSize =10
                    TabIndex =5
                    Name ="Text24"
                    ControlSource ="=\"صفحه \" & [Page] & \" از \" & [Pages]"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    BorderWidth =1
                    Left =9330
                    Top =1485
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line189"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =7890
                    Top =1530
                    Width =1425
                    Height =390
                    FontSize =10
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label191"
                    Caption ="باركد / شماره فني"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    Left =7531
                    Top =113
                    Width =2226
                    Height =375
                    FontSize =10
                    TabIndex =6
                    BorderColor =32768
                    Name ="Text170"
                    ControlSource ="=[Forms]![BASEKNOW]![ISO_FROOSH]"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =9798
                            Top =113
                            Width =690
                            Height =390
                            FontSize =10
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label195"
                            Caption ="كد فرم :"
                            FontName ="WeblogmaYekan"
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
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =9375
                    Top =15
                    Width =519
                    Height =405
                    FontSize =10
                    BorderColor =32768
                    Name ="IANBAR"
                    ControlSource ="IANBAR"
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
                    Left =9943
                    Top =16
                    Width =504
                    Height =405
                    FontSize =10
                    TabIndex =1
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
                    Left =2627
                    Top =16
                    Width =4524
                    Height =405
                    FontSize =10
                    TabIndex =2
                    BorderColor =32768
                    Name ="KALA"
                    ControlSource ="KALA"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1185
                    Top =15
                    Width =1320
                    Height =405
                    FontSize =10
                    TabIndex =3
                    BorderColor =32768
                    Name ="VAHED_K"
                    ControlSource ="VNAMES"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Top =435
                    Width =10485
                    BorderColor =0
                    Name ="Line100"
                End
                Begin Line
                    BorderWidth =3
                    Left =10485
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line101"
                End
                Begin Line
                    BorderWidth =1
                    Left =9927
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line102"
                End
                Begin Line
                    BorderWidth =1
                    Left =7824
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line103"
                End
                Begin Line
                    BorderWidth =1
                    Left =2607
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line104"
                End
                Begin Line
                    BorderWidth =1
                    Left =1155
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
                    DecimalPlaces =2
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =57
                    Width =1029
                    Height =405
                    FontSize =10
                    TabIndex =4
                    BorderColor =32768
                    Name ="MEGH_MAR"
                    ControlSource ="MEGH_MAR"
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
                    Left =7245
                    Top =15
                    Width =564
                    Height =390
                    FontSize =10
                    TabIndex =5
                    BorderColor =32768
                    Name ="CODE"
                    ControlSource ="CODE"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    BorderWidth =1
                    Left =7215
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line179"
                End
                Begin Line
                    BorderWidth =1
                    Left =9330
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line190"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =7875
                    Top =15
                    Width =1419
                    Height =375
                    FontSize =10
                    TabIndex =6
                    BorderColor =32768
                    Name ="Text193"
                    ControlSource ="N_FANI"
                    FontName ="WeblogmaYekan"
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            CanGrow = NotDefault
            Height =1587
            Name ="GroupFooter3"
            Begin
                Begin Line
                    BorderWidth =3
                    Width =10510
                    BorderColor =0
                    Name ="Line110"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =9206
                    Top =283
                    Width =1080
                    Height =405
                    FontSize =10
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label181"
                    Caption ="تنظيم كننده:"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3530
                    Top =465
                    Width =2601
                    Height =375
                    FontSize =10
                    BorderColor =32768
                    Name ="TAH"
                    ControlSource ="TAH"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextFontFamily =2
                            Left =6185
                            Top =465
                            Width =1080
                            Height =390
                            FontSize =10
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label155"
                            Caption ="تحويل دهنده:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =60
                    Top =465
                    Width =2841
                    Height =375
                    FontSize =10
                    TabIndex =1
                    BorderColor =32768
                    Name ="Text158"
                    ControlSource ="MOLAH"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =2903
                            Top =465
                            Width =555
                            Height =390
                            FontSize =10
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label159"
                            Caption ="توسط:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =9184
                    Top =737
                    Width =1155
                    Height =390
                    FontSize =10
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label187"
                    Caption ="شماره فاكتور:"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =8820
                    Top =1140
                    Width =1515
                    Height =390
                    FontSize =10
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label188"
                    Caption ="شماره برگ خروج:"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextFontFamily =2
                    Left =6180
                    Top =1155
                    Width =1095
                    Height =390
                    FontSize =10
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label197"
                    Caption ="تحويل گيرنده:"
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
                    Left =15
                    Width =1155
                    Height =390
                    FontSize =10
                    TabIndex =2
                    BorderColor =4210752
                    Name ="Text198"
                    ControlSource ="=Sum([MEGH_MAR])"
                    FontName ="WeblogmaYekan"
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
' See "HAVLAH_ANBAR_khback.cls"
