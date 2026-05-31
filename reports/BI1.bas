Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    DefaultView =0
    AllowUpdating =4
    TabularCharSet =178
    TabularFamily =42
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =10660
    DatasheetFontHeight =10
    ItemSuffix =198
    Left =930
    Top =1050
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x318f23107cede340
    End
    RecordSource ="QSL_BARGAH_ENTEGHAL"
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
        Begin BreakLevel
            GroupHeader = NotDefault
            GroupFooter = NotDefault
            ControlSource ="[NUMBER]"
        End
        Begin BreakLevel
            ControlSource ="RADIF"
        End
        Begin PageHeader
            Height =2475
            Name ="PageHeaderSection"
            Begin
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =4140
                    Width =1680
                    Height =540
                    FontSize =14
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label88"
                    Caption ="برگه انتقال كالا"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2265
                    Top =495
                    Width =5613
                    Height =540
                    FontSize =14
                    BorderColor =32768
                    Name ="Text90"
                    ControlSource ="=[Forms]![Baseknow]![WIDTH_D]"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextAlign =2
                            TextFontFamily =2
                            Left =8560
                            Top =1984
                            Width =1425
                            Height =390
                            FontSize =10
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label191"
                            Caption ="باركد"
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
                    Left =60
                    Top =540
                    Width =1506
                    Height =450
                    FontSize =12
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
                            Left =1603
                            Top =540
                            Width =660
                            Height =480
                            FontSize =12
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
                    Left =7545
                    Top =510
                    Width =1686
                    Height =435
                    FontSize =12
                    TabIndex =2
                    BorderColor =32768
                    Name ="NUMBER"
                    ControlSource ="NUMBER"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =9555
                            Top =510
                            Width =1005
                            Height =405
                            FontSize =12
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label93"
                            Caption ="شماره برگه:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin Line
                    BorderWidth =3
                    Top =1050
                    Width =10656
                    BorderColor =0
                    Name ="Line94"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =30
                    Top =1485
                    Width =4266
                    Height =375
                    FontSize =10
                    TabIndex =3
                    BorderColor =32768
                    Name ="HEAD_LST.CUST_NO"
                    ControlSource ="TAH"
                    FontName ="WeblogmaYekan"
                    EventProcPrefix ="HEAD_LST_CUST_NO"
                    Begin
                        Begin Label
                            TextFontFamily =2
                            Left =4362
                            Top =1485
                            Width =1170
                            Height =405
                            FontSize =10
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label95"
                            Caption ="تحويل گيرنده:"
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
                    Left =120
                    Top =1080
                    Width =4176
                    Height =375
                    FontSize =10
                    TabIndex =4
                    BorderColor =32768
                    Name ="ADDRESS"
                    ControlSource ="MOLAH"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextFontFamily =2
                            Left =4377
                            Top =1080
                            Width =1215
                            Height =405
                            FontSize =10
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label97"
                            Caption ="تحويل دهنده:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin Line
                    BorderWidth =3
                    Top =1950
                    Width =10656
                    BorderColor =0
                    Name ="Line99"
                End
                Begin Line
                    BorderWidth =3
                    Left =10635
                    Top =1965
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line130"
                End
                Begin Line
                    Left =10020
                    Top =1965
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line131"
                End
                Begin Line
                    Left =3499
                    Top =1965
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line133"
                End
                Begin Line
                    Left =2295
                    Top =1965
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line134"
                End
                Begin Line
                    Left =1290
                    Top =1965
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line135"
                End
                Begin Line
                    BorderWidth =3
                    Left =30
                    Top =1965
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line138"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =158
                    Top =2010
                    Width =1095
                    Height =405
                    FontSize =10
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label141"
                    Caption ="مقدار كل"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =1440
                    Top =2010
                    Width =765
                    Height =405
                    FontSize =10
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label142"
                    Caption ="مقدار"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =2539
                    Top =2010
                    Width =705
                    Height =375
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
                    Left =4200
                    Top =1980
                    Width =4275
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
                    Left =10083
                    Top =2006
                    Width =510
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
                    Top =2475
                    Width =10656
                    BorderColor =0
                    Name ="Line147"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5790
                    Top =1080
                    Width =4131
                    Height =375
                    FontSize =10
                    TabIndex =5
                    BorderColor =32768
                    Name ="Text153"
                    ControlSource ="ANBAR_NAMES"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =9972
                            Top =1080
                            Width =600
                            Height =405
                            FontSize =10
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label154"
                            Caption ="از انبار:"
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
                    Left =5835
                    Top =1503
                    Width =4086
                    Height =375
                    FontSize =10
                    TabIndex =6
                    BorderColor =32768
                    Name ="Text155"
                    ControlSource ="ANBAR_1_NAMES"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =9952
                            Top =1509
                            Width =630
                            Height =405
                            FontSize =10
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label156"
                            Caption ="به انبار:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =8050
                    Top =113
                    TabIndex =7
                    Name ="TAG"
                    ControlSource ="TAG"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    Left =8535
                    Top =1965
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line162"
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =0
            Name ="GroupHeader2"
        End
        Begin Section
            KeepTogether = NotDefault
            Height =465
            Name ="Detail"
            Begin
                Begin TextBox
                    DecimalPlaces =0
                    RunningSum =1
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =10063
                    Top =16
                    Width =504
                    Height =405
                    FontSize =10
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
                    Left =3544
                    Top =16
                    Width =4959
                    Height =405
                    FontSize =10
                    TabIndex =1
                    BorderColor =32768
                    Name ="kala"
                    ControlSource ="kala"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1293
                    Top =16
                    Width =1029
                    Height =405
                    FontSize =10
                    TabIndex =2
                    BorderColor =32768
                    Name ="MEGH"
                    ControlSource ="MEGH"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =83
                    Top =16
                    Width =1179
                    Height =405
                    FontSize =10
                    TabIndex =3
                    BorderColor =32768
                    Name ="MEGHk"
                    ControlSource ="MEGHk"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2295
                    Top =15
                    Width =1185
                    Height =405
                    FontSize =10
                    TabIndex =4
                    BorderColor =32768
                    Name ="VAHED_K"
                    ControlSource ="NAMES"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Top =450
                    Width =10641
                    BorderColor =0
                    Name ="Line100"
                End
                Begin Line
                    BorderWidth =2
                    Left =10617
                    Top =15
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line101"
                End
                Begin Line
                    BorderWidth =1
                    Left =10020
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line103"
                End
                Begin Line
                    BorderWidth =1
                    Left =3499
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line104"
                End
                Begin Line
                    BorderWidth =1
                    Left =2295
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line105"
                End
                Begin Line
                    BorderWidth =1
                    Left =1290
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
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =8610
                    Top =30
                    Width =1374
                    Height =390
                    FontSize =10
                    TabIndex =5
                    BorderColor =32768
                    Name ="N_FANI"
                    ControlSource ="N_FANI"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    BorderWidth =1
                    Left =8535
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line163"
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =4195
            Name ="GroupFooter3"
            Begin
                Begin Line
                    BorderWidth =3
                    Top =450
                    Width =10632
                    BorderColor =0
                    Name ="Line110"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5160
                    Top =1360
                    Width =3696
                    Height =375
                    FontSize =10
                    BorderColor =32768
                    Name ="Text157"
                    ControlSource ="MOLAH"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextFontFamily =2
                            Left =8865
                            Top =1365
                            Width =1695
                            Height =405
                            FontSize =10
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label158"
                            Caption =" امضاء تحويل دهنده:"
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
                    Left =120
                    Top =2325
                    Width =3801
                    Height =375
                    FontSize =10
                    TabIndex =1
                    BorderColor =32768
                    Name ="Text159"
                    ControlSource ="TAH"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextFontFamily =2
                            Left =3968
                            Top =2325
                            Width =1305
                            Height =405
                            FontSize =10
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label160"
                            Caption ="  تحويل گيرنده:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5272
                    Top =510
                    Width =2276
                    FontSize =8
                    TabIndex =2
                    Name ="Text172"
                    ControlSource ="=Format(Farsidate(Now()),\"00\\/00\\/00\") & \" - \" & Format(Now(),\"Long Time\""
                        ")"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1293
                    Top =16
                    Width =969
                    Height =405
                    FontSize =10
                    TabIndex =3
                    BorderColor =32768
                    Name ="Text164"
                    ControlSource ="=Sum([MEGH])"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =83
                    Top =16
                    Width =1179
                    Height =405
                    FontSize =10
                    TabIndex =4
                    BorderColor =32768
                    Name ="Text165"
                    ControlSource ="=Sum([MEGHk])"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    BorderWidth =1
                    Left =1290
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line167"
                End
                Begin Line
                    BorderWidth =2
                    Left =10617
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line169"
                End
                Begin Line
                    BorderWidth =1
                    Left =3499
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line174"
                End
                Begin Line
                    BorderWidth =1
                    Left =2295
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line175"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =7755
                    Top =510
                    Width =1761
                    Height =375
                    FontSize =10
                    TabIndex =5
                    BorderColor =32768
                    Name ="Text176"
                    ControlSource ="NUMBER"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextFontFamily =2
                            Left =9540
                            Top =510
                            Width =1050
                            Height =405
                            FontSize =10
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label177"
                            Caption ="شماره برگه:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =2
                    Left =585
                    Top =960
                    Width =10065
                    Height =375
                    FontSize =10
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label178"
                    Caption ="در صورت هرگونه مغايرت در محموله ارسالي در زمان تخليه و با حضور راننده اعلام شود "
                        "در غير اينصورت دفتر مركزي هيچگونه تعهدي ندارد."
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =2
                    Left =3911
                    Top =1360
                    Width =1140
                    Height =405
                    FontSize =10
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label179"
                    Caption ="شماره تماس:"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =2
                    Left =623
                    Top =1360
                    Width =1140
                    Height =405
                    FontSize =10
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label180"
                    Caption ="امضاء:"
                    FontName ="WeblogmaYekan"
                End
                Begin Rectangle
                    Left =5319
                    Top =1814
                    Width =5282
                    Height =2285
                    Name ="Box181"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =2
                    Left =5380
                    Top =1870
                    Width =5190
                    Height =315
                    FontSize =8
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label182"
                    Caption ="طبق صورت فوق تمام كالاها صحيح و سالم و بدون كسري تحويل اينجانب داده شده است:"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =2
                    Left =56
                    Top =1870
                    Width =5205
                    Height =315
                    FontSize =8
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label183"
                    Caption ="طبق صورت فوق تمام كالاها صحيح و سالم و بدون كسري تحويل اينجانب داده شده است:"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =2
                    Left =8235
                    Top =2265
                    Width =2355
                    Height =405
                    FontSize =10
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label184"
                    Caption =" امضاء تحويل گيرنده/ راننده:"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =2
                    Left =9435
                    Top =2835
                    Width =1140
                    Height =405
                    FontSize =10
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label185"
                    Caption ="شماره تماس:"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =2
                    Left =5839
                    Top =2834
                    Width =1140
                    Height =405
                    FontSize =10
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label186"
                    Caption ="امضاء:"
                    FontName ="WeblogmaYekan"
                End
                Begin Rectangle
                    Left =60
                    Top =1815
                    Width =5282
                    Height =2285
                    Name ="Box187"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =2
                    Left =4125
                    Top =2835
                    Width =1140
                    Height =405
                    FontSize =10
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label188"
                    Caption ="شماره تماس:"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =2
                    Left =396
                    Top =2834
                    Width =1140
                    Height =405
                    FontSize =10
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label189"
                    Caption ="امضاء:"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    BorderWidth =1
                    Left =56
                    Top =3401
                    Width =10545
                    BorderColor =0
                    Name ="Line190"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =2
                    Left =9600
                    Top =3480
                    Width =960
                    Height =345
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label192"
                    Caption ="شماره ماشين:"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =2
                    Left =7111
                    Top =3458
                    Width =1020
                    Height =345
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label193"
                    Caption ="شماره بارنامه :"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =2
                    Left =4478
                    Top =3458
                    Width =705
                    Height =345
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label194"
                    Caption ="توضيحات:"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextFontFamily =2
                    Left =2381
                    Width =1050
                    Height =405
                    FontSize =10
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label196"
                    Caption ="جمع كل:"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    BorderWidth =2
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line197"
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
' See "BI1.cls"
