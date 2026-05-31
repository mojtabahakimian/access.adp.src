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
    Width =10755
    DatasheetFontHeight =10
    ItemSuffix =388
    Left =930
    Top =1260
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xd255f5fc3b2ae440
    End
    ServerFilter ="TAG = 1 and NUMBER =2517"
    RecordSource ="RASID_ANBAR"
    Caption ="قبض باسكول"
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
            ControlSource ="RADIF"
        End
        Begin PageHeader
            Height =1386
            Name ="PageHeader"
            Begin
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =6802
                    Width =3873
                    Height =300
                    FontSize =7
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
                    Left =5602
                    Top =525
                    Width =696
                    Height =285
                    FontSize =6
                    TabIndex =1
                    BorderColor =32768
                    Name ="DATE_N"
                    ControlSource ="DATE_N"
                    FontName ="WeblogmaYekan"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            TextFontFamily =2
                            Left =6325
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
                    Left =5617
                    Top =240
                    Width =681
                    Height =270
                    FontSize =7
                    TabIndex =2
                    BorderColor =32768
                    Name ="NUMBER"
                    ControlSource ="NUMBER"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextFontFamily =2
                            Left =6323
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
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =7150
                    Top =810
                    Width =2916
                    Height =300
                    FontSize =7
                    TabIndex =3
                    BorderColor =32768
                    Name ="HEAD_LST.CUST_NO"
                    ControlSource ="HESAB"
                    FontName ="WeblogmaYekan"
                    EventProcPrefix ="HEAD_LST_CUST_NO"
                    Begin
                        Begin Label
                            TextFontFamily =2
                            Left =10075
                            Top =810
                            Width =630
                            Height =300
                            FontSize =7
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label95"
                            Caption ="فروشنده:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin Line
                    BorderWidth =1
                    Left =5587
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
                    Left =5605
                    Top =810
                    Width =1011
                    Height =300
                    FontSize =7
                    TabIndex =4
                    BorderColor =32768
                    Name ="TEL"
                    ControlSource ="TEL"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextFontFamily =2
                            Left =6655
                            Top =810
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
                Begin Line
                    BorderWidth =1
                    Left =6772
                    Top =15
                    Width =0
                    Height =805
                    BorderColor =0
                    Name ="Line219"
                End
                Begin Line
                    BorderWidth =1
                    Left =6772
                    Top =540
                    Width =3948
                    BorderColor =0
                    Name ="Line220"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =7975
                    Top =525
                    Width =1305
                    Height =300
                    FontSize =7
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label221"
                    Caption ="قبض باسكول كالاي وارده"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5602
                    Top =1066
                    Width =4671
                    Height =285
                    FontSize =7
                    TabIndex =5
                    BorderColor =32768
                    Name ="ADDRESS"
                    ControlSource ="ADDRESS"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextFontFamily =2
                            Left =10294
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
                    Left =5587
                    Width =5133
                    Height =1386
                    Name ="Box237"
                End
                Begin Label
                    TextFontFamily =2
                    Left =5587
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
                    Left =1215
                    Width =3873
                    Height =300
                    FontSize =7
                    TabIndex =6
                    BorderColor =32768
                    Name ="Text335"
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
                    Top =525
                    Width =696
                    Height =285
                    FontSize =6
                    TabIndex =7
                    BorderColor =32768
                    Name ="Text336"
                    ControlSource ="DATE_N"
                    FontName ="WeblogmaYekan"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            TextFontFamily =2
                            Left =738
                            Top =525
                            Width =420
                            Height =300
                            FontSize =7
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label337"
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
                    Left =30
                    Top =240
                    Width =681
                    Height =270
                    FontSize =7
                    TabIndex =8
                    BorderColor =32768
                    Name ="Text338"
                    ControlSource ="NUMBER"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextFontFamily =2
                            Left =736
                            Top =240
                            Width =495
                            Height =300
                            FontSize =7
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label339"
                            Caption ="شماره :"
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
                    Left =1563
                    Top =810
                    Width =2916
                    Height =300
                    FontSize =7
                    TabIndex =9
                    BorderColor =32768
                    Name ="Text340"
                    ControlSource ="HESAB"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextFontFamily =2
                            Left =4488
                            Top =810
                            Width =630
                            Height =300
                            FontSize =7
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label341"
                            Caption ="فروشنده:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin Line
                    BorderWidth =1
                    Top =810
                    Width =5133
                    BorderColor =0
                    Name ="Line342"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =18
                    Top =810
                    Width =1011
                    Height =300
                    FontSize =7
                    TabIndex =10
                    BorderColor =32768
                    Name ="Text343"
                    ControlSource ="TEL"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextFontFamily =2
                            Left =1068
                            Top =810
                            Width =375
                            Height =300
                            FontSize =7
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label344"
                            Caption ="تلفن:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin Line
                    BorderWidth =1
                    Left =1185
                    Top =15
                    Width =0
                    Height =805
                    BorderColor =0
                    Name ="Line349"
                End
                Begin Line
                    BorderWidth =1
                    Left =1185
                    Top =540
                    Width =3948
                    BorderColor =0
                    Name ="Line350"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =2388
                    Top =525
                    Width =1305
                    Height =300
                    FontSize =7
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label351"
                    Caption ="قبض باسكول كالاي وارده"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =15
                    Top =1066
                    Width =4671
                    Height =285
                    FontSize =7
                    TabIndex =11
                    BorderColor =32768
                    Name ="Text352"
                    ControlSource ="ADDRESS"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextFontFamily =2
                            Left =4707
                            Top =1066
                            Width =390
                            Height =300
                            FontSize =7
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label353"
                            Caption ="نشاني:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin Rectangle
                    Width =5133
                    Height =1386
                    Name ="Box354"
                End
                Begin Label
                    TextFontFamily =2
                    Width =990
                    Height =270
                    FontSize =7
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label355"
                    Caption ="(نسخه حسابداري)"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =6808
                    Top =255
                    Width =996
                    Height =285
                    FontSize =7
                    TabIndex =12
                    BorderColor =32768
                    Name ="ECODE"
                    ControlSource ="=[Forms]![Baseknow]![ECODE]"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextFontFamily =2
                            Left =7814
                            Top =255
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
                    Left =8995
                    Top =255
                    Width =1161
                    Height =285
                    FontSize =7
                    TabIndex =13
                    BorderColor =32768
                    Name ="Text208"
                    ControlSource ="=[Forms]![Baseknow]![MCODEM]"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextFontFamily =2
                            Left =10172
                            Top =255
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
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1230
                    Top =255
                    Width =997
                    Height =285
                    FontSize =7
                    TabIndex =14
                    BorderColor =32768
                    Name ="Text252"
                    ControlSource ="=[Forms]![Baseknow]![ECODE]"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextFontFamily =2
                            Left =2251
                            Top =255
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
                    Left =3435
                    Top =255
                    Width =1161
                    Height =285
                    FontSize =7
                    TabIndex =15
                    BorderColor =32768
                    Name ="Text254"
                    ControlSource ="=[Forms]![Baseknow]![MCODEM]"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextFontFamily =2
                            Left =4612
                            Top =255
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
            End
        End
        Begin Section
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =5848
            Name ="Detail"
            Begin
                Begin TextBox
                    OldBorderStyle =1
                    BorderWidth =1
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5655
                    Top =91
                    Width =3279
                    Height =405
                    FontSize =10
                    Name ="NAMEG"
                    ControlSource ="NAMEG"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    OldBorderStyle =1
                    BorderWidth =1
                    TextAlign =2
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =8955
                    Top =91
                    Width =681
                    Height =405
                    FontSize =10
                    TabIndex =1
                    Name ="CODE"
                    ControlSource ="CODEG"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    OldBorderStyle =1
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =8280
                    Top =615
                    Width =1341
                    Height =405
                    FontSize =10
                    TabIndex =2
                    BorderColor =32768
                    Name ="CAMIUN_NUM"
                    ControlSource ="CAMIUN_NUM"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            BackStyle =1
                            TextFontFamily =2
                            Left =9659
                            Top =615
                            Width =1051
                            Height =405
                            FontSize =10
                            FontWeight =400
                            BackColor =10092543
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label232"
                            Caption ="شماره ماشين:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =1
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5655
                    Top =615
                    Width =1626
                    Height =405
                    FontSize =10
                    TabIndex =3
                    BorderColor =32768
                    Name ="DRIVER"
                    ControlSource ="DRIVER"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            BackStyle =1
                            TextFontFamily =2
                            Left =7306
                            Top =615
                            Width =915
                            Height =405
                            FontSize =10
                            FontWeight =400
                            BackColor =10092543
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
                    OldBorderStyle =1
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =8265
                    Top =1791
                    Width =1371
                    Height =405
                    FontSize =10
                    TabIndex =4
                    BorderColor =32768
                    Name ="Text235"
                    ControlSource ="BARNAMEH"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            BackStyle =1
                            TextFontFamily =2
                            Left =9659
                            Top =1791
                            Width =1052
                            Height =405
                            FontSize =10
                            FontWeight =400
                            BackColor =10092543
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label236"
                            Caption ="شماره حواله:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin Label
                    BackStyle =1
                    TextFontFamily =2
                    Left =9659
                    Top =91
                    Width =1051
                    Height =405
                    FontSize =10
                    FontWeight =400
                    BackColor =10092543
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label318"
                    Caption ="كد/نام كالا:"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    OldBorderStyle =1
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =8280
                    Top =1230
                    Width =1341
                    Height =405
                    FontSize =10
                    TabIndex =5
                    BorderColor =32768
                    Name ="CAMIUN"
                    ControlSource ="CAMIUN"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            BackStyle =1
                            TextFontFamily =2
                            Left =9659
                            Top =1230
                            Width =1051
                            Height =405
                            FontSize =10
                            FontWeight =400
                            BackColor =10092543
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label320"
                            Caption ="نوع ماشين:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =1
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =8265
                    Top =2415
                    Width =1371
                    Height =405
                    FontSize =10
                    TabIndex =6
                    BorderColor =32768
                    Name ="VAZNH"
                    ControlSource ="VAZNH"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            BackStyle =1
                            TextFontFamily =2
                            Left =9659
                            Top =2415
                            Width =1052
                            Height =405
                            FontSize =10
                            FontWeight =400
                            BackColor =10092543
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label322"
                            Caption ="وزن حواله:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =1
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5655
                    Top =1230
                    Width =1626
                    Height =405
                    FontSize =10
                    TabIndex =7
                    BorderColor =32768
                    Name ="Text323"
                    ControlSource ="CAM_POORG"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            BackStyle =1
                            TextFontFamily =2
                            Left =7306
                            Top =1230
                            Width =915
                            Height =405
                            FontSize =10
                            FontWeight =400
                            BackColor =10092543
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label324"
                            Caption ="وزن پر:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =1
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5655
                    Top =1791
                    Width =1626
                    Height =405
                    FontSize =10
                    TabIndex =8
                    BorderColor =32768
                    Name ="Text325"
                    ControlSource ="CAM_KHALYG"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            BackStyle =1
                            TextFontFamily =2
                            Left =7306
                            Top =1791
                            Width =915
                            Height =405
                            FontSize =10
                            FontWeight =400
                            BackColor =10092543
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label326"
                            Caption ="وزن ماشين:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =1
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5655
                    Top =2415
                    Width =1626
                    Height =405
                    FontSize =10
                    TabIndex =9
                    BorderColor =32768
                    Name ="Text327"
                    ControlSource ="MEGHkG"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            BackStyle =1
                            TextFontFamily =2
                            Left =7306
                            Top =2415
                            Width =915
                            Height =405
                            FontSize =10
                            FontWeight =400
                            BackColor =10092543
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label328"
                            Caption ="وزن خالص:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =1
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5670
                    Top =3030
                    Width =3966
                    Height =1170
                    FontSize =10
                    TabIndex =10
                    BorderColor =32768
                    Name ="Text329"
                    ControlSource ="TOZIHG"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            BackStyle =1
                            TextFontFamily =2
                            Left =9659
                            Top =3030
                            Width =1052
                            Height =405
                            FontSize =10
                            FontWeight =400
                            BackColor =10092543
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label330"
                            Caption ="توضيحات:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =8959
                    Top =4561
                    Width =1671
                    Height =330
                    FontSize =8
                    TabIndex =11
                    BorderColor =32768
                    Name ="Text209"
                    ControlSource ="=UCURRENTUSER()"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =9070
                            Top =4279
                            Width =1455
                            Height =330
                            FontSize =8
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label210"
                            Caption ="مهر و امضاء صادر كننده:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =2
                    Left =6135
                    Top =4260
                    Width =1170
                    Height =330
                    FontSize =8
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label331"
                    Caption ="مهر و امضاء راننده:"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    BorderWidth =1
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5715
                    Top =4562
                    Width =2016
                    Height =330
                    FontSize =8
                    TabIndex =12
                    BorderColor =32768
                    Name ="Text332"
                    ControlSource ="DRIVER"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    Left =5595
                    Top =5415
                    Width =5130
                    BorderColor =0
                    Name ="Line125"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =7170
                    Top =5434
                    Width =3006
                    Height =300
                    FontSize =8
                    TabIndex =13
                    BorderColor =32768
                    Name ="Text199"
                    ControlSource ="=[Forms]![Baseknow]![TFADDRESS]"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =10215
                            Top =5434
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
                    Left =5670
                    Top =5434
                    Width =1086
                    Height =270
                    FontSize =7
                    TabIndex =14
                    BorderColor =32768
                    Name ="Text197"
                    ControlSource ="=[Forms]![Baseknow]![TFTEL]"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =6780
                            Top =5434
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
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =8157
                    Top =5145
                    Width =2529
                    FontSize =7
                    TabIndex =15
                    BorderColor =32768
                    Name ="Text211"
                    ControlSource ="=Format(Now(),\"hh:nn:ss ampm\")"
                    FontName ="WeblogmaYekan"
                End
                Begin Rectangle
                    Left =5587
                    Width =5133
                    Height =5811
                    Name ="Box334"
                End
                Begin TextBox
                    OldBorderStyle =1
                    BorderWidth =1
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =68
                    Top =91
                    Width =3279
                    Height =405
                    FontSize =10
                    TabIndex =16
                    Name ="Text356"
                    ControlSource ="NAMEG"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    OldBorderStyle =1
                    BorderWidth =1
                    TextAlign =2
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =3368
                    Top =91
                    Width =681
                    Height =405
                    FontSize =10
                    TabIndex =17
                    Name ="Text357"
                    ControlSource ="CODEG"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    OldBorderStyle =1
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2693
                    Top =615
                    Width =1341
                    Height =405
                    FontSize =10
                    TabIndex =18
                    BorderColor =32768
                    Name ="Text358"
                    ControlSource ="CAMIUN_NUM"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            BackStyle =1
                            TextFontFamily =2
                            Left =4072
                            Top =615
                            Width =1051
                            Height =405
                            FontSize =10
                            FontWeight =400
                            BackColor =10092543
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label359"
                            Caption ="شماره ماشين:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =1
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =68
                    Top =615
                    Width =1626
                    Height =405
                    FontSize =10
                    TabIndex =19
                    BorderColor =32768
                    Name ="Text360"
                    ControlSource ="DRIVER"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            BackStyle =1
                            TextFontFamily =2
                            Left =1719
                            Top =615
                            Width =915
                            Height =405
                            FontSize =10
                            FontWeight =400
                            BackColor =10092543
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label361"
                            Caption ="نام راننده:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =1
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2678
                    Top =1791
                    Width =1371
                    Height =405
                    FontSize =10
                    TabIndex =20
                    BorderColor =32768
                    Name ="Text362"
                    ControlSource ="BARNAMEH"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            BackStyle =1
                            TextFontFamily =2
                            Left =4072
                            Top =1791
                            Width =1052
                            Height =405
                            FontSize =10
                            FontWeight =400
                            BackColor =10092543
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label363"
                            Caption ="شماره حواله:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin Label
                    BackStyle =1
                    TextFontFamily =2
                    Left =4072
                    Top =91
                    Width =1051
                    Height =405
                    FontSize =10
                    FontWeight =400
                    BackColor =10092543
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label364"
                    Caption ="كد/نام كالا:"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    OldBorderStyle =1
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2693
                    Top =1230
                    Width =1341
                    Height =405
                    FontSize =10
                    TabIndex =21
                    BorderColor =32768
                    Name ="Text365"
                    ControlSource ="CAMIUN"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            BackStyle =1
                            TextFontFamily =2
                            Left =4072
                            Top =1230
                            Width =1051
                            Height =405
                            FontSize =10
                            FontWeight =400
                            BackColor =10092543
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label366"
                            Caption ="نوع ماشين:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =1
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2678
                    Top =2415
                    Width =1371
                    Height =405
                    FontSize =10
                    TabIndex =22
                    BorderColor =32768
                    Name ="Text367"
                    ControlSource ="VAZNH"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            BackStyle =1
                            TextFontFamily =2
                            Left =4072
                            Top =2415
                            Width =1052
                            Height =405
                            FontSize =10
                            FontWeight =400
                            BackColor =10092543
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label368"
                            Caption ="وزن حواله:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =1
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =68
                    Top =1230
                    Width =1626
                    Height =405
                    FontSize =10
                    TabIndex =23
                    BorderColor =32768
                    Name ="Text369"
                    ControlSource ="CAM_POORG"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            BackStyle =1
                            TextFontFamily =2
                            Left =1719
                            Top =1230
                            Width =915
                            Height =405
                            FontSize =10
                            FontWeight =400
                            BackColor =10092543
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label370"
                            Caption ="وزن پر:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =1
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =68
                    Top =1791
                    Width =1626
                    Height =405
                    FontSize =10
                    TabIndex =24
                    BorderColor =32768
                    Name ="Text371"
                    ControlSource ="CAM_KHALYG"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            BackStyle =1
                            TextFontFamily =2
                            Left =1719
                            Top =1791
                            Width =915
                            Height =405
                            FontSize =10
                            FontWeight =400
                            BackColor =10092543
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label372"
                            Caption ="وزن ماشين:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =1
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =68
                    Top =2415
                    Width =1626
                    Height =405
                    FontSize =10
                    TabIndex =25
                    BorderColor =32768
                    Name ="Text373"
                    ControlSource ="MEGHkG"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            BackStyle =1
                            TextFontFamily =2
                            Left =1719
                            Top =2415
                            Width =915
                            Height =405
                            FontSize =10
                            FontWeight =400
                            BackColor =10092543
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label374"
                            Caption ="وزن خالص:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =1
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =83
                    Top =3030
                    Width =3966
                    Height =1170
                    FontSize =10
                    TabIndex =26
                    BorderColor =32768
                    Name ="Text375"
                    ControlSource ="TOZIHG"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            BackStyle =1
                            TextFontFamily =2
                            Left =4072
                            Top =3030
                            Width =1052
                            Height =405
                            FontSize =10
                            FontWeight =400
                            BackColor =10092543
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label376"
                            Caption ="توضيحات:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3372
                    Top =4561
                    Width =1671
                    Height =330
                    FontSize =8
                    TabIndex =27
                    BorderColor =32768
                    Name ="Text377"
                    ControlSource ="=UCURRENTUSER()"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =3483
                            Top =4279
                            Width =1455
                            Height =330
                            FontSize =8
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label378"
                            Caption ="مهر و امضاء صادر كننده:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =2
                    Left =548
                    Top =4260
                    Width =1170
                    Height =330
                    FontSize =8
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label379"
                    Caption ="مهر و امضاء راننده:"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    BorderWidth =1
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =128
                    Top =4562
                    Width =2016
                    Height =330
                    FontSize =8
                    TabIndex =28
                    BorderColor =32768
                    Name ="Text380"
                    ControlSource ="DRIVER"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    Left =8
                    Top =5415
                    Width =5130
                    BorderColor =0
                    Name ="Line381"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1583
                    Top =5434
                    Width =3006
                    Height =300
                    FontSize =8
                    TabIndex =29
                    BorderColor =32768
                    Name ="Text382"
                    ControlSource ="=[Forms]![Baseknow]![TFADDRESS]"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =4628
                            Top =5434
                            Width =495
                            Height =300
                            FontSize =7
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label383"
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
                    Left =83
                    Top =5434
                    Width =1086
                    Height =270
                    FontSize =7
                    TabIndex =30
                    BorderColor =32768
                    Name ="Text384"
                    ControlSource ="=[Forms]![Baseknow]![TFTEL]"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =1193
                            Top =5434
                            Width =375
                            Height =300
                            FontSize =7
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label385"
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
                    Left =2570
                    Top =5145
                    Width =2529
                    FontSize =7
                    TabIndex =31
                    BorderColor =32768
                    Name ="Text386"
                    ControlSource ="=Format(Now(),\"hh:nn:ss ampm\")"
                    FontName ="WeblogmaYekan"
                End
                Begin Rectangle
                    Width =5133
                    Height =5811
                    Name ="Box387"
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
' See "IK2.cls"
