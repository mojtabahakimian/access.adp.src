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
    GridX =50
    GridY =50
    Width =10680
    DatasheetFontHeight =10
    ItemSuffix =174
    Left =525
    Top =660
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x4218033c91b8e240
    End
    RecordSource ="RASID_ANBAR"
    Caption ="فاكتور خريد"
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
            Height =2085
            Name ="PageHeaderSection"
            Begin
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =3738
                    Width =2340
                    Height =540
                    FontSize =14
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label88"
                    Caption ="رسيد انبار"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2163
                    Top =495
                    Width =5388
                    Height =540
                    FontSize =14
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
                    Left =48
                    Top =540
                    Width =1386
                    Height =405
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
                            Left =1471
                            Top =540
                            Width =660
                            Height =405
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
                    Left =7548
                    Top =510
                    Width =1686
                    Height =405
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
                            Left =9258
                            Top =510
                            Width =1305
                            Height =405
                            FontSize =12
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label93"
                            Caption ="شماره رسيد:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin Line
                    BorderWidth =3
                    Top =1050
                    Width =10660
                    BorderColor =0
                    Name ="Line94"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =3
                    Top =1545
                    Width =10660
                    BorderColor =0
                    Name ="Line99"
                End
                Begin Line
                    BorderWidth =3
                    Left =10638
                    Top =1575
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line130"
                End
                Begin Line
                    Left =10068
                    Top =1575
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line131"
                End
                Begin Line
                    Left =7965
                    Top =1575
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line132"
                End
                Begin Line
                    Left =4362
                    Top =1575
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line133"
                End
                Begin Line
                    Left =3537
                    Top =1575
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line134"
                End
                Begin Line
                    Left =2148
                    Top =1575
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line135"
                End
                Begin Line
                    BorderWidth =2
                    Top =1575
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line138"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =1203
                    Top =1620
                    Width =915
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
                    Left =3574
                    Top =1620
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
                    Left =2160
                    Top =1620
                    Width =1350
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
                    Left =4473
                    Top =1620
                    Width =3285
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
                    Left =8040
                    Top =1620
                    Width =1890
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
                    Left =10089
                    Top =1616
                    Width =495
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
                    LineSlant = NotDefault
                    BorderWidth =3
                    Top =2085
                    Width =10671
                    BorderColor =0
                    Name ="Line147"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =11
                    Top =1095
                    Width =3531
                    Height =405
                    FontSize =11
                    TabIndex =3
                    BorderColor =32768
                    Name ="TAH"
                    ControlSource ="TAH"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =3581
                            Top =1095
                            Width =1440
                            Height =405
                            FontSize =11
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label155"
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
                    Left =5104
                    Top =1095
                    Width =4206
                    Height =405
                    FontSize =11
                    TabIndex =4
                    BorderColor =32768
                    Name ="Text158"
                    ControlSource ="MOLAH"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =9318
                            Top =1102
                            Width =1305
                            Height =405
                            FontSize =11
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label159"
                            Caption ="تحويل دهنده:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin Line
                    Left =1158
                    Top =1575
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line162"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =33
                    Top =1635
                    Width =1125
                    Height =405
                    FontSize =10
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label163"
                    Caption ="مقدار تحويلي"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Top =237
                    Width =2336
                    Height =300
                    FontSize =10
                    TabIndex =5
                    Name ="Text24"
                    ControlSource ="=\"صفحه \" & [Page] & \" از \" & [Pages]"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    Left =7545
                    Top =120
                    Width =2226
                    Height =420
                    FontSize =12
                    TabIndex =6
                    BorderColor =32768
                    Name ="Text170"
                    ControlSource ="=[Forms]![BASEKNOW]![ISO_KHAREED]"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =9735
                            Top =75
                            Width =825
                            Height =465
                            FontSize =12
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label171"
                            Caption ="كد فرم :"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =405
            Name ="Detail"
            Begin
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =7986
                    Top =16
                    Width =2019
                    Height =360
                    FontSize =8
                    BorderColor =32768
                    Name ="NAMES"
                    ControlSource ="NAMES"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =0
                    RunningSum =1
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =10073
                    Top =16
                    Width =504
                    Height =360
                    FontSize =8
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
                    Left =4356
                    Top =16
                    Width =3534
                    Height =360
                    FontSize =8
                    TabIndex =2
                    BorderColor =32768
                    Name ="CODE"
                    ControlSource ="KALA"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3551
                    Top =16
                    Width =789
                    Height =360
                    FontSize =8
                    TabIndex =3
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
                    Left =1207
                    Top =16
                    Width =894
                    Height =360
                    FontSize =8
                    TabIndex =4
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
                    Left =2167
                    Top =15
                    Width =1320
                    Height =360
                    FontSize =8
                    TabIndex =5
                    BorderColor =32768
                    Name ="VAHED_K"
                    ControlSource ="VNAMES"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Top =405
                    Width =10660
                    BorderColor =0
                    Name ="Line100"
                End
                Begin Line
                    BorderWidth =3
                    Left =10638
                    Width =0
                    Height =405
                    BorderColor =0
                    Name ="Line101"
                End
                Begin Line
                    BorderWidth =1
                    Left =10069
                    Width =0
                    Height =405
                    BorderColor =0
                    Name ="Line102"
                End
                Begin Line
                    BorderWidth =1
                    Left =7966
                    Width =0
                    Height =405
                    BorderColor =0
                    Name ="Line103"
                End
                Begin Line
                    BorderWidth =1
                    Left =4362
                    Width =0
                    Height =405
                    BorderColor =0
                    Name ="Line104"
                End
                Begin Line
                    BorderWidth =1
                    Left =3537
                    Width =0
                    Height =405
                    BorderColor =0
                    Name ="Line105"
                End
                Begin Line
                    BorderWidth =1
                    Left =2145
                    Width =0
                    Height =405
                    BorderColor =0
                    Name ="Line106"
                End
                Begin Line
                    BorderWidth =2
                    Width =0
                    Height =405
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
                    Left =22
                    Width =1089
                    Height =360
                    FontSize =8
                    TabIndex =6
                    BorderColor =32768
                    Name ="MEGH_R"
                    ControlSource ="MEGH_R"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    BorderWidth =1
                    Left =1155
                    Width =0
                    Height =405
                    BorderColor =0
                    Name ="Line165"
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =1414
            Name ="GroupFooter3"
            Begin
                Begin Line
                    BorderWidth =3
                    Width =10660
                    BorderColor =0
                    Name ="Line110"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =2151
                    Top =565
                    Width =525
                    Height =405
                    FontSize =10
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label154"
                    Caption ="امضاء"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =7939
                    Top =566
                    Width =525
                    Height =405
                    FontSize =10
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label157"
                    Caption ="امضاء"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =247
                    Top =165
                    Width =4371
                    Height =375
                    FontSize =10
                    BorderColor =32768
                    Name ="Text166"
                    ControlSource ="MOLAH"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5992
                    Top =165
                    Width =4371
                    Height =375
                    FontSize =10
                    TabIndex =1
                    BorderColor =32768
                    Name ="Text168"
                    ControlSource ="TAH"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =8401
                    Top =1001
                    Width =2276
                    FontSize =8
                    TabIndex =2
                    Name ="Text172"
                    ControlSource ="=Format(Farsidate(Now()),\"00\\/00\\/00\") & \" - \" & Format(Now(),\"Long Time\""
                        ")"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Top =1174
                    Width =2276
                    FontSize =8
                    TabIndex =3
                    Name ="Text173"
                    ControlSource ="=Format(Farsidate(Now()),\"00\\/00\\/00\") & \" - \" & Format(Now(),\"Long Time\""
                        ")"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextFontFamily =2
                    Left =45
                    Top =900
                    Width =3585
                    Height =345
                    FontSize =10
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="vazn"
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
' See "HAVALAH_ANVAR_VROUD.cls"
