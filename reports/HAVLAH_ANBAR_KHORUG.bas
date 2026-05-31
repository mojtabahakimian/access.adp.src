Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    FilterOn = NotDefault
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
    Width =10601
    DatasheetFontHeight =10
    ItemSuffix =196
    Left =600
    Top =1935
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x18f666671ba9e240
    End
    RecordSource ="QSL_HAVALEH_KHUROG"
    Caption ="حواله خروج كالا از انبار"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial (Arabic)"
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
            SortOrder = NotDefault
            GroupHeader = NotDefault
            GroupFooter = NotDefault
            KeepTogether =1
            ControlSource ="[NUMBER] & \", \" &[TAG]"
        End
        Begin BreakLevel
            ControlSource ="SNAME"
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =3045
            Name ="GroupHeader2"
            Begin
                Begin Label
                    TextFontFamily =10
                    Left =3432
                    Width =2445
                    Height =540
                    FontSize =14
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label88"
                    Caption ="حواله خروج كالا از انبار"
                    FontName ="Titr"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2757
                    Top =495
                    Width =4698
                    Height =540
                    FontSize =14
                    BorderColor =32768
                    Name ="Text90"
                    ControlSource ="=[Forms]![Baseknow]![WIDTH_D]"
                    FontName ="Titr"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =372
                    Top =540
                    Width =1596
                    Height =450
                    FontSize =12
                    TabIndex =1
                    BorderColor =32768
                    Name ="DATE_N"
                    ControlSource ="DATE_N"
                    FontName ="Traffic"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =10
                            Left =2005
                            Top =540
                            Width =660
                            Height =480
                            FontSize =12
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label92"
                            Caption ="تاريخ :"
                            FontName ="Traffic"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =7452
                    Top =510
                    Width =1686
                    Height =435
                    FontSize =12
                    TabIndex =2
                    BorderColor =32768
                    Name ="HEAD_LST.NUMBER"
                    ControlSource ="NUMBER"
                    FontName ="Traffic"
                    EventProcPrefix ="HEAD_LST_NUMBER"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =10
                            Left =9207
                            Top =510
                            Width =1260
                            Height =480
                            FontSize =12
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label93"
                            Caption ="شماره حواله:"
                            FontName ="Traffic"
                        End
                    End
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =3
                    Left =87
                    Top =1065
                    Width =10507
                    BorderColor =0
                    Name ="Line94"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2757
                    Top =1107
                    Width =6921
                    Height =450
                    FontSize =12
                    TabIndex =3
                    BorderColor =32768
                    Name ="HEAD_LST.CUST_NO"
                    ControlSource ="NAME"
                    FontName ="Traffic"
                    EventProcPrefix ="HEAD_LST_CUST_NO"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =10
                            Left =9714
                            Top =1107
                            Width =780
                            Height =480
                            FontSize =12
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label95"
                            Caption ="خريدار:"
                            FontName ="Traffic"
                        End
                    End
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =327
                    Top =1095
                    Width =1926
                    Height =450
                    FontSize =12
                    TabIndex =4
                    BorderColor =32768
                    Name ="hes"
                    ControlSource ="hes"
                    FontName ="Traffic"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =10
                            Left =2244
                            Top =1095
                            Width =405
                            Height =480
                            FontSize =12
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label96"
                            Caption ="كد:"
                            FontName ="Traffic"
                        End
                    End
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2532
                    Top =1569
                    Width =7131
                    Height =450
                    FontSize =12
                    TabIndex =5
                    BorderColor =32768
                    Name ="ADDRESS"
                    ControlSource ="ADDRESS"
                    FontName ="Traffic"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =10
                            Left =9699
                            Top =1569
                            Width =765
                            Height =480
                            FontSize =12
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label97"
                            Caption ="آدرس:"
                            FontName ="Traffic"
                        End
                    End
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =312
                    Top =1560
                    Width =1596
                    Height =450
                    FontSize =12
                    TabIndex =6
                    BorderColor =32768
                    Name ="TEL"
                    ControlSource ="TEL"
                    FontName ="Traffic"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =10
                            Left =1929
                            Top =1560
                            Width =555
                            Height =480
                            FontSize =12
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label98"
                            Caption ="تلفن:"
                            FontName ="Traffic"
                        End
                    End
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =3
                    Left =57
                    Top =2520
                    Width =10507
                    BorderColor =0
                    Name ="Line99"
                End
                Begin Line
                    BorderWidth =3
                    Left =10542
                    Top =2535
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line130"
                End
                Begin Line
                    Left =9972
                    Top =2535
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line131"
                End
                Begin Line
                    Left =7782
                    Top =2535
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line132"
                End
                Begin Line
                    Left =3162
                    Top =2535
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line133"
                End
                Begin Line
                    Left =2332
                    Top =2535
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line134"
                End
                Begin Line
                    Left =1092
                    Top =2535
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line135"
                End
                Begin Line
                    BorderWidth =3
                    Left =87
                    Top =2535
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line138"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =237
                    Top =2580
                    Width =810
                    Height =405
                    FontSize =10
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label141"
                    Caption ="مقدار كل"
                    FontName ="Traffic"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =2509
                    Top =2580
                    Width =525
                    Height =405
                    FontSize =10
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label142"
                    Caption ="مقدار"
                    FontName ="Traffic"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =1287
                    Top =2580
                    Width =810
                    Height =405
                    FontSize =10
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label143"
                    Caption ="واحد كالا"
                    FontName ="Traffic"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =4264
                    Top =2580
                    Width =1710
                    Height =405
                    FontSize =10
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label144"
                    Caption ="شــــــــــرح  كالا"
                    FontName ="Traffic"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =7857
                    Top =2580
                    Width =1980
                    Height =405
                    FontSize =10
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label145"
                    Caption ="انبار"
                    FontName ="Traffic"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =9978
                    Top =2576
                    Width =540
                    Height =405
                    FontSize =10
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label146"
                    Caption ="رديف"
                    FontName ="Traffic"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =267
                    Top =2040
                    Width =9186
                    Height =450
                    FontSize =12
                    TabIndex =7
                    BorderColor =32768
                    Name ="MOLAH"
                    ControlSource ="MOLAH"
                    FontName ="Traffic"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =10
                            Left =9477
                            Top =2040
                            Width =1035
                            Height =480
                            FontSize =12
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label155"
                            Caption ="ملاحظات:"
                            FontName ="Traffic"
                        End
                    End
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =3
                    Left =87
                    Top =3045
                    Width =10479
                    BorderColor =0
                    Name ="Line147"
                End
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Left =510
                    Top =56
                    TabIndex =8
                    Name ="TAG"
                    ControlSource ="TAG"
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
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =7814
                    Top =16
                    Width =2109
                    Height =405
                    FontSize =10
                    FontWeight =700
                    BorderColor =32768
                    Name ="NAMES"
                    ControlSource ="NAMES"
                    FontName ="Traffic"
                End
                Begin TextBox
                    DecimalPlaces =0
                    RunningSum =1
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =9988
                    Top =16
                    Width =504
                    Height =405
                    FontSize =10
                    FontWeight =700
                    TabIndex =1
                    BorderColor =32768
                    Name ="RADIF123"
                    ControlSource ="=1"
                    FontName ="Traffic"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3174
                    Top =16
                    Width =4569
                    Height =405
                    FontSize =10
                    FontWeight =700
                    TabIndex =2
                    BorderColor =32768
                    Name ="CODE"
                    ControlSource ="SNAME"
                    FontName ="Traffic"
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =2
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2362
                    Top =16
                    Width =789
                    Height =405
                    FontSize =10
                    FontWeight =700
                    TabIndex =3
                    BorderColor =32768
                    Name ="MEGH"
                    ControlSource ="MEGH"
                    FontName ="Traffic"
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =2
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =162
                    Top =16
                    Width =894
                    Height =405
                    FontSize =10
                    FontWeight =700
                    TabIndex =4
                    BorderColor =32768
                    Name ="MEGHk"
                    ControlSource ="MEGHk"
                    FontName ="Traffic"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1092
                    Top =15
                    Width =1215
                    Height =405
                    FontSize =10
                    FontWeight =700
                    TabIndex =5
                    BorderColor =32768
                    Name ="VAHED_K"
                    ControlSource ="vNAMES"
                    FontName ="Traffic"
                End
                Begin Line
                    BorderWidth =1
                    Left =87
                    Top =420
                    Width =10461
                    BorderColor =0
                    Name ="Line100"
                End
                Begin Line
                    BorderWidth =3
                    Left =10542
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line101"
                End
                Begin Line
                    BorderWidth =1
                    Left =9972
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line102"
                End
                Begin Line
                    BorderWidth =1
                    Left =7782
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line103"
                End
                Begin Line
                    BorderWidth =1
                    Left =3162
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line104"
                End
                Begin Line
                    BorderWidth =1
                    Left =2332
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line105"
                End
                Begin Line
                    BorderWidth =1
                    Left =1092
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line106"
                End
                Begin Line
                    BorderWidth =2
                    Left =57
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line109"
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =1190
            Name ="GroupFooter3"
            Begin
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =3
                    Left =57
                    Width =10507
                    BorderColor =0
                    Name ="Line110"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =1520
                    Top =225
                    Width =2085
                    Height =405
                    FontSize =10
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label191"
                    Caption ="نام ونام خانوادگي انبار دار:"
                    FontName ="Traffic"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =7550
                    Top =225
                    Width =2625
                    Height =405
                    FontSize =10
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label192"
                    Caption ="نام و نام خانوادگي تحويل گيرنده:"
                    FontName ="Traffic"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =6597
                    Top =675
                    Width =570
                    Height =405
                    FontSize =10
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label193"
                    Caption ="امضاء "
                    FontName ="Traffic"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =784
                    Top =750
                    Width =570
                    Height =405
                    FontSize =10
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label194"
                    Caption ="امضاء "
                    FontName ="Traffic"
                End
            End
        End
    End
End
CodeBehindForm
' See "HAVLAH_ANBAR_KHORUG.cls"
