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
    Width =10401
    DatasheetFontHeight =10
    ItemSuffix =158
    Left =930
    Top =1050
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x84b6c3d197a9e240
    End
    ServerFilter ="NUMBER =2"
    RecordSource ="QSL_RASID_MOSTAGHIM_LARGE"
    Caption ="برگ رسيد كالاي مصرف مستقيم"
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
            ControlSource ="NUMBER"
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =2115
            Name ="GroupHeader2"
            Begin
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =3712
                    Width =3120
                    Height =540
                    FontSize =14
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label88"
                    Caption ="برگ رسيد كالاي مصرف مستقيم"
                    FontName ="Titr"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2115
                    Top =495
                    Width =5613
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
                    Left =105
                    Top =540
                    Width =1311
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
                            Left =1453
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
                    Left =7755
                    Top =510
                    Width =1221
                    Height =435
                    FontSize =12
                    TabIndex =2
                    BorderColor =32768
                    Name ="NUMBER"
                    ControlSource ="NUMBER"
                    FontName ="Traffic"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =10
                            Left =9000
                            Top =510
                            Width =1305
                            Height =480
                            FontSize =12
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label93"
                            Caption ="شماره رسيد:"
                            FontName ="Traffic"
                        End
                    End
                End
                Begin Line
                    BorderWidth =3
                    Top =1050
                    Width =10401
                    BorderColor =0
                    Name ="Line94"
                End
                Begin Line
                    BorderWidth =3
                    Top =1590
                    Width =10401
                    BorderColor =0
                    Name ="Line99"
                End
                Begin Line
                    BorderWidth =3
                    Left =10380
                    Top =1605
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line130"
                End
                Begin Line
                    Left =9810
                    Top =1605
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line131"
                End
                Begin Line
                    Left =7707
                    Top =1605
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line132"
                End
                Begin Line
                    Left =3647
                    Top =1605
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line133"
                End
                Begin Line
                    Left =2822
                    Top =1605
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line134"
                End
                Begin Line
                    Left =1058
                    Top =1605
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line135"
                End
                Begin Line
                    BorderWidth =3
                    Left =15
                    Top =1605
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line138"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =203
                    Top =1650
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
                    Left =2994
                    Top =1650
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
                    Left =1235
                    Top =1650
                    Width =1560
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
                    Left =4460
                    Top =1647
                    Width =2520
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
                    Left =7782
                    Top =1650
                    Width =1890
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
                    Left =9816
                    Top =1646
                    Width =540
                    Height =405
                    FontSize =10
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label146"
                    Caption ="رديف"
                    FontName ="Traffic"
                End
                Begin Line
                    BorderWidth =3
                    Top =2115
                    Width =10401
                    BorderColor =0
                    Name ="Line147"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =170
                    Top =1077
                    Width =8781
                    Height =450
                    FontSize =12
                    TabIndex =3
                    BorderColor =32768
                    Name ="MOLAH"
                    ControlSource ="MOLAH"
                    FontName ="Traffic"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =10
                            Left =8970
                            Top =1080
                            Width =1395
                            Height =480
                            FontSize =12
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label155"
                            Caption ="تحويل دهنده:"
                            FontName ="Traffic"
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
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =7739
                    Top =16
                    Width =2019
                    Height =405
                    FontSize =10
                    FontWeight =700
                    BorderColor =32768
                    Name ="INVO_LST.ANBAR"
                    ControlSource ="NAMES"
                    FontName ="Traffic"
                    EventProcPrefix ="INVO_LST_ANBAR"
                End
                Begin TextBox
                    DecimalPlaces =0
                    RunningSum =1
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =9826
                    Top =16
                    Width =504
                    Height =405
                    FontSize =10
                    FontWeight =700
                    TabIndex =1
                    BorderColor =32768
                    Name ="RADIF"
                    ControlSource ="=1"
                    FontName ="Traffic"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3674
                    Top =16
                    Width =3969
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
                    Left =2847
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
                    Left =128
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
                    Left =1073
                    Top =15
                    Width =1710
                    Height =405
                    FontSize =10
                    FontWeight =700
                    TabIndex =5
                    BorderColor =32768
                    Name ="VAHED_K"
                    ControlSource ="VNAMES"
                    FontName ="Traffic"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Top =450
                    Width =10386
                    BorderColor =0
                    Name ="Line100"
                End
                Begin Line
                    BorderWidth =3
                    Left =10380
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line101"
                End
                Begin Line
                    BorderWidth =1
                    Left =9810
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line102"
                End
                Begin Line
                    BorderWidth =1
                    Left =7707
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line103"
                End
                Begin Line
                    BorderWidth =1
                    Left =3647
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line104"
                End
                Begin Line
                    BorderWidth =1
                    Left =2822
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line105"
                End
                Begin Line
                    BorderWidth =1
                    Left =1058
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line106"
                End
                Begin Line
                    BorderWidth =2
                    Left =15
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line109"
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =1247
            Name ="GroupFooter3"
            Begin
                Begin Line
                    BorderWidth =3
                    Left =15
                    Top =30
                    Width =10375
                    BorderColor =0
                    Name ="Line110"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =1050
                    Top =165
                    Width =2130
                    Height =405
                    FontSize =10
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label153"
                    Caption ="نام و نام خانوادگي انبار دار:"
                    FontName ="Traffic"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =540
                    Top =690
                    Width =525
                    Height =405
                    FontSize =10
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label154"
                    Caption ="امضاء"
                    FontName ="Traffic"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =7455
                    Top =165
                    Width =2595
                    Height =405
                    FontSize =10
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label156"
                    Caption ="نام و نام خانوادگي تحويل دهنده:"
                    FontName ="Traffic"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =7181
                    Top =695
                    Width =525
                    Height =405
                    FontSize =10
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label157"
                    Caption ="امضاء"
                    FontName ="Traffic"
                End
            End
        End
    End
End
CodeBehindForm
' See "RASID_MOSTAGHIM_LARG.cls"
