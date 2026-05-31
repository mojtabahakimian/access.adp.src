Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    FilterOn = NotDefault
    DefaultView =0
    AllowUpdating =4
    TabularCharSet =178
    TabularFamily =42
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =60
    GridY =60
    Width =4071
    DatasheetFontHeight =10
    ItemSuffix =241
    Left =270
    Top =210
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x911d697514c9e440
    End
    RecordSource ="QSL_INVOICE_FROOSH"
    Caption ="فاكتور فروش"
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
            Height =1487
            Name ="PageHeader"
            Begin
                Begin Label
                    TextFontFamily =10
                    Left =1590
                    Width =825
                    Height =330
                    FontSize =8
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label88"
                    Caption ="فاكتور فروش"
                    FontName ="Titr"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1035
                    Top =300
                    Width =1953
                    Height =330
                    FontSize =8
                    BorderColor =32768
                    Name ="Text90"
                    ControlSource ="=[Forms]![Baseknow]![WIDTH_D]"
                    FontName ="Titr"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =15
                    Width =681
                    Height =300
                    FontSize =8
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
                            Left =718
                            Width =390
                            Height =330
                            FontSize =8
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
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2430
                    Width =786
                    Height =300
                    FontSize =8
                    TabIndex =2
                    BorderColor =32768
                    Name ="NUMBER1"
                    ControlSource ="NUMBER"
                    FontName ="Traffic"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =3222
                            Width =825
                            Height =315
                            FontSize =8
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label93"
                            Caption ="شماره فاكتور:"
                            FontName ="Traffic"
                        End
                    End
                End
                Begin TextBox
                    TextFontCharSet =0
                    TextAlign =3
                    TextFontFamily =34
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1376
                    Top =690
                    Width =2076
                    Height =300
                    FontSize =8
                    TabIndex =3
                    BorderColor =32768
                    Name ="HEAD_LST.CUST_NO"
                    ControlSource ="HESAB"
                    FontName ="Tahoma"
                    EventProcPrefix ="HEAD_LST_CUST_NO"
                    Begin
                        Begin Label
                            TextFontFamily =10
                            Left =3458
                            Top =690
                            Width =585
                            Height =330
                            FontSize =8
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label95"
                            Caption ="خريدار:"
                            FontName ="Traffic"
                        End
                    End
                End
                Begin Line
                    BorderWidth =2
                    Top =660
                    Width =4071
                    BorderColor =0
                    Name ="Line164"
                End
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Left =105
                    Top =180
                    Width =291
                    Height =255
                    TabIndex =4
                    Name ="TAG"
                    ControlSource ="TAG"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =2
                    Top =1065
                    Width =4045
                    BorderColor =0
                    Name ="Line230"
                End
                Begin Line
                    BorderWidth =2
                    Left =4036
                    Top =1095
                    Width =0
                    Height =391
                    BorderColor =0
                    Name ="Line130"
                End
                Begin Line
                    BorderWidth =1
                    Left =1485
                    Top =1096
                    Width =0
                    Height =391
                    BorderColor =0
                    Name ="Line136"
                End
                Begin Line
                    BorderWidth =1
                    Left =865
                    Top =1095
                    Width =0
                    Height =391
                    BorderColor =0
                    Name ="Line137"
                End
                Begin Line
                    BorderWidth =2
                    Top =1095
                    Width =0
                    Height =391
                    BorderColor =0
                    Name ="Line138"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =15
                    Top =1125
                    Width =795
                    Height =345
                    FontSize =8
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label139"
                    Caption ="مبــلغ"
                    FontName ="Titr"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =910
                    Top =1155
                    Width =555
                    Height =330
                    FontSize =8
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label140"
                    Caption ="فـي"
                    FontName ="Titr"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =1535
                    Top =1155
                    Width =435
                    Height =330
                    FontSize =8
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label141"
                    Caption ="مقدار "
                    FontName ="Titr"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =2000
                    Top =1139
                    Width =2010
                    Height =330
                    FontSize =8
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label144"
                    Caption ="شــــــــــرح  كالا"
                    FontName ="Titr"
                End
                Begin Line
                    BorderWidth =1
                    Left =1970
                    Top =1095
                    Width =0
                    Height =391
                    BorderColor =0
                    Name ="Line205"
                End
                Begin Line
                    BorderWidth =2
                    Top =1470
                    Width =4060
                    BorderColor =0
                    Name ="Line2"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3029
                    Top =349
                    Width =1011
                    Height =300
                    FontSize =8
                    TabIndex =5
                    BorderColor =32768
                    Name ="NUMBER"
                    ControlSource ="NUMBER"
                    FontName ="Traffic"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Top =690
                    Width =996
                    Height =300
                    FontSize =8
                    TabIndex =6
                    BorderColor =32768
                    Name ="DEPNAME"
                    ControlSource ="DEPNAME"
                    FontName ="Traffic"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =10
                            Left =1020
                            Top =690
                            Width =345
                            Height =330
                            FontSize =8
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label206"
                            Caption ="واحد:"
                            FontName ="Traffic"
                        End
                    End
                End
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Top =330
                    Width =999
                    Height =270
                    FontSize =8
                    TabIndex =7
                    BorderColor =32768
                    Name ="Text211"
                    ControlSource ="=Format(Now(),\"hh:nn:ss ampm\")"
                    FontName ="Zar"
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
                    Left =3780
                    Top =15
                    Width =234
                    Height =330
                    FontSize =8
                    FontWeight =700
                    BorderColor =32768
                    Name ="RADIFff"
                    ControlSource ="=1"
                    FontName ="Titr Mazar"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1980
                    Top =16
                    Width =1794
                    Height =330
                    FontSize =8
                    FontWeight =700
                    TabIndex =1
                    BorderColor =32768
                    Name ="Text218"
                    ControlSource ="KALA"
                    FontName ="Titr Mazar"
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1500
                    Top =15
                    Width =459
                    Height =330
                    FontSize =8
                    TabIndex =2
                    BorderColor =32768
                    Name ="Text219"
                    ControlSource ="MEGHk"
                    FontName ="Traffic"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =885
                    Top =15
                    Width =564
                    Height =330
                    FontSize =8
                    TabIndex =3
                    BorderColor =32768
                    Name ="Text220"
                    ControlSource ="MABL"
                    Format ="#,###"
                    FontName ="Traffic"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =45
                    Top =15
                    Width =789
                    Height =330
                    FontSize =8
                    TabIndex =4
                    BorderColor =32768
                    Name ="Text221"
                    ControlSource ="MABL_K"
                    Format ="#,###"
                    FontName ="Traffic"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =2
                    Left =4036
                    Width =0
                    Height =375
                    BorderColor =0
                    Name ="Line222"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Left =3811
                    Width =0
                    Height =375
                    BorderColor =0
                    Name ="Line223"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Left =1485
                    Width =0
                    Height =375
                    BorderColor =0
                    Name ="Line225"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Left =865
                    Width =0
                    Height =375
                    BorderColor =0
                    Name ="Line226"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =2
                    Width =0
                    Height =375
                    BorderColor =0
                    Name ="Line227"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Left =1970
                    Width =0
                    Height =375
                    BorderColor =0
                    Name ="Line208"
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =675
            Name ="GroupFooter3"
            Begin
                Begin Line
                    BorderWidth =2
                    Width =0
                    Height =651
                    BorderColor =0
                    Name ="Line119"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =2
                    Top =661
                    Width =4060
                    BorderColor =0
                    Name ="Line120"
                End
                Begin Line
                    BorderWidth =2
                    Left =4036
                    Width =0
                    Height =646
                    BorderColor =0
                    Name ="Line127"
                End
                Begin Line
                    LineSlant = NotDefault
                    Top =645
                    Width =4053
                    BorderColor =0
                    Name ="Line128"
                End
                Begin Label
                    TextAlign =0
                    TextFontFamily =2
                    Left =1185
                    Top =90
                    Width =675
                    Height =420
                    FontSize =11
                    FontWeight =400
                    ForeColor =0
                    Name ="Label178"
                    Caption ="جمع کل "
                    FontName ="Zar"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    NumeralShapes =2
                    Top =120
                    Width =1170
                    Height =360
                    FontSize =10
                    ForeColor =0
                    Name ="JF"
                    Caption ="جمع كل فاكتور:"
                    FontName ="Titr Mazar"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =2
                    Left =1860
                    Width =2130
                    Height =675
                    FontSize =11
                    BorderColor =32768
                    ForeColor =0
                    Name ="HR"
                    Caption ="مبلغ به حروف:"
                    FontName ="Kamran"
                End
                Begin Line
                    LineSlant = NotDefault
                    Width =4053
                    BorderColor =0
                    Name ="Line235"
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
' See "i23.cls"
