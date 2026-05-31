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
    Width =3959
    DatasheetFontHeight =10
    ItemSuffix =258
    Left =915
    Top =1035
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x193fdc3f7492e540
    End
    RecordSource ="qsl_invoice_froosh_y"
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
            ControlSource ="RADAH"
        End
        Begin PageHeader
            Height =1187
            Name ="PageHeader"
            Begin
                Begin TextBox
                    BorderWidth =2
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1200
                    Top =30
                    Width =2718
                    Height =720
                    FontSize =12
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
                    Top =393
                    Width =681
                    Height =300
                    FontSize =7
                    TabIndex =1
                    BorderColor =32768
                    Name ="DATE_N"
                    ControlSource ="DATE_N"
                    FontName ="WeblogmaYekan"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            TextFontFamily =2
                            Left =718
                            Top =393
                            Width =450
                            Height =300
                            FontSize =8
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
                    Left =15
                    Top =33
                    Width =651
                    Height =300
                    FontSize =10
                    TabIndex =2
                    BorderColor =32768
                    Name ="NUMBER"
                    ControlSource ="NUMBER"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextFontFamily =2
                            Left =660
                            Top =30
                            Width =510
                            Height =315
                            FontSize =8
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
                    Visible = NotDefault
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =2157
                    Top =434
                    Width =291
                    Height =225
                    TabIndex =3
                    Name ="TAG"
                    ControlSource ="TAG"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    BorderWidth =2
                    Left =3909
                    Top =795
                    Width =0
                    Height =392
                    BorderColor =0
                    Name ="Line130"
                End
                Begin Line
                    BorderWidth =1
                    Left =1276
                    Top =795
                    Width =0
                    Height =392
                    BorderColor =0
                    Name ="Line137"
                End
                Begin Line
                    BorderWidth =2
                    Top =795
                    Width =0
                    Height =392
                    BorderColor =0
                    Name ="Line138"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =30
                    Top =795
                    Width =630
                    Height =345
                    FontSize =8
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
                    Left =1365
                    Top =825
                    Width =435
                    Height =315
                    FontSize =8
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label140"
                    Caption ="مصرف"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =1875
                    Top =840
                    Width =390
                    Height =300
                    FontSize =7
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label141"
                    Caption ="مقدار "
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =2295
                    Top =810
                    Width =1545
                    Height =330
                    FontSize =8
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label144"
                    Caption ="شـرح  كالا"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    BorderWidth =2
                    Top =1170
                    Width =3909
                    BorderColor =0
                    Name ="Line2"
                End
                Begin Rectangle
                    BorderWidth =2
                    Left =15
                    Width =3912
                    Height =778
                    Name ="Box239"
                End
                Begin Line
                    BorderWidth =1
                    Left =2280
                    Top =780
                    Width =0
                    Height =392
                    BorderColor =0
                    Name ="Line242"
                End
                Begin Line
                    BorderWidth =1
                    Left =1845
                    Top =795
                    Width =0
                    Height =392
                    BorderColor =0
                    Name ="Line243"
                End
                Begin Line
                    BorderWidth =1
                    Left =676
                    Top =795
                    Width =0
                    Height =392
                    BorderColor =0
                    Name ="Line254"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =765
                    Top =825
                    Width =465
                    Height =315
                    FontSize =8
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label256"
                    Caption ="فروش"
                    FontName ="WeblogmaYekan"
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =375
            OnFormat ="[Event Procedure]"
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
                    Left =3690
                    Top =15
                    Width =204
                    Height =330
                    FontSize =8
                    BorderColor =32768
                    Name ="RADIFff"
                    ControlSource ="=1"
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
                    Width =1374
                    Height =345
                    FontSize =6
                    TabIndex =1
                    BorderColor =32768
                    Name ="kala3"
                    ControlSource ="=IIf([Forms]![baseknow]![TFTPAGE]=1,[kala],[kala])"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1875
                    Top =15
                    Width =384
                    Height =330
                    FontSize =8
                    TabIndex =2
                    BorderColor =32768
                    Name ="Text219"
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
                    Left =1290
                    Top =15
                    Width =534
                    Height =330
                    FontSize =8
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
                    Left =15
                    Top =15
                    Width =639
                    Height =330
                    FontSize =8
                    TabIndex =4
                    BorderColor =32768
                    Name ="MABL_K"
                    ControlSource ="MABL_K"
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =2
                    Left =3909
                    Width =0
                    Height =375
                    BorderColor =0
                    Name ="Line222"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Left =1275
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
                    Left =2281
                    Width =0
                    Height =375
                    BorderColor =0
                    Name ="Line208"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Left =1845
                    Width =0
                    Height =375
                    BorderColor =0
                    Name ="Line246"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Left =675
                    Width =0
                    Height =375
                    BorderColor =0
                    Name ="Line255"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =705
                    Top =15
                    Width =549
                    Height =330
                    FontSize =8
                    TabIndex =5
                    BorderColor =32768
                    Name ="fikh1"
                    ControlSource ="=[MABL]-[fikh]"
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =2490
            Name ="GroupFooter3"
            Begin
                Begin Line
                    LineSlant = NotDefault
                    Left =1530
                    Width =0
                    Height =1895
                    BorderColor =0
                    Name ="Line126"
                End
                Begin Label
                    TextAlign =0
                    TextFontFamily =2
                    Left =1605
                    Width =1500
                    Height =435
                    FontSize =12
                    FontWeight =400
                    ForeColor =0
                    Name ="Label178"
                    Caption ="جمع مصرف كننده:"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =0
                    TextFontFamily =2
                    Left =1609
                    Top =1228
                    Width =1425
                    Height =585
                    FontSize =12
                    FontWeight =400
                    ForeColor =0
                    Name ="Label181"
                    Caption =" قابل پرداخت:"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    NumeralShapes =2
                    Width =1515
                    Height =585
                    FontSize =16
                    FontWeight =400
                    ForeColor =0
                    Name ="JF"
                    Caption ="جمع كل فاكتور:"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    NumeralShapes =2
                    Top =1230
                    Width =1515
                    Height =585
                    FontSize =16
                    FontWeight =400
                    ForeColor =0
                    Name ="GABEL"
                    Caption ="مبلغ قابل پرداخت:"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextFontFamily =2
                    ReadingOrder =2
                    NumeralShapes =2
                    Top =1890
                    Width =1815
                    Height =300
                    FontSize =7
                    FontWeight =400
                    BackColor =128
                    ForeColor =0
                    Name ="Label197"
                    Caption ="حسابداري دنا فراز تلفن 38259800"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =825
                    Width =3033
                    BorderColor =0
                    Name ="Line235"
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =15
                    Top =2160
                    Width =3828
                    Height =330
                    FontSize =10
                    BorderColor =32768
                    Name ="Text153"
                    ControlSource ="=[Forms]![Baseknow]![HIGH_D]"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =15
                    Top =615
                    Width =3885
                    BorderColor =0
                    Name ="Line236"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1875
                    Top =1860
                    Width =1656
                    Height =300
                    FontSize =8
                    TabIndex =1
                    BorderColor =32768
                    Name ="Text209"
                    ControlSource ="=UCURRENTUSER()"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextFontFamily =2
                            Left =3542
                            Top =1908
                            Width =330
                            Height =285
                            FontSize =7
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label210"
                            Caption ="کاربر:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin Rectangle
                    BorderWidth =2
                    Width =3929
                    Height =1858
                    Name ="Box247"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =15
                    Top =1230
                    Width =3885
                    BorderColor =0
                    Name ="Line248"
                End
                Begin Label
                    TextAlign =0
                    TextFontFamily =2
                    Left =1605
                    Top =645
                    Width =1440
                    Height =585
                    FontSize =12
                    FontWeight =400
                    ForeColor =0
                    Name ="Label249"
                    Caption ="سود شما::"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    NumeralShapes =2
                    Top =647
                    Width =1515
                    Height =585
                    FontSize =16
                    FontWeight =400
                    ForeColor =0
                    Name ="TF"
                    Caption ="سود شما:"
                    FontName ="WeblogmaYekan"
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
' See "I24.cls"
