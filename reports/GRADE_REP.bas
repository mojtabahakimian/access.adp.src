Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    DefaultView =0
    AllowUpdating =4
    TabularCharSet =178
    TabularFamily =42
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =10350
    DatasheetFontHeight =10
    ItemSuffix =249
    Left =600
    Top =780
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x6a40be21328ee540
    End
    ServerFilter ="HES='113-1-1             '"
    RecordSource ="GRADEVIEW"
    Caption ="رده بندي مشتري"
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
            ControlSource ="HES"
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            GroupFooter = NotDefault
            KeepTogether =1
            ControlSource ="GCNAME"
        End
        Begin BreakLevel
            ControlSource ="GCGRPNAME"
        End
        Begin PageHeader
            Height =1774
            BackColor =13434828
            Name ="PageHeader"
            Begin
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =3
                    Top =1249
                    Width =10335
                    BorderColor =0
                    Name ="Line99"
                End
                Begin Line
                    BorderWidth =3
                    Left =10335
                    Top =1264
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line130"
                End
                Begin Line
                    Left =9765
                    Top =1264
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line131"
                End
                Begin Line
                    Left =5895
                    Top =1264
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line133"
                End
                Begin Line
                    Left =5055
                    Top =1264
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line134"
                End
                Begin Line
                    Left =3930
                    Top =1264
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line135"
                End
                Begin Line
                    BorderWidth =3
                    Top =1264
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line138"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =5265
                    Top =1309
                    Width =435
                    Height =285
                    FontSize =8
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label142"
                    Caption ="ضريب"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =7065
                    Top =1309
                    Width =915
                    Height =285
                    FontSize =8
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label144"
                    Caption ="عنوان"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =9854
                    Top =1305
                    Width =375
                    Height =285
                    FontSize =8
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
                    Top =1774
                    Width =10350
                    BorderColor =0
                    Name ="Line147"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =2385
                    Top =397
                    Height =285
                    FontSize =8
                    Name ="HES"
                    ControlSource ="HES"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextAlign =0
                            TextFontFamily =2
                            Left =4260
                            Top =394
                            Width =645
                            Height =285
                            FontSize =8
                            FontWeight =400
                            Name ="Label196"
                            Caption ="کد حساب:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =7596
                    Top =844
                    Height =285
                    FontSize =8
                    TabIndex =1
                    Name ="TOPETEB"
                    ControlSource ="TOPETEB"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextAlign =0
                            TextFontFamily =2
                            Left =9345
                            Top =844
                            Width =855
                            Height =285
                            FontSize =8
                            FontWeight =400
                            Name ="Label198"
                            Caption ="سقف اعتبار:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =4592
                    Top =844
                    Height =285
                    FontSize =8
                    TabIndex =2
                    Name ="TOPMEGH"
                    ControlSource ="TOPMEGH"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextAlign =0
                            TextFontFamily =2
                            Left =6345
                            Top =844
                            Width =1230
                            Height =285
                            FontSize =8
                            FontWeight =400
                            Name ="Label199"
                            Caption ="سقف اعتبار مقداري:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =4913
                    Top =390
                    Width =4611
                    Height =285
                    FontSize =8
                    TabIndex =3
                    Name ="NAME"
                    ControlSource ="NAME"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextAlign =0
                            TextFontFamily =2
                            Left =9585
                            Top =394
                            Width =735
                            Height =285
                            FontSize =8
                            FontWeight =400
                            Name ="Label200"
                            Caption ="نام مشتري:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =113
                    Top =390
                    Height =285
                    FontSize =8
                    TabIndex =4
                    Name ="MOBILE"
                    ControlSource ="MOBILE"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextAlign =0
                            TextFontFamily =2
                            Left =1920
                            Top =394
                            Width =495
                            Height =285
                            FontSize =8
                            FontWeight =400
                            Name ="Label201"
                            Caption ="موبايل:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin CheckBox
                    SpecialEffect =2
                    Left =1303
                    Top =876
                    TabIndex =5
                    Name ="Check202"
                    ControlSource ="TOPASN"
                    Begin
                        Begin Label
                            TextFontCharSet =0
                            TextAlign =0
                            TextFontFamily =34
                            Left =1747
                            Top =844
                            Width =2775
                            Height =240
                            FontSize =8
                            FontWeight =400
                            ForeColor =0
                            Name ="Label6"
                            Caption ="اسناد پاس نشده از اعتبار كسر شود"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =4251
                    Top =1302
                    Width =555
                    Height =285
                    FontSize =8
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label143"
                    Caption ="امتياز"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =2895
                    Top =1309
                    Width =915
                    Height =285
                    FontSize =8
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label224"
                    Caption ="امتياز با ضريب"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    Left =2775
                    Top =1264
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line226"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =1815
                    Top =1309
                    Width =915
                    Height =285
                    FontSize =8
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label233"
                    Caption ="امتياز گروه"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    Left =1530
                    Top =1264
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line235"
                End
                Begin Label
                    TextAlign =0
                    TextFontFamily =2
                    Left =4365
                    Width =1545
                    Height =405
                    FontSize =12
                    FontWeight =400
                    ForeColor =0
                    Name ="Label248"
                    Caption ="رده بندي مشتريان"
                    FontName ="WeblogmaYekan"
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =450
            BreakLevel =1
            BackColor =10092543
            Name ="GroupHeader0"
            Begin
                Begin TextBox
                    DecimalPlaces =0
                    RunningSum =1
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =9766
                    Width =504
                    Height =345
                    FontSize =10
                    FontWeight =700
                    BorderColor =32768
                    Name ="Text213"
                    ControlSource ="=1"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5937
                    Width =3759
                    Height =345
                    FontSize =10
                    FontWeight =700
                    TabIndex =1
                    BorderColor =32768
                    Name ="Text214"
                    ControlSource ="GCNAME"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5070
                    Width =789
                    Height =345
                    FontSize =10
                    FontWeight =700
                    TabIndex =2
                    BorderColor =32768
                    Name ="GCZARIB"
                    ControlSource ="GCZARIB"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    BorderWidth =3
                    Left =10335
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line217"
                End
                Begin Line
                    BorderWidth =1
                    Left =9764
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line218"
                End
                Begin Line
                    BorderWidth =1
                    Left =5895
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line219"
                End
                Begin Line
                    BorderWidth =1
                    Left =5055
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line220"
                End
                Begin Line
                    BorderWidth =1
                    Left =3930
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line221"
                End
                Begin Line
                    BorderWidth =2
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line222"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2777
                    Width =1110
                    Height =345
                    FontSize =10
                    FontWeight =700
                    TabIndex =3
                    BorderColor =32768
                    Name ="EMKOLZ"
                    ControlSource ="=[EMKOL]*[GCZARIB]"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    BorderWidth =1
                    Left =2775
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line227"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3975
                    Width =1110
                    Height =345
                    FontSize =10
                    FontWeight =700
                    TabIndex =4
                    BorderColor =32768
                    Name ="EMKOL"
                    ControlSource ="=[EMTZARSS]/[GCGRPZARIBS]"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =1
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1590
                    Width =1110
                    Height =345
                    FontSize =10
                    FontWeight =700
                    TabIndex =5
                    BorderColor =32768
                    Name ="EMKOLZS"
                    ControlSource ="=[EMKOLZ]/[JAMZARIB]"
                    Format ="#"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    BorderWidth =1
                    Left =1530
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line236"
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =450
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
                    Left =9781
                    Top =16
                    Width =504
                    Height =285
                    FontSize =8
                    BorderColor =32768
                    Name ="RADIFwe"
                    ControlSource ="=1"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5952
                    Top =16
                    Width =3129
                    Height =285
                    FontSize =8
                    TabIndex =1
                    BorderColor =32768
                    Name ="GCGRPNAME"
                    ControlSource ="GCGRPNAME"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5085
                    Top =16
                    Width =789
                    Height =285
                    FontSize =8
                    TabIndex =2
                    BorderColor =32768
                    Name ="GCGRPZARIB"
                    ControlSource ="GCGRPZARIB"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3920
                    Top =15
                    Width =1110
                    Height =285
                    FontSize =8
                    TabIndex =3
                    BorderColor =32768
                    Name ="GCGRPGRADE"
                    ControlSource ="GCGRPGRADE"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    BorderWidth =1
                    Top =450
                    Width =10335
                    BorderColor =0
                    Name ="Line100"
                End
                Begin Line
                    BorderWidth =3
                    Left =10335
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line101"
                End
                Begin Line
                    BorderWidth =1
                    Left =9765
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line102"
                End
                Begin Line
                    BorderWidth =1
                    Left =5895
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line104"
                End
                Begin Line
                    BorderWidth =1
                    Left =5055
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line105"
                End
                Begin Line
                    BorderWidth =1
                    Left =3930
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
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2775
                    Top =30
                    Width =1110
                    Height =285
                    FontSize =8
                    TabIndex =4
                    BorderColor =32768
                    Name ="EMTZAR"
                    ControlSource ="EMTZAR"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    BorderWidth =1
                    Left =2775
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line228"
                End
                Begin Line
                    BorderWidth =1
                    Left =1530
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line237"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1650
                    Width =1110
                    Height =285
                    FontSize =8
                    TabIndex =5
                    BorderColor =32768
                    Name ="Text238"
                    ControlSource ="=[EMTZAR]/[GCGRPZARIBS]"
                    Format ="#"
                    FontName ="WeblogmaYekan"
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =60
            BreakLevel =1
            Name ="GroupFooter3"
            Begin
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =3
                    Top =30
                    Width =10350
                    BorderColor =0
                    Name ="Line110"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Width =1110
                    Height =60
                    FontSize =8
                    BorderColor =32768
                    Name ="EMTZARSS"
                    ControlSource ="=Sum([EMTZAR])"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =8107
                    Width =1110
                    Height =0
                    FontSize =8
                    TabIndex =1
                    BorderColor =32768
                    Name ="GCGRPZARIBS"
                    ControlSource ="=Sum([GCGRPZARIB])"
                    FontName ="WeblogmaYekan"
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            ForceNewPage =2
            Height =405
            BackColor =16777164
            Name ="GroupFooter0"
            Begin
                Begin TextBox
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =7313
                    Height =405
                    FontSize =12
                    Name ="JAMZARIB"
                    ControlSource ="JAMZARIB"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =1
                    TextAlign =2
                    TextFontFamily =2
                    IMESentenceMode =3
                    Height =405
                    FontSize =12
                    TabIndex =1
                    Name ="EMTIAZ"
                    ControlSource ="EMTIAZ"
                    Format ="#"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =1687
                    Width =825
                    Height =405
                    FontSize =12
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label246"
                    Caption ="امتياز کل:"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =9067
                    Width =1020
                    Height =405
                    FontSize =12
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label247"
                    Caption ="جمع ضرايب:"
                    FontName ="WeblogmaYekan"
                End
            End
        End
        Begin PageFooter
            Visible = NotDefault
            Height =170
            Name ="PageFooter"
        End
    End
End
CodeBehindForm
' See "GRADE_REP.cls"
