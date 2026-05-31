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
    Width =6406
    DatasheetFontHeight =10
    ItemSuffix =236
    Left =600
    Top =1935
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x1cce23814f17e540
    End
    RecordSource ="SELECT PRGHEAD.PRGID, PRGHEAD.PRG_DATE, PRGHEAD.PCODE, STUF_DEF.NAME, STUF_DEF.C"
        "ODE, QPROGPAS2.CODB, QPROGPAS2.VAHED, QPROGPAS2.SumOfMEGH, QPROGPAS2.SumOfMEGHK,"
        " QPROGPAS2.SumOfPERT, QPROGPAS2.SumOfKOLMAV, QPROGPAS2.SumOfMABL, QPROGPAS2.SumO"
        "fMABLK, STUF_DEF_1.NAME AS MAVAD, STUF_DEF_1.CODE AS MAVCO FROM PRGHEAD INNER JO"
        "IN STUF_DEF ON PRGHEAD.PCODE = STUF_DEF.CODE INNER JOIN QPROGPAS2 ON PRGHEAD.PRG"
        "ID = QPROGPAS2.PRGID INNER JOIN STUF_DEF STUF_DEF_1 ON QPROGPAS2.CODB = STUF_DEF"
        "_1.CODE"
    Caption ="چاپ فرمول"
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
            ControlSource ="PRGID"
        End
        Begin BreakLevel
            SortOrder = NotDefault
            ControlSource ="SumOfMABLK"
        End
        Begin PageHeader
            Height =1470
            Name ="PageHeader"
            Begin
                Begin Label
                    TextFontFamily =2
                    Left =2625
                    Width =1005
                    Height =390
                    FontSize =10
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label88"
                    Caption ="فرمول ساخت"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1425
                    Top =300
                    Width =3288
                    Height =330
                    FontSize =8
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
                    Top =312
                    Width =846
                    Height =300
                    FontSize =8
                    TabIndex =1
                    BorderColor =32768
                    Name ="DATE_ACTIV"
                    ControlSource ="PRG_DATE"
                    FontName ="WeblogmaYekan"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =883
                            Top =312
                            Width =465
                            Height =330
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
                    Left =4700
                    Top =300
                    Width =651
                    Height =315
                    FontSize =8
                    TabIndex =2
                    BorderColor =32768
                    Name ="FNUMB"
                    ControlSource ="PRGID"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =5385
                            Top =300
                            Width =960
                            Height =330
                            FontSize =8
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label93"
                            Caption ="شماره برنامه:"
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
                    Left =1455
                    Top =689
                    Width =4266
                    Height =330
                    FontSize =8
                    TabIndex =3
                    BorderColor =32768
                    Name ="KALA"
                    ControlSource ="NAME"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =5730
                            Top =690
                            Width =630
                            Height =330
                            FontSize =8
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label95"
                            Caption ="نام كالا:"
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
                    Top =690
                    Width =801
                    Height =330
                    FontSize =8
                    TabIndex =4
                    BorderColor =32768
                    Name ="COKAL"
                    ControlSource ="CODE"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =836
                            Top =691
                            Width =570
                            Height =330
                            FontSize =8
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label96"
                            Caption ="كد  كالا:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Top =1065
                    Width =6366
                    BorderColor =0
                    Name ="Line99"
                End
                Begin Line
                    BorderWidth =1
                    Left =6341
                    Top =1080
                    Width =0
                    Height =390
                    BorderColor =0
                    Name ="Line130"
                End
                Begin Line
                    Left =5891
                    Top =1080
                    Width =0
                    Height =390
                    BorderColor =0
                    Name ="Line131"
                End
                Begin Line
                    Left =2495
                    Top =1080
                    Width =0
                    Height =390
                    BorderColor =0
                    Name ="Line133"
                End
                Begin Line
                    Left =1950
                    Top =1080
                    Width =0
                    Height =390
                    BorderColor =0
                    Name ="Line136"
                End
                Begin Line
                    Left =1125
                    Top =1080
                    Width =0
                    Height =390
                    BorderColor =0
                    Name ="Line137"
                End
                Begin Line
                    BorderWidth =1
                    Top =1080
                    Width =0
                    Height =390
                    BorderColor =0
                    Name ="Line138"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =300
                    Top =1110
                    Width =525
                    Height =330
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
                    Left =1238
                    Top =1125
                    Width =630
                    Height =330
                    FontSize =8
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
                    Left =2015
                    Top =1125
                    Width =390
                    Height =330
                    FontSize =8
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label141"
                    Caption ="پرت"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =3732
                    Top =1110
                    Width =1995
                    Height =330
                    FontSize =8
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
                    Left =5925
                    Top =1125
                    Width =390
                    Height =330
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
                    BorderWidth =1
                    Top =1455
                    Width =6366
                    BorderColor =0
                    Name ="Line147"
                End
                Begin Line
                    Left =3315
                    Top =1080
                    Width =0
                    Height =390
                    BorderColor =0
                    Name ="Line220"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =2520
                    Top =1125
                    Width =780
                    Height =330
                    FontSize =8
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label221"
                    Caption ="مقدار مصرف"
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
                    Left =5925
                    Top =15
                    Width =414
                    Height =330
                    FontSize =8
                    BorderColor =32768
                    Name ="RADIF"
                    ControlSource ="=1"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3345
                    Top =16
                    Width =2529
                    Height =330
                    FontSize =8
                    TabIndex =1
                    BorderColor =32768
                    Name ="MAVAD"
                    ControlSource ="MAVAD"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1965
                    Top =15
                    Width =504
                    Height =330
                    FontSize =10
                    TabIndex =2
                    BorderColor =32768
                    Name ="MEGHk"
                    ControlSource ="SumOfPERT"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1170
                    Top =15
                    Width =774
                    Height =330
                    FontSize =8
                    TabIndex =3
                    BorderColor =32768
                    Name ="MABL"
                    ControlSource ="SumOfMABL"
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
                    Left =60
                    Top =15
                    Width =1044
                    Height =330
                    FontSize =8
                    TabIndex =4
                    BorderColor =32768
                    Name ="SumOfMABLK"
                    ControlSource ="SumOfMABLK"
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    BorderWidth =1
                    Top =375
                    Width =6366
                    BorderColor =0
                    Name ="Line100"
                End
                Begin Line
                    BorderWidth =1
                    Left =6341
                    Width =0
                    Height =375
                    BorderColor =0
                    Name ="Line101"
                End
                Begin Line
                    BorderWidth =1
                    Left =5891
                    Width =0
                    Height =375
                    BorderColor =0
                    Name ="Line102"
                End
                Begin Line
                    BorderWidth =1
                    Left =2495
                    Width =0
                    Height =375
                    BorderColor =0
                    Name ="Line104"
                End
                Begin Line
                    BorderWidth =1
                    Left =1950
                    Width =0
                    Height =375
                    BorderColor =0
                    Name ="Line107"
                End
                Begin Line
                    BorderWidth =1
                    Left =1125
                    Width =0
                    Height =375
                    BorderColor =0
                    Name ="Line108"
                End
                Begin Line
                    BorderWidth =1
                    Width =0
                    Height =375
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
                    Left =2505
                    Top =15
                    Width =789
                    Height =330
                    FontSize =10
                    TabIndex =5
                    BorderColor =32768
                    Name ="Text222"
                    ControlSource ="SumOfMEGHK"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    BorderWidth =1
                    Left =3315
                    Width =0
                    Height =375
                    BorderColor =0
                    Name ="Line223"
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =1606
            Name ="GroupFooter3"
            Begin
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Width =6325
                    BorderColor =0
                    Name ="Line110"
                End
                Begin Line
                    BorderWidth =1
                    Top =1170
                    Width =6355
                    BorderColor =0
                    Name ="Line120"
                End
                Begin Line
                    BorderWidth =1
                    Left =6345
                    Width =0
                    Height =1606
                    BorderColor =0
                    Name ="Line127"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =1
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =15
                    Top =390
                    Width =1086
                    Height =360
                    Name ="IMBIBE_MANF"
                    ControlSource ="=[Forms]![HEAD_MANF]![DASTR]"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextFontFamily =2
                            Left =1140
                            Top =390
                            Width =1380
                            Height =360
                            FontWeight =400
                            ForeColor =0
                            Name ="Label224"
                            Caption ="جذب هزينه دستمزد:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =1
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =17
                    Top =775
                    Width =1071
                    Height =360
                    TabIndex =1
                    Name ="IMBIBE_SAR"
                    ControlSource ="=[Forms]![HEAD_MANF]![SARR]"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextFontFamily =2
                            Left =1124
                            Top =775
                            Width =1290
                            Height =360
                            FontWeight =400
                            ForeColor =0
                            Name ="Label225"
                            Caption ="جذب هزينه سربار:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =1
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =28
                    Width =1071
                    Height =360
                    TabIndex =2
                    Name ="smb"
                    ControlSource ="=Sum([SumOfMABLK])"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextFontFamily =2
                            Left =1135
                            Width =1290
                            Height =360
                            FontWeight =400
                            ForeColor =0
                            Name ="Label227"
                            Caption ="جمع مواد مصرفي:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =1
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =15
                    Top =1200
                    Width =1071
                    Height =360
                    TabIndex =3
                    Name ="Text228"
                    ControlSource ="=[Forms]![HEAD_MANF]![MAVADR]+[Forms]![HEAD_MANF]![DASTR]+[Forms]![HEAD_MANF]![S"
                        "ARR]"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextFontFamily =2
                            Left =1122
                            Top =1200
                            Width =1290
                            Height =360
                            FontWeight =400
                            ForeColor =0
                            Name ="Label229"
                            Caption ="قيمت تمام شده:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin Line
                    BorderWidth =1
                    Top =1590
                    Width =6355
                    BorderColor =0
                    Name ="Line230"
                End
                Begin Line
                    BorderWidth =1
                    Width =0
                    Height =1606
                    BorderColor =0
                    Name ="Line231"
                End
                Begin Line
                    BorderWidth =1
                    Top =765
                    Width =2605
                    BorderColor =0
                    Name ="Line232"
                End
                Begin Line
                    BorderWidth =1
                    Top =375
                    Width =2605
                    BorderColor =0
                    Name ="Line233"
                End
                Begin Line
                    BorderWidth =1
                    Left =2595
                    Width =0
                    Height =1606
                    BorderColor =0
                    Name ="Line234"
                End
                Begin Line
                    BorderWidth =1
                    Left =1125
                    Width =0
                    Height =1606
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
' See "PRINT_FORMULAR.cls"
