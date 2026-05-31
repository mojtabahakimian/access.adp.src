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
    ItemSuffix =196
    Left =525
    Top =1830
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x40c829fb56efe540
    End
    RecordSource ="SELECT HEAD_LST.NUMBER, HEAD_LST.TAG, HEAD_LST.DATE_N, HEAD_LST.TAH, STUF_DEF.CO"
        "DE, STUF_DEF.NAME, STUF_DEF.N_FANI, TCOD_VAHEDS.NAMES AS VNAMES, INVO_LST.MEGH, "
        "INVO_LST.MEGHk, TCOD_ANBAR.NAMES, ISNULL(STUF_DEF.NAME, N' ') + N' ' + ISNULL(IN"
        "VO_LST.MANDAH, N' ') AS KALA, INVO_LST.RADIF, INVO_LST.ANBAR, SALA_DTL.EMZA AS e"
        "mza1, SALA_DTL_1.EMZA AS emza2, SALA_DTL_2.EMZA AS emza3, HEAD_LST.SGN1, HEAD_LS"
        "T.SGN2, HEAD_LST.SGN3, HEAD_LST.SGN4, HEAD_LST.sgn1usid, HEAD_LST.sgn2usid, HEAD"
        "_LST.sgn3usid FROM HEAD_LST INNER JOIN INVO_LST ON HEAD_LST.NUMBER = INVO_LST.NU"
        "MBER AND HEAD_LST.TAG = INVO_LST.TAG INNER JOIN STUF_DEF ON INVO_LST.CODE = STUF"
        "_DEF.CODE INNER JOIN TCOD_VAHEDS ON STUF_DEF.VAHED = TCOD_VAHEDS.CODE INNER JOIN"
        " TCOD_ANBAR ON INVO_LST.ANBAR = TCOD_ANBAR.CODE LEFT OUTER JOIN SALA_DTL SALA_DT"
        "L_2 ON HEAD_LST.sgn3usid = SALA_DTL_2.IDD LEFT OUTER JOIN SALA_DTL SALA_DTL_1 ON"
        " HEAD_LST.sgn2usid = SALA_DTL_1.IDD LEFT OUTER JOIN SALA_DTL ON HEAD_LST.sgn1usi"
        "d = SALA_DTL.IDD WHERE (HEAD_LST.TAG = 28)"
    Caption ="فاكتور خريد"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial (Arabic)"
    OnError ="[Event Procedure]"
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
            Height =1665
            Name ="PageHeaderSection"
            Begin
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =4170
                    Width =2385
                    Height =540
                    FontSize =14
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label88"
                    Caption ="درخواست کالا از انبار"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2655
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
                    Left =8560
                    Top =588
                    Width =1281
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
                            Left =9878
                            Top =588
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
                    Left =7530
                    Top =90
                    Width =1191
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
                            Left =8757
                            Top =90
                            Width =1785
                            Height =405
                            FontSize =12
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label93"
                            Caption ="شماره درخواست:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin Line
                    BorderWidth =3
                    Top =1110
                    Width =10660
                    BorderColor =0
                    Name ="Line94"
                End
                Begin Line
                    Left =3234
                    Top =1125
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line133"
                End
                Begin Line
                    Left =2409
                    Top =1125
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line134"
                End
                Begin Line
                    Left =1020
                    Top =1125
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line135"
                End
                Begin Line
                    BorderWidth =2
                    Top =1125
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line138"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =75
                    Top =1170
                    Width =900
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
                    Left =2581
                    Top =1170
                    Width =525
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
                    Left =1107
                    Top =1170
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
                    Left =3705
                    Top =1170
                    Width =3945
                    Height =405
                    FontSize =10
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label144"
                    Caption ="شــــــــــرح  كالا"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =3
                    Top =1635
                    Width =10671
                    BorderColor =0
                    Name ="Line147"
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
                    TabIndex =3
                    Name ="Text24"
                    ControlSource ="=\"صفحه \" & [Page] & \" از \" & [Pages]"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    Top =735
                    Width =1371
                    Height =300
                    FontSize =8
                    TabIndex =4
                    BorderColor =32768
                    Name ="Text170"
                    ControlSource ="=[Forms]![BASEKNOW]![ISO_KHAREED]"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =1410
                            Top =690
                            Width =750
                            Height =330
                            FontSize =8
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label171"
                            Caption ="شماره فرم :"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin Line
                    BorderWidth =3
                    Left =10635
                    Top =1080
                    Width =0
                    Height =585
                    BorderColor =0
                    Name ="Line130"
                End
                Begin Line
                    BorderWidth =1
                    Left =10077
                    Top =1125
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line131"
                End
                Begin Line
                    BorderWidth =1
                    Left =7905
                    Top =1155
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line132"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =9555
                    Top =1170
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
                    Left =10083
                    Top =1166
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
                    BorderWidth =1
                    Left =9480
                    Top =1125
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line189"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =7980
                    Top =1170
                    Width =1500
                    Height =390
                    FontSize =10
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label195"
                    Caption ="كد"
                    FontName ="WeblogmaYekan"
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
                    Left =3228
                    Top =16
                    Width =4644
                    Height =360
                    FontSize =10
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
                    Left =2423
                    Top =16
                    Width =789
                    Height =360
                    FontSize =10
                    TabIndex =1
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
                    Left =79
                    Top =16
                    Width =894
                    Height =360
                    FontSize =10
                    TabIndex =2
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
                    Left =1039
                    Top =15
                    Width =1320
                    Height =360
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
                    Top =405
                    Width =10660
                    BorderColor =0
                    Name ="Line100"
                End
                Begin Line
                    BorderWidth =1
                    Left =3234
                    Width =0
                    Height =405
                    BorderColor =0
                    Name ="Line104"
                End
                Begin Line
                    BorderWidth =1
                    Left =2409
                    Width =0
                    Height =405
                    BorderColor =0
                    Name ="Line105"
                End
                Begin Line
                    BorderWidth =1
                    Left =1009
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
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =9546
                    Top =15
                    Width =519
                    Height =405
                    FontSize =10
                    TabIndex =4
                    BorderColor =32768
                    Name ="IANBAR"
                    ControlSource ="ANBAR"
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
                    Left =10114
                    Top =16
                    Width =504
                    Height =405
                    FontSize =10
                    TabIndex =5
                    BorderColor =32768
                    Name ="RADIF123"
                    ControlSource ="=1"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    BorderWidth =3
                    Left =10636
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line101"
                End
                Begin Line
                    BorderWidth =1
                    Left =10098
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line102"
                End
                Begin Line
                    BorderWidth =1
                    Left =7905
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line103"
                End
                Begin Line
                    BorderWidth =1
                    Left =9501
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
                    Left =7971
                    Top =15
                    Width =1494
                    Height =360
                    FontSize =10
                    TabIndex =6
                    BorderColor =32768
                    Name ="N_FANI"
                    ControlSource ="N_FANI"
                    FontName ="WeblogmaYekan"
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =2377
            Name ="GroupFooter3"
            Begin
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =45
                    Top =219
                    Width =2276
                    FontSize =8
                    Name ="Text172"
                    ControlSource ="=Format(Farsidate(Now()),\"00\\/00\\/00\") & \" - \" & Format(Now(),\"Long Time\""
                        ")"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =9490
                    Top =163
                    Width =1035
                    Height =405
                    FontSize =10
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label184"
                    Caption ="اقدام كننده:"
                    FontName ="WeblogmaYekan"
                End
                Begin Rectangle
                    Left =15
                    Top =45
                    Width =10624
                    Height =614
                    Name ="Box186"
                End
                Begin BoundObjectFrame
                    SizeMode =3
                    OldBorderStyle =0
                    Left =6840
                    Top =676
                    Width =1701
                    Height =1701
                    TabIndex =1
                    Name ="EMZA1"
                    ControlSource ="EMZA1"
                End
                Begin BoundObjectFrame
                    SizeMode =3
                    OldBorderStyle =0
                    Left =3435
                    Top =661
                    Width =1701
                    Height =1701
                    TabIndex =2
                    Name ="EMZA2"
                    ControlSource ="EMZA2"
                End
                Begin BoundObjectFrame
                    SizeMode =3
                    OldBorderStyle =0
                    Top =657
                    Width =1701
                    Height =1701
                    TabIndex =3
                    Name ="EMZA3"
                    ControlSource ="EMZA3"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =8550
                    Top =1096
                    Width =1725
                    Height =285
                    FontSize =8
                    FontWeight =400
                    BackColor =128
                    ForeColor =0
                    Name ="semat1"
                    Caption ="."
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =5128
                    Top =1096
                    Width =1725
                    Height =285
                    FontSize =8
                    FontWeight =400
                    BackColor =128
                    ForeColor =0
                    Name ="semat2"
                    Caption ="."
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =1710
                    Top =1096
                    Width =1725
                    Height =285
                    FontSize =8
                    FontWeight =400
                    BackColor =128
                    ForeColor =0
                    Name ="semat3"
                    Caption ="."
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =8547
                    Top =763
                    Width =1725
                    Height =285
                    FontSize =8
                    FontWeight =400
                    BackColor =128
                    ForeColor =0
                    Name ="nemz1"
                    Caption ="."
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =5130
                    Top =763
                    Width =1725
                    Height =285
                    FontSize =8
                    FontWeight =400
                    BackColor =128
                    ForeColor =0
                    Name ="nemz2"
                    Caption ="."
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =1695
                    Top =763
                    Width =1725
                    Height =285
                    FontSize =8
                    FontWeight =400
                    BackColor =128
                    ForeColor =0
                    Name ="nemz3"
                    Caption ="."
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Left =5952
                    Top =1272
                    TabIndex =4
                    Name ="sgn1usid"
                    ControlSource ="sgn1usid"
                    Begin
                        Begin Label
                            TextAlign =0
                            Left =4251
                            Top =1272
                            Width =930
                            Height =255
                            Name ="Label99"
                            Caption ="sgn1usid:"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Left =5952
                    Top =1612
                    TabIndex =5
                    Name ="sgn2usid"
                    ControlSource ="sgn2usid"
                    Begin
                        Begin Label
                            TextAlign =0
                            Left =4251
                            Top =1612
                            Width =930
                            Height =255
                            Name ="Label100"
                            Caption ="sgn2usid:"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Left =5952
                    Top =1953
                    TabIndex =6
                    Name ="sgn3usid"
                    ControlSource ="sgn3usid"
                    Begin
                        Begin Label
                            TextAlign =0
                            Left =4251
                            Top =1953
                            Width =930
                            Height =255
                            Name ="Label101"
                            Caption ="sgn3usid:"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Left =1984
                    Top =1499
                    TabIndex =7
                    Name ="SGN1"
                    ControlSource ="SGN1"
                End
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Left =1984
                    Top =1818
                    TabIndex =8
                    Name ="SGN2"
                    ControlSource ="SGN2"
                End
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Left =1984
                    Top =2137
                    TabIndex =9
                    Name ="SGN3"
                    ControlSource ="SGN3"
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
' See "DARKHST_KALA_ANBAR.cls"
