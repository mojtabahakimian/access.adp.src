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
    Width =7413
    DatasheetFontHeight =10
    ItemSuffix =328
    Left =600
    Top =2100
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xc2d1de97392be440
    End
    RecordSource ="SELECT TOP 100 PERCENT HEAD_LST.NUMBER, HEAD_LST.DATE_N, HEAD_LST.TAG, INVO_LST."
        "RADIF, INVO_LST.CODE, INVO_LST.MEGH, INVO_LST.MEGHk, CUST_HESAB.NAME AS HESAB, C"
        "UST_HESAB.ADDRESS, CUST_HESAB.TEL, HEAD_LST.CUST_NO, CUST_HESAB.ECODE, CUST_HESA"
        "B.PCODE, CUST_HESAB.IYALAT, CUST_HESAB.MCODEM, CUST_HESAB.CITY, CUST_HESAB.CODE_"
        "E, OTHER_DTL.REQUEST_NO, OTHER_DTL.BARNAMEH, OTHER_DTL.DRIVER, OTHER_DTL.DRIVER_"
        "MOB, OTHER_DTL.CAMIUN_NUM, OTHER_DTL.MAGHSAD, OTHER_DTL.CAM_KHALY, OTHER_DTL.CAM"
        "_POOR, OTHER_DTL.TOZIH, OTHER_DTL.CAMIUN, OTHER_DTL_SUB.CODE AS CODEG, OTHER_DTL"
        "_SUB.CAM_KHALY AS CAM_KHALYG, OTHER_DTL_SUB.CAM_POOR AS CAM_POORG, OTHER_DTL_SUB"
        ".MEGHk AS MEGHkG, OTHER_DTL_SUB.TOZIH AS TOZIHG, OTHER_DTL_SUB.VAZNH, STUF_DEF_1"
        ".NAME AS NAMEG FROM OTHER_DTL_SUB RIGHT OUTER JOIN CUST_HESAB RIGHT OUTER JOIN H"
        "EAD_LST ON CUST_HESAB.hes = HEAD_LST.CUST_NO RIGHT OUTER JOIN STUF_DEF STUF_DEF_"
        "1 INNER JOIN INVO_LST ON STUF_DEF_1.CODE = INVO_LST.CODE ON HEAD_LST.TAG = INVO_"
        "LST.TAG AND HEAD_LST.NUMBER = INVO_LST.NUMBER ON OTHER_DTL_SUB.CODE = INVO_LST.C"
        "ODE AND OTHER_DTL_SUB.NUMBER = INVO_LST.NUMBER AND OTHER_DTL_SUB.TAGG = INVO_LST"
        ".TAG LEFT OUTER JOIN OTHER_DTL ON HEAD_LST.NUMBER = OTHER_DTL.NUMBER AND HEAD_LS"
        "T.TAG = OTHER_DTL.TAG WHERE (HEAD_LST.TAG = 20)"
    Caption ="فاكتور فروش"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
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
            GroupFooter = NotDefault
            KeepTogether =1
            ControlSource ="NUMBER"
        End
        Begin BreakLevel
            ControlSource ="RADIF"
        End
        Begin PageHeader
            Height =1866
            OnFormat ="[Event Procedure]"
            Name ="PageHeader"
            Begin
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1260
                    Width =6108
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
                    Top =285
                    Width =771
                    Height =285
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
                            Left =798
                            Top =285
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
                    Left =30
                    Width =681
                    Height =270
                    FontSize =8
                    TabIndex =2
                    BorderColor =32768
                    Name ="NUMBER"
                    ControlSource ="NUMBER"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextFontFamily =2
                            Left =736
                            Width =495
                            Height =300
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
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1615
                    Top =642
                    Width =5286
                    Height =300
                    FontSize =8
                    TabIndex =3
                    BorderColor =32768
                    Name ="CUST_NO2"
                    ControlSource ="=[CUST_NO] & \"   /    \" & [HESAB]"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextFontFamily =2
                            Left =6914
                            Top =642
                            Width =465
                            Height =300
                            FontSize =8
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label95"
                            Caption ="مشتري:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin Line
                    BorderWidth =1
                    Left =10
                    Top =585
                    Width =7398
                    BorderColor =0
                    Name ="Line164"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Top =882
                    Width =1011
                    Height =300
                    FontSize =8
                    TabIndex =4
                    BorderColor =32768
                    Name ="TEL"
                    ControlSource ="TEL"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextFontFamily =2
                            Left =1050
                            Top =882
                            Width =405
                            Height =330
                            FontSize =8
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
                    Left =1275
                    Top =15
                    Width =0
                    Height =585
                    BorderColor =0
                    Name ="Line219"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =3690
                    Top =240
                    Width =1320
                    Height =405
                    FontSize =10
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label221"
                    Caption ="حواله انبار فروش"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Top =600
                    Width =1026
                    Height =285
                    FontSize =8
                    TabIndex =5
                    BorderColor =32768
                    Name ="MCODEM"
                    ControlSource ="MCODEM"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextFontFamily =2
                            Left =1049
                            Top =600
                            Width =570
                            Height =330
                            FontSize =8
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label227"
                            Caption ="ش.ملي:"
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
                    Left =1540
                    Top =898
                    Width =5376
                    Height =285
                    FontSize =8
                    TabIndex =6
                    BorderColor =32768
                    Name ="ADDRESS"
                    ControlSource ="ADDRESS"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextFontFamily =2
                            Left =6922
                            Top =898
                            Width =450
                            Height =300
                            FontSize =8
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
                    BackStyle =1
                    Left =10
                    Top =1545
                    Width =7394
                    Height =302
                    BackColor =10092543
                    Name ="Box224"
                End
                Begin Line
                    Left =6940
                    Top =1545
                    Width =0
                    Height =300
                    BorderColor =0
                    Name ="Line131"
                End
                Begin Line
                    Left =4705
                    Top =1545
                    Width =0
                    Height =300
                    BorderColor =0
                    Name ="Line133"
                End
                Begin Line
                    Left =4000
                    Top =1545
                    Width =0
                    Height =300
                    BorderColor =0
                    Name ="Line136"
                End
                Begin Line
                    Left =3030
                    Top =1545
                    Width =0
                    Height =300
                    BorderColor =0
                    Name ="Line137"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =2370
                    Top =1560
                    Width =405
                    Height =270
                    FontSize =7
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label139"
                    Caption ="وزن پر"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =3240
                    Top =1560
                    Width =585
                    Height =270
                    FontSize =7
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label140"
                    Caption ="وزن خالي"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =4075
                    Top =1545
                    Width =510
                    Height =285
                    FontSize =7
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label141"
                    Caption ="وزن كالا"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =5385
                    Top =1545
                    Width =1035
                    Height =285
                    FontSize =7
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
                    Left =6982
                    Top =1545
                    Width =375
                    Height =285
                    FontSize =7
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label146"
                    Caption ="رديف"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5455
                    Top =1230
                    Width =1071
                    Height =330
                    FontSize =8
                    TabIndex =7
                    BorderColor =32768
                    Name ="CAMIUN_NUM"
                    ControlSource ="CAMIUN_NUM"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextFontFamily =2
                            Left =6504
                            Top =1234
                            Width =855
                            Height =330
                            FontSize =8
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label232"
                            Caption ="شماره ماشين:"
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
                    Left =2910
                    Top =1200
                    Width =1551
                    Height =330
                    FontSize =8
                    TabIndex =8
                    BorderColor =32768
                    Name ="DRIVER"
                    ControlSource ="DRIVER"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextFontFamily =2
                            Left =4484
                            Top =1204
                            Width =690
                            Height =330
                            FontSize =8
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label234"
                            Caption ="نام راننده:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin Rectangle
                    Width =7398
                    Height =1866
                    Name ="Box237"
                End
                Begin Line
                    BorderWidth =1
                    Left =10
                    Top =1212
                    Width =7395
                    BorderColor =0
                    Name ="Line317"
                End
                Begin Line
                    Left =2070
                    Top =1545
                    Width =0
                    Height =300
                    BorderColor =0
                    Name ="Line319"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =1335
                    Top =1575
                    Width =585
                    Height =270
                    FontSize =7
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label320"
                    Caption ="وزن خالص"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    Left =1155
                    Top =1545
                    Width =0
                    Height =300
                    BorderColor =0
                    Name ="Line323"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =30
                    Top =1560
                    Width =1080
                    Height =270
                    FontSize =7
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label324"
                    Caption ="تحويل دهنده"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =5026
                    Top =274
                    TabIndex =9
                    Name ="CUST_NO"
                    ControlSource ="CUST_NO"
                    FontName ="WeblogmaYekan"
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =567
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
                    Left =6990
                    Top =150
                    Width =369
                    Height =330
                    FontSize =8
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
                    Left =4695
                    Top =135
                    Width =2229
                    Height =330
                    FontSize =8
                    TabIndex =1
                    BorderColor =32768
                    Name ="NAMEG"
                    ControlSource ="NAMEG"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4030
                    Top =105
                    Width =639
                    Height =330
                    FontSize =8
                    TabIndex =2
                    BorderColor =32768
                    Name ="MEGHk"
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
                    Left =3095
                    Top =105
                    Width =849
                    Height =330
                    FontSize =8
                    TabIndex =3
                    BorderColor =32768
                    Name ="CAM_KHALYG"
                    ControlSource ="CAM_KHALYG"
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
                    Left =2100
                    Top =105
                    Width =879
                    Height =330
                    FontSize =8
                    TabIndex =4
                    BorderColor =32768
                    Name ="CAM_POORG"
                    ControlSource ="CAM_POORG"
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    BorderWidth =1
                    Top =567
                    Width =7413
                    BorderColor =0
                    Name ="Line100"
                End
                Begin Line
                    BorderWidth =1
                    Left =6940
                    Width =0
                    Height =567
                    BorderColor =0
                    Name ="Line102"
                End
                Begin Line
                    BorderWidth =1
                    Left =4705
                    Width =0
                    Height =567
                    BorderColor =0
                    Name ="Line104"
                End
                Begin Line
                    BorderWidth =1
                    Left =4000
                    Width =0
                    Height =567
                    BorderColor =0
                    Name ="Line107"
                End
                Begin Line
                    BorderWidth =1
                    Left =3030
                    Width =0
                    Height =567
                    BorderColor =0
                    Name ="Line108"
                End
                Begin Line
                    BorderWidth =1
                    Width =0
                    Height =567
                    BorderColor =0
                    Name ="Line109"
                End
                Begin TextBox
                    Visible = NotDefault
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =4808
                    Top =47
                    Height =195
                    FontSize =7
                    TabIndex =5
                    Name ="CODE"
                    ControlSource ="CODE"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Left =7405
                    Width =0
                    Height =567
                    BorderColor =0
                    Name ="Line318"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1185
                    Top =105
                    Width =849
                    Height =330
                    FontSize =8
                    TabIndex =6
                    BorderColor =32768
                    Name ="MEGHkG"
                    ControlSource ="MEGHkG"
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    BorderWidth =1
                    Left =2070
                    Width =0
                    Height =567
                    BorderColor =0
                    Name ="Line322"
                End
                Begin Line
                    BorderWidth =1
                    Left =1155
                    Width =0
                    Height =567
                    BorderColor =0
                    Name ="Line326"
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =1461
            Name ="GroupFooter3"
            Begin
                Begin Line
                    Top =1095
                    Width =7410
                    BorderColor =0
                    Name ="Line125"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =6307
                    Width =1029
                    Height =270
                    FontSize =8
                    BorderColor =32768
                    Name ="Text211"
                    ControlSource ="=Format(Now(),\"hh:nn:ss ampm\")"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =1
                    BorderWidth =1
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4000
                    Width =714
                    Height =391
                    FontSize =10
                    TabIndex =1
                    Name ="Text200"
                    ControlSource ="=Sum([MEGHk])"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextFontFamily =2
                    Left =4930
                    Width =735
                    Height =390
                    FontSize =10
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label226"
                    Caption ="جمع كل:"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =122
                    Top =736
                    Width =1446
                    Height =300
                    FontSize =7
                    TabIndex =2
                    BorderColor =32768
                    Name ="Text209"
                    ControlSource ="=UCURRENTUSER()"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =471
                            Top =510
                            Width =675
                            Height =270
                            FontSize =7
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label210"
                            Caption ="صادر كننده:"
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
                    Left =1915
                    Top =1110
                    Width =4911
                    Height =330
                    FontSize =10
                    TabIndex =3
                    BorderColor =32768
                    Name ="Text199"
                    ControlSource ="=[Forms]![Baseknow]![TFADDRESS]"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextFontFamily =2
                            Left =6850
                            Top =1110
                            Width =510
                            Height =330
                            FontSize =10
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
                    Left =45
                    Top =1110
                    Width =1431
                    Height =330
                    FontSize =10
                    TabIndex =4
                    BorderColor =32768
                    Name ="Text197"
                    ControlSource ="=[Forms]![Baseknow]![TFTEL]"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =1500
                            Top =1110
                            Width =375
                            Height =330
                            FontSize =10
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label198"
                            Caption ="تلفن:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin Rectangle
                    Left =-5
                    Width =7413
                    Height =1461
                    Name ="Box238"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =2
                    Left =1633
                    Top =603
                    Width =5700
                    Height =465
                    FontSize =12
                    FontWeight =400
                    BackColor =128
                    ForeColor =0
                    Name ="MANDAH"
                    Caption ="مانده حساب:"
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
' See "P2.cls"
