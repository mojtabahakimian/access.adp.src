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
    GridX =64
    GridY =64
    Width =10559
    DatasheetFontHeight =10
    ItemSuffix =212
    Left =600
    Top =930
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xc3a366fd9ae4e340
    End
    RecordSource ="SELECT HEAD_LST.NUMBER, HEAD_LST.TAG, HEAD_LST.DATE_N, HEAD_LST.CUST_NO, HEAD_LS"
        "T.ANBAR, HEAD_LST.NUMBER1, HEAD_LST.TAH, HEAD_LST.MAS, HEAD_LST.VAS, HEAD_LST.N_"
        "S, HEAD_LST.M_NAGHD, HEAD_LST.MABL_VAR, HEAD_LST.MOIN_VAR, HEAD_LST.MABL_HAV, HE"
        "AD_LST.MOIN_HAV, HEAD_LST.MABL_HAZ, HEAD_LST.MOIN_HAZ, HEAD_LST.TAKHFIF, HEAD_LS"
        "T.MOIN_KHF, HEAD_LST.ANBARF, HEAD_LST.FNUMCO, INVO_LST.NUMBER AS INUMBER, INVO_L"
        "ST.TAG AS ITAG, INVO_LST.ANBAR AS IANBAR, INVO_LST.RADIF, INVO_LST.CODE, INVO_LS"
        "T.MEGH, INVO_LST.MEGHk, INVO_LST.MEGH_MAR, INVO_LST.MANDAH, INVO_LST.MABL, INVO_"
        "LST.MABL_K, INVO_LST.FROM_A, INVO_LST.N_RASID, INVO_LST.MEGH_R, INVO_LST.CUST_NO"
        " AS ICUST_NO, TCOD_ANBAR.NAMES, ISNULL(STUF_DEF.NAME, N' ') + N' ' + ISNULL(INVO"
        "_LST.MANDAH, N' ') AS KALA, TCOD_VAHEDS.NAMES AS VNAMES, HEAD_LST.MOLAH, HEAD_LS"
        "T.USER_NAME, CUST_HESAB.hes, CUST_HESAB.NAME, CUST_HESAB.ADDRESS, STUF_DEF.N_FAN"
        "I, HEAD_LST.SHARAYET, INVO_LST.VAHED_K, CUST_HESAB.TEL FROM TCOD_ANBAR INNER JOI"
        "N STUF_DEF INNER JOIN TCOD_VAHEDS INNER JOIN INVO_LST ON TCOD_VAHEDS.CODE = INVO"
        "_LST.VAHED_K ON STUF_DEF.CODE = INVO_LST.CODE ON TCOD_ANBAR.CODE = INVO_LST.ANBA"
        "R RIGHT OUTER JOIN HEAD_LST ON INVO_LST.NUMBER = HEAD_LST.NUMBER AND INVO_LST.TA"
        "G = HEAD_LST.TAG LEFT OUTER JOIN CUST_HESAB ON HEAD_LST.CUST_NO = CUST_HESAB.hes"
        " WHERE (HEAD_LST.TAG = 2)"
    Caption ="حواله انبار"
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
            Height =2415
            OnFormat ="[Event Procedure]"
            Name ="PageHeaderSection"
            Begin
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =3585
                    Width =2550
                    Height =540
                    FontSize =14
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label88"
                    Caption ="برگ خروج كالا از انبار"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2307
                    Top =435
                    Width =5103
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
                    Left =15
                    Top =540
                    Width =1146
                    Height =375
                    FontSize =10
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
                            Left =1213
                            Top =540
                            Width =570
                            Height =390
                            FontSize =10
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
                    Left =8130
                    Top =540
                    Width =1311
                    Height =375
                    FontSize =10
                    TabIndex =2
                    BorderColor =32768
                    Name ="NUMBER"
                    ControlSource ="NUMBER"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =9480
                            Top =540
                            Width =1020
                            Height =390
                            FontSize =10
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label93"
                            Caption ="شماره حواله:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin Line
                    BorderWidth =3
                    Top =975
                    Width =10521
                    BorderColor =0
                    Name ="Line94"
                End
                Begin Line
                    BorderWidth =3
                    Top =1860
                    Width =10506
                    BorderColor =0
                    Name ="Line99"
                End
                Begin Line
                    BorderWidth =3
                    Left =10485
                    Top =1830
                    Width =0
                    Height =585
                    BorderColor =0
                    Name ="Line130"
                End
                Begin Line
                    BorderWidth =1
                    Left =9927
                    Top =1875
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line131"
                End
                Begin Line
                    BorderWidth =1
                    Left =7824
                    Top =1875
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line132"
                End
                Begin Line
                    BorderWidth =1
                    Left =3369
                    Top =1875
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line133"
                End
                Begin Line
                    BorderWidth =1
                    Left =2184
                    Top =1875
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line134"
                End
                Begin Line
                    BorderWidth =1
                    Left =1155
                    Top =1875
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line135"
                End
                Begin Line
                    BorderWidth =2
                    Top =1875
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line138"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =2190
                    Top =1935
                    Width =1215
                    Height =405
                    FontSize =10
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label142"
                    Caption ="واحد"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =1155
                    Top =1920
                    Width =1035
                    Height =405
                    FontSize =10
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label143"
                    Caption ="مقدار"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =3585
                    Top =1920
                    Width =3975
                    Height =405
                    FontSize =10
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label144"
                    Caption ="شــــــــــرح  كالا"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =9405
                    Top =1920
                    Width =435
                    Height =405
                    FontSize =10
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label145"
                    Caption ="انبار"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =9933
                    Top =1916
                    Width =540
                    Height =405
                    FontSize =10
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label146"
                    Caption ="رديف"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    BorderWidth =3
                    Top =2400
                    Width =10506
                    BorderColor =0
                    Name ="Line147"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =15
                    Top =1890
                    Width =1110
                    Height =480
                    FontSize =10
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label163"
                    Caption ="مقدار كل"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4395
                    Top =1020
                    Width =5406
                    Height =375
                    FontSize =10
                    TabIndex =3
                    BorderColor =32768
                    Name ="NAME"
                    ControlSource ="NAME"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextFontFamily =2
                            Left =9855
                            Top =1020
                            Width =645
                            Height =390
                            FontSize =10
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label171"
                            Caption ="مشتري:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1337
                    Top =1452
                    Width =8586
                    Height =375
                    FontSize =10
                    TabIndex =4
                    BorderColor =32768
                    Name ="ADDRESS"
                    ControlSource ="ADDRESS"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            Visible = NotDefault
                            TextAlign =3
                            TextFontFamily =2
                            Left =9887
                            Top =1452
                            Width =600
                            Height =390
                            FontSize =10
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label173"
                            Caption ="آدرس:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin Label
                    Visible = NotDefault
                    TextAlign =2
                    TextFontFamily =2
                    Left =6000
                    Top =1920
                    Width =510
                    Height =390
                    FontSize =10
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label176"
                    Caption ="كد"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Top =226
                    Width =2336
                    Height =300
                    FontSize =10
                    FontWeight =700
                    TabIndex =5
                    Name ="Text24"
                    ControlSource ="=\"صفحه \" & [Page] & \" از \" & [Pages]"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    BorderWidth =1
                    Left =9330
                    Top =1875
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line189"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =7875
                    Top =1920
                    Width =1455
                    Height =390
                    FontSize =10
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label191"
                    Caption ="كد"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    Left =7531
                    Top =113
                    Width =1836
                    Height =375
                    FontSize =10
                    TabIndex =6
                    BorderColor =32768
                    Name ="FNUMCO"
                    ControlSource ="FNUMCO"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =9360
                            Top =120
                            Width =1125
                            Height =390
                            FontSize =10
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label195"
                            Caption ="شماره داخلي :"
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
                    Left =26
                    Top =1036
                    Width =3561
                    Height =375
                    FontSize =10
                    TabIndex =7
                    BorderColor =32768
                    Name ="TEL"
                    ControlSource ="TEL"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextFontFamily =2
                            Left =3641
                            Top =1036
                            Width =645
                            Height =390
                            FontSize =10
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label205"
                            Caption ="تلفن:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin Label
                    TextFontFamily =2
                    Left =915
                    Top =1515
                    Width =405
                    Height =300
                    FontSize =8
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label170"
                    Caption ="كاربر:"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Top =1515
                    Width =894
                    Height =300
                    FontSize =8
                    TabIndex =8
                    BorderColor =32768
                    Name ="USER_NAME"
                    ControlSource ="USER_NAME"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    Visible = NotDefault
                    TextAlign =2
                    TextFontFamily =2
                    Left =6112
                    Top =15
                    Width =990
                    Height =600
                    FontSize =16
                    BorderColor =32768
                    ForeColor =255
                    Name ="Label208"
                    Caption ="اصلاحيه"
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
                    Left =9375
                    Top =15
                    Width =519
                    Height =405
                    FontSize =10
                    FontWeight =700
                    BorderColor =32768
                    Name ="IANBAR"
                    ControlSource ="IANBAR"
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
                    Left =9943
                    Top =16
                    Width =504
                    Height =405
                    FontSize =10
                    FontWeight =700
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
                    Left =3389
                    Top =16
                    Width =4389
                    Height =405
                    FontSize =10
                    FontWeight =700
                    TabIndex =2
                    BorderColor =32768
                    Name ="KALA"
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
                    Left =2209
                    Top =16
                    Width =1134
                    Height =405
                    FontSize =10
                    FontWeight =700
                    TabIndex =3
                    BorderColor =32768
                    Name ="MEGH"
                    ControlSource ="MEGH"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1185
                    Top =15
                    Width =975
                    Height =405
                    FontSize =10
                    FontWeight =700
                    TabIndex =4
                    BorderColor =32768
                    Name ="VNAMES"
                    ControlSource ="VNAMES"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Top =435
                    Width =10485
                    BorderColor =0
                    Name ="Line100"
                End
                Begin Line
                    BorderWidth =3
                    Left =10485
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line101"
                End
                Begin Line
                    BorderWidth =1
                    Left =9927
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line102"
                End
                Begin Line
                    BorderWidth =1
                    Left =7824
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line103"
                End
                Begin Line
                    BorderWidth =1
                    Left =3369
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line104"
                End
                Begin Line
                    BorderWidth =1
                    Left =2184
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line105"
                End
                Begin Line
                    BorderWidth =1
                    Left =1155
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
                    DecimalPlaces =2
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =57
                    Width =1029
                    Height =405
                    FontSize =10
                    FontWeight =700
                    TabIndex =5
                    BorderColor =32768
                    Name ="MEGHk"
                    ControlSource ="MEGHk"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    RunningSum =1
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5970
                    Top =15
                    Width =564
                    Height =390
                    FontSize =10
                    FontWeight =700
                    TabIndex =6
                    BorderColor =32768
                    Name ="CODE"
                    ControlSource ="CODE"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    BorderWidth =1
                    Left =9330
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
                    Left =7875
                    Top =15
                    Width =1374
                    Height =390
                    FontSize =10
                    FontWeight =700
                    TabIndex =7
                    BorderColor =32768
                    Name ="N_FANI"
                    ControlSource ="N_FANI"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =4110
                    Top =106
                    TabIndex =8
                    Name ="VAHED_K"
                    ControlSource ="VAHED_K"
                    FontName ="WeblogmaYekan"
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            CanGrow = NotDefault
            Height =2994
            Name ="GroupFooter3"
            Begin
                Begin Line
                    BorderWidth =3
                    Width =10510
                    BorderColor =0
                    Name ="Line110"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =9208
                    Top =1260
                    Width =1080
                    Height =405
                    FontSize =10
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label181"
                    Caption ="تنظيم كننده:"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3530
                    Top =1260
                    Width =2601
                    Height =375
                    FontSize =10
                    BorderColor =32768
                    Name ="TAH"
                    ControlSource ="TAH"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextFontFamily =2
                            Left =6185
                            Top =1260
                            Width =1080
                            Height =390
                            FontSize =10
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label155"
                            Caption ="تحويل دهنده:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =60
                    Top =1260
                    Width =2841
                    Height =375
                    FontSize =10
                    TabIndex =1
                    BorderColor =32768
                    Name ="Text158"
                    ControlSource ="MOLAH"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =2903
                            Top =1260
                            Width =555
                            Height =390
                            FontSize =10
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label159"
                            Caption ="توسط:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =9186
                    Top =1785
                    Width =1155
                    Height =390
                    FontSize =10
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label187"
                    Caption ="شماره فاكتور:"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =8820
                    Top =2066
                    Width =1515
                    Height =390
                    FontSize =10
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label188"
                    Caption ="شماره برگ خروج:"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextFontFamily =2
                    Left =6180
                    Top =2081
                    Width =1095
                    Height =390
                    FontSize =10
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label197"
                    Caption ="تحويل گيرنده:"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Top =2754
                    Width =2276
                    FontSize =8
                    TabIndex =2
                    Name ="Text172"
                    ControlSource ="=Format(Farsidate(Now()),\"00\\/00\\/00\") & \" - \" & Format(Now(),\"Long Time\""
                        ")"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =1
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2190
                    Width =1194
                    Height =390
                    FontSize =10
                    FontWeight =700
                    TabIndex =3
                    Name ="Text198"
                    ControlSource ="=Sum([MEGH])"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextFontFamily =2
                    Left =3435
                    Width =810
                    Height =375
                    FontSize =10
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label199"
                    Caption ="جمع :"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =1
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Width =1163
                    Height =391
                    FontSize =10
                    FontWeight =700
                    TabIndex =4
                    Name ="Text203"
                    ControlSource ="=Sum([MEGHk])"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =2
                    Left =450
                    Top =435
                    Width =10065
                    Height =405
                    FontSize =10
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label206"
                    Caption ="طبق صورت به تعداد صحيح و سالم تحويل گرفتم.شماره ماشين:                          "
                        "        شماره بارنامه:    "
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =2
                    Top =855
                    Width =10515
                    Height =345
                    ForeColor =0
                    Name ="Label207"
                    Caption ="درصورت هرگونه مغايرت در محموله ارسالي در زمان تخليه و باحضور راننده اعلام شود در"
                        " غير اين صورت شركت هيچ تعهدي ندارد"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4304
                    Top =61
                    Width =6216
                    Height =375
                    FontSize =10
                    FontWeight =700
                    TabIndex =5
                    BorderColor =32768
                    ForeColor =128
                    Name ="SHARAYET"
                    ControlSource ="SHARAYET"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =6180
                    Top =2591
                    Width =1095
                    Height =390
                    FontSize =10
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label209"
                    Caption ="مدير مالي "
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =9194
                    Top =2577
                    Width =1095
                    Height =390
                    FontSize =10
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label210"
                    Caption ="مدير عامل"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextFontFamily =2
                    Left =2490
                    Top =2096
                    Width =1095
                    Height =390
                    FontSize =10
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label211"
                    Caption ="تلفن راننده:"
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
' See "HAVLAH_ANBAR_02.cls"
