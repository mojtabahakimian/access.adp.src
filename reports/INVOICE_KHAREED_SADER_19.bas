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
    GridY =10
    Width =10155
    DatasheetFontHeight =10
    ItemSuffix =245
    Left =1005
    Top =930
    DatasheetGridlinesColor =12632256
    OnPage ="[Event Procedure]"
    RecSrcDt = Begin
        0x71533ece5d94e540
    End
    RecordSource ="SELECT HEAD_BACK_ANBAR.NUMBER1, HEAD_BACK_ANBAR.NUMBER, HEAD_BACK_ANBAR.DATE_N, "
        "INVO_LST.NUMBER AS INUMBER, HEAD_BACK_ANBAR.HTAG, INVO_LST.RADIF, INVO_LST.CODE,"
        " INVO_LST.MEGH, INVO_LST.MEGHk, INVO_LST.MABL, INVO_LST.MABL_K, STUF_DEF.NAME, H"
        "EAD_BACK_ANBAR.MOLAH, CUST_HESAB.NAME AS HESAB, CUST_HESAB.ADDRESS + N' ' + HEAD"
        "_BACK_ANBAR.MOLAH AS ADDD, CUST_HESAB.TEL, HEAD_BACK_ANBAR.CUST_NO, INVO_LST.N_K"
        "OL, INVO_LST.N_MOIN, HEAD_BACK_ANBAR.FNUMCO, CUST_HESAB.CODE_E, STUF_DEF.TOZIH, "
        "INVO_LST.TOTALARZ, INVO_LST.N_TAF, CUST_HESAB.TOZIH AS BUYER, HEAD_BACK_ANBAR.AN"
        "BARF, HEAD_BACK_ANBAR.MABL_HAZ FROM STUF_DEF RIGHT OUTER JOIN INVO_LST ON STUF_D"
        "EF.CODE = INVO_LST.CODE RIGHT OUTER JOIN CUST_HESAB RIGHT OUTER JOIN HEAD_BACK_A"
        "NBAR ON CUST_HESAB.hes = HEAD_BACK_ANBAR.CUST_NO ON INVO_LST.NUMBER = HEAD_BACK_"
        "ANBAR.NUMBER AND INVO_LST.TAG = HEAD_BACK_ANBAR.HTAG WHERE (HEAD_BACK_ANBAR.HTAG"
        " = 1)"
    Caption ="فاكتور فروش"
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
        Begin PageBreak
            Width =283
        End
        Begin BreakLevel
            ControlSource ="RADIF"
        End
        Begin PageHeader
            Height =2538
            Name ="PageHeader"
            Begin
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    TextAlign =2
                    TextFontFamily =34
                    Width =10155
                    Height =540
                    FontSize =14
                    FontWeight =400
                    BackColor =10092543
                    ForeColor =0
                    Name ="Label88"
                    Caption ="KASHI ABADIS  CO.(L.T.D)"
                    FontName ="@Arial Unicode MS"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =1
                    TextFontFamily =34
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =1
                    Left =1592
                    Top =1053
                    Width =1191
                    Height =435
                    FontSize =12
                    BorderColor =32768
                    Name ="ANBARF"
                    ControlSource ="ANBARF"
                    FontName ="@Arial Unicode MS"
                    Begin
                        Begin Label
                            TextFontFamily =34
                            NumeralShapes =1
                            Left =8
                            Top =1053
                            Width =1545
                            Height =435
                            FontSize =12
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label93"
                            Caption ="INVOICE NO:"
                            FontName ="@Arial Unicode MS"
                        End
                    End
                End
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =34
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =1
                    Left =1757
                    Top =566
                    Width =6066
                    Height =450
                    FontSize =12
                    TabIndex =1
                    BorderColor =32768
                    Name ="BUYER"
                    ControlSource ="BUYER"
                    FontName ="@Arial Unicode MS"
                    Begin
                        Begin Label
                            TextFontFamily =34
                            NumeralShapes =1
                            Left =4
                            Top =566
                            Width =1695
                            Height =435
                            FontSize =12
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label95"
                            Caption ="SELER NAME:"
                            FontName ="@Arial Unicode MS"
                        End
                    End
                End
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =34
                    ReadingOrder =2
                    ScrollBarAlign =1
                    NumeralShapes =1
                    Left =851
                    Top =1474
                    Width =6846
                    Height =450
                    FontSize =12
                    TabIndex =2
                    BorderColor =32768
                    Name ="ADDD"
                    ControlSource ="ADDD"
                    FontName ="@Arial Unicode MS"
                    Begin
                        Begin Label
                            TextFontFamily =34
                            NumeralShapes =1
                            Top =1474
                            Width =795
                            Height =450
                            FontSize =12
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label97"
                            Caption ="ADD:"
                            FontName ="@Arial Unicode MS"
                        End
                    End
                End
                Begin Rectangle
                    BackStyle =1
                    Top =2025
                    Width =10098
                    Height =513
                    BackColor =10092543
                    Name ="Box199"
                End
                Begin Line
                    Left =5370
                    Top =2025
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line131"
                End
                Begin Line
                    Left =540
                    Top =2025
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line133"
                End
                Begin Line
                    Left =4320
                    Top =2025
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line135"
                End
                Begin Line
                    Left =8549
                    Top =2025
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line136"
                End
                Begin Line
                    Left =10101
                    Top =2017
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line137"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =34
                    Left =8677
                    Top =2069
                    Width =1350
                    Height =405
                    FontSize =8
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label139"
                    Caption ="PRICE"
                    FontName ="@Arial Unicode MS"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =34
                    Left =6802
                    Top =2073
                    Width =1695
                    Height =405
                    FontSize =8
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label140"
                    Caption ="UNIT PRICE F.O.T"
                    FontName ="@Arial Unicode MS"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =34
                    Left =4342
                    Top =2084
                    Width =1020
                    Height =405
                    FontSize =8
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label141"
                    Caption ="SQ.METER"
                    FontName ="@Arial Unicode MS"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =34
                    Left =5415
                    Top =2070
                    Width =1275
                    Height =405
                    FontSize =8
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label142"
                    Caption ="CARTON NO."
                    FontName ="@Arial Unicode MS"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =34
                    Left =1597
                    Top =2084
                    Width =1710
                    Height =405
                    FontSize =8
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label144"
                    Caption ="DESCRIPTION"
                    FontName ="@Arial Unicode MS"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =34
                    Left =7
                    Top =2080
                    Width =510
                    Height =405
                    FontSize =8
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label146"
                    Caption ="NO"
                    FontName ="@Arial Unicode MS"
                End
                Begin Line
                    Left =6720
                    Top =2025
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line197"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =2
                    TextFontFamily =34
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =1
                    IMESentenceMode =3
                    Left =9426
                    Top =563
                    Width =651
                    Height =390
                    FontSize =10
                    TabIndex =3
                    Name ="NUMBER"
                    ControlSource ="NUMBER"
                    FontName ="@Arial Unicode MS"
                End
                Begin TextBox
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =8735
                    Top =963
                    Width =1356
                    Height =405
                    FontSize =12
                    TabIndex =4
                    Name ="DATE_N"
                    ControlSource ="DATE_N"
                    FontName ="@Arial Unicode MS"
                    InputMask ="####/##/##"
                End
                Begin Label
                    TextFontFamily =34
                    NumeralShapes =1
                    Left =7905
                    Top =960
                    Width =825
                    Height =435
                    FontSize =12
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label228"
                    Caption ="DATE :"
                    FontName ="@Arial Unicode MS"
                End
                Begin Label
                    TextFontFamily =34
                    NumeralShapes =1
                    Left =7880
                    Top =566
                    Width =1650
                    Height =375
                    FontSize =8
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="ddd"
                    Caption ="Delivery Order no.:"
                    FontName ="@Arial Unicode MS"
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
                    TextFontFamily =34
                    ReadingOrder =1
                    ScrollBarAlign =1
                    NumeralShapes =1
                    Left =10
                    Top =16
                    Width =504
                    Height =405
                    FontSize =10
                    BorderColor =32768
                    Name ="RADIFwe"
                    ControlSource ="=1"
                    FontName ="@Arial Unicode MS"
                End
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =34
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =1
                    Left =566
                    Top =16
                    Width =3759
                    Height =405
                    FontSize =10
                    TabIndex =1
                    BorderColor =32768
                    Name ="CODE1"
                    ControlSource ="TOZIH"
                    FontName ="@Arial Unicode MS"
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =2
                    TextFontFamily =34
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =1
                    Left =5427
                    Width =1254
                    Height =405
                    FontSize =10
                    TabIndex =2
                    BorderColor =32768
                    Name ="MEGH"
                    ControlSource ="MEGH"
                    FontName ="@Arial Unicode MS"
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =2
                    TextFontFamily =34
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =1
                    Left =4350
                    Top =16
                    Width =954
                    Height =405
                    FontSize =10
                    TabIndex =3
                    BorderColor =32768
                    Name ="MEGHk"
                    ControlSource ="MEGHk"
                    FontName ="@Arial Unicode MS"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =34
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =1
                    Left =6795
                    Top =15
                    Width =1704
                    Height =405
                    FontSize =10
                    TabIndex =4
                    BorderColor =32768
                    Name ="N_TAF"
                    ControlSource ="N_TAF"
                    Format ="General Number"
                    FontName ="@Arial Unicode MS"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =34
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =1
                    Left =8620
                    Width =1419
                    Height =405
                    FontSize =10
                    TabIndex =5
                    BorderColor =32768
                    Name ="MABL_K"
                    ControlSource ="TOTALARZ"
                    Format ="General Number"
                    FontName ="@Arial Unicode MS"
                End
                Begin Line
                    BorderWidth =1
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line220"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =2
                    RunningSum =2
                    TextAlign =2
                    TextFontFamily =34
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =1
                    Left =5670
                    Top =56
                    Width =969
                    Height =210
                    FontSize =10
                    TabIndex =6
                    BorderColor =32768
                    Name ="MEGHS"
                    ControlSource ="MEGH"
                    FontName ="@Arial Unicode MS"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =2
                    RunningSum =2
                    TextAlign =2
                    TextFontFamily =34
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4593
                    Top =72
                    Width =669
                    Height =210
                    FontSize =10
                    TabIndex =7
                    BorderColor =32768
                    Name ="MEGHkS"
                    ControlSource ="MEGHk"
                    FontName ="@Arial Unicode MS"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    RunningSum =2
                    TextAlign =2
                    TextFontFamily =34
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =8578
                    Top =56
                    Width =1419
                    Height =210
                    FontSize =10
                    TabIndex =8
                    BorderColor =32768
                    Name ="MABL_KS"
                    ControlSource ="TOTALARZ"
                    Format ="#,###"
                    FontName ="@Arial Unicode MS"
                End
            End
        End
        Begin PageFooter
            Height =3180
            Name ="PageFooter"
            Begin
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =2
                    TextAlign =2
                    TextFontFamily =34
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =1
                    Left =8605
                    Top =60
                    Width =1419
                    Height =405
                    FontSize =10
                    BorderColor =32768
                    Name ="MABL_HAZ1"
                    ControlSource ="=Round([Forms]![HEAD_LST_KHAREED1_SADER]![MABL_HAZ]/[Forms]![HEAD_LST_KHAREED1_S"
                        "ADER]![ARZD],2)"
                    Format ="General Number"
                    FontName ="@Arial Unicode MS"
                End
                Begin Label
                    Visible = NotDefault
                    TextAlign =2
                    TextFontFamily =34
                    NumeralShapes =1
                    Left =2070
                    Top =60
                    Width =1845
                    Height =405
                    FontSize =10
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label208"
                    Caption ="st of transport"
                    FontName ="@Arial Unicode MS"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =34
                    NumeralShapes =1
                    Left =7440
                    Top =1080
                    Width =2055
                    Height =405
                    FontSize =10
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label211"
                    Caption ="Dr A Damaki"
                    FontName ="@Arial Unicode MS"
                End
                Begin Line
                    BorderWidth =2
                    Top =435
                    Width =10091
                    BorderColor =0
                    Name ="Line212"
                End
                Begin Line
                    BorderWidth =2
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line213"
                End
                Begin Line
                    BorderWidth =2
                    Width =10091
                    BorderColor =0
                    Name ="Line214"
                End
                Begin Rectangle
                    Top =885
                    Width =10098
                    Height =1579
                    Name ="Box215"
                End
                Begin Rectangle
                    BackStyle =1
                    Top =2361
                    Width =10098
                    Height =784
                    BackColor =10092543
                    Name ="Box216"
                End
                Begin Label
                    TextFontFamily =34
                    NumeralShapes =1
                    Top =2580
                    Width =4822
                    Height =375
                    FontSize =10
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label217"
                    Caption =".FACTORY:YAZD IRAN 4 th PAKNEJAD STR"
                    FontName ="@Arial Unicode MS"
                End
                Begin Label
                    TextFontFamily =34
                    NumeralShapes =1
                    Left =5670
                    Top =2385
                    Width =3285
                    Height =375
                    FontSize =10
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label218"
                    Caption ="TEL:(+9835) 37269168-37284271"
                    FontName ="@Arial Unicode MS"
                End
                Begin Label
                    TextFontFamily =34
                    NumeralShapes =1
                    Left =5670
                    Top =2805
                    Width =3165
                    Height =375
                    FontSize =10
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label219"
                    Caption ="FAX:(+9835) 37284277"
                    FontName ="@Arial Unicode MS"
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =2
                    TextFontFamily =34
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =1
                    Left =5412
                    Top =510
                    Width =1254
                    Height =405
                    FontSize =10
                    TabIndex =1
                    BorderColor =32768
                    Name ="Text232"
                    ControlSource ="=[MEGHS]"
                    FontName ="@Arial Unicode MS"
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =2
                    TextFontFamily =34
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =1
                    Left =4335
                    Top =510
                    Width =954
                    Height =405
                    FontSize =10
                    TabIndex =2
                    BorderColor =32768
                    Name ="Text233"
                    ControlSource ="=[MEGHkS]"
                    FontName ="@Arial Unicode MS"
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =2
                    TextFontFamily =34
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =1
                    Left =8605
                    Top =510
                    Width =1419
                    Height =405
                    FontSize =10
                    TabIndex =3
                    BorderColor =32768
                    Name ="Text234"
                    ControlSource ="=[MABL_KS]+[MABL_HAZ1]"
                    Format ="General Number"
                    FontName ="@Arial Unicode MS"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =34
                    NumeralShapes =1
                    Left =2115
                    Top =510
                    Width =1755
                    Height =405
                    FontSize =10
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label235"
                    Caption ="Total"
                    FontName ="@Arial Unicode MS"
                End
                Begin Line
                    BorderWidth =2
                    Top =885
                    Width =10091
                    BorderColor =0
                    Name ="Line236"
                End
                Begin Line
                    BorderWidth =2
                    Top =450
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line237"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =34
                    NumeralShapes =1
                    Left =315
                    Top =1605
                    Width =1875
                    Height =405
                    FontSize =10
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label240"
                    Caption =".Purchase  Dep"
                    FontName ="@Arial Unicode MS"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =34
                    NumeralShapes =1
                    Left =3975
                    Top =1605
                    Width =1875
                    Height =405
                    FontSize =10
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label241"
                    Caption =".Accounting Dep"
                    FontName ="@Arial Unicode MS"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =34
                    NumeralShapes =1
                    Left =7485
                    Top =1605
                    Width =1965
                    Height =405
                    FontSize =10
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label242"
                    Caption =".Comerical Manager"
                    FontName ="@Arial Unicode MS"
                End
            End
        End
    End
End
CodeBehindForm
' See "INVOICE_KHAREED_SADER_19.cls"
