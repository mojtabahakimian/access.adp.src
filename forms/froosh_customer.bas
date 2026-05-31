Version =20
VersionRequired =20
Begin Form
    DefaultView =2
    AllowUpdating =4
    ViewsAllowed =2
    TabularCharSet =178
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =5669
    RowHeight =375
    ItemSuffix =19
    Left =1050
    Top =930
    Right =15090
    Bottom =8655
    DatasheetForeColor =33554432
    DatasheetGridlinesColor =12632256
    MaxRecords =100000
    RecSrcDt = Begin
        0x7a262116d4c3e440
    End
    RecordSource ="SELECT TOP 100 PERCENT HEAD_LST.TAG, HEAD_LST.DATE_N, HEAD_LST.CUST_NO, STUF_DEF"
        ".NAME, INVO_LST.MEGHk, INVO_LST.MABL_K, HEAD_LST.NUMBER, CUST_HESAB.NAME AS cust"
        "omer, INVO_LST.MEGH, TCOD_VAHEDS.NAMES AS VAHED, INVO_LST.MABL, INVO_LST.N_KOL, "
        "INVO_LST.N_MOIN FROM HEAD_LST INNER JOIN INVO_LST ON HEAD_LST.NUMBER = INVO_LST."
        "NUMBER AND HEAD_LST.TAG = INVO_LST.TAG INNER JOIN STUF_DEF ON INVO_LST.CODE = ST"
        "UF_DEF.CODE INNER JOIN CUST_HESAB ON HEAD_LST.CUST_NO = CUST_HESAB.hes INNER JOI"
        "N TCOD_VAHEDS ON INVO_LST.VAHED_K = TCOD_VAHEDS.CODE WHERE (HEAD_LST.CUST_NO = N"
        "'115-1-1') ORDER BY HEAD_LST.DATE_N DESC"
    Caption ="فروش كالا"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Tahoma"
    AllowFormView =0
    Begin
        Begin Label
            BackStyle =0
            TextFontCharSet =178
            FontName ="Tahoma"
        End
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            OldBorderStyle =0
            TextFontCharSet =178
            Width =1701
            LabelX =-1701
            FontName ="Tahoma"
            AsianLineBreak =255
        End
        Begin Section
            Height =5798
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2551
                    Top =1247
                    ColumnWidth =1020
                    ColumnOrder =0
                    Name ="DATE_N"
                    ControlSource ="DATE_N"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =850
                            Top =1247
                            Width =885
                            Height =240
                            Name ="Label1"
                            Caption ="تاريخ "
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2551
                    Top =1587
                    ColumnWidth =1230
                    ColumnOrder =1
                    TabIndex =1
                    Name ="CUST_NO"
                    ControlSource ="CUST_NO"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =850
                            Top =1587
                            Width =1185
                            Height =240
                            Name ="Label2"
                            Caption ="شماره مشتري"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2551
                    Top =1927
                    ColumnWidth =4905
                    ColumnOrder =2
                    TabIndex =2
                    Name ="NAME"
                    ControlSource ="NAME"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =850
                            Top =1927
                            Width =555
                            Height =240
                            Name ="Label3"
                            Caption ="نام كالا"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2551
                    Top =2267
                    ColumnOrder =5
                    TabIndex =3
                    Name ="MEGHk"
                    ControlSource ="MEGHk"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =855
                            Top =2265
                            Width =690
                            Height =240
                            Name ="Label4"
                            Caption ="مقدار كل"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2551
                    Top =2607
                    ColumnWidth =1470
                    ColumnOrder =7
                    TabIndex =4
                    Name ="MABL_K"
                    ControlSource ="MABL_K"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =850
                            Top =2607
                            Width =705
                            Height =240
                            Name ="Label5"
                            Caption ="مبلغ"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2551
                    Top =2948
                    ColumnWidth =1005
                    ColumnOrder =10
                    TabIndex =5
                    Name ="NUMBER"
                    ControlSource ="NUMBER"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =850
                            Top =2948
                            Width =1020
                            Height =240
                            Name ="Label6"
                            Caption ="شماره فاكتور"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2556
                    Top =3465
                    ColumnWidth =2610
                    ColumnOrder =11
                    TabIndex =6
                    Name ="customer"
                    ControlSource ="customer"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =855
                            Top =3465
                            Width =870
                            Height =240
                            Name ="Label8"
                            Caption ="نام مشتري"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2664
                    Top =4478
                    ColumnWidth =765
                    ColumnOrder =3
                    TabIndex =7
                    Name ="MEGH"
                    ControlSource ="MEGH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =963
                            Top =4478
                            Width =630
                            Height =240
                            Name ="Label10"
                            Caption ="مقدار"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2664
                    Top =4838
                    ColumnWidth =930
                    ColumnOrder =4
                    TabIndex =8
                    Name ="VAHED"
                    ControlSource ="VAHED"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =963
                            Top =4838
                            Width =630
                            Height =240
                            Name ="Label12"
                            Caption ="واحد"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2556
                    Top =3855
                    ColumnWidth =945
                    ColumnOrder =6
                    TabIndex =9
                    Name ="MABL"
                    ControlSource ="MABL"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =855
                            Top =3855
                            Width =705
                            Height =240
                            Name ="Label14"
                            Caption ="في"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2664
                    Top =5198
                    ColumnWidth =600
                    ColumnOrder =8
                    TabIndex =10
                    Name ="N_KOL"
                    ControlSource ="N_KOL"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =963
                            Top =5198
                            Width =630
                            Height =240
                            Name ="Label16"
                            Caption ="درصد"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2664
                    Top =5558
                    ColumnWidth =1050
                    ColumnOrder =9
                    TabIndex =11
                    Name ="N_MOIN"
                    ControlSource ="N_MOIN"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =963
                            Top =5558
                            Width =630
                            Height =240
                            Name ="Label18"
                            Caption ="تخفيف"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "froosh_customer.cls"
