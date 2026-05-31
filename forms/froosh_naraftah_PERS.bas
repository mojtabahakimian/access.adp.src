Version =20
VersionRequired =20
Begin Form
    AllowDeletions = NotDefault
    AllowAdditions = NotDefault
    DefaultView =2
    AllowUpdating =4
    ViewsAllowed =2
    TabularCharSet =178
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =4940
    RowHeight =360
    ItemSuffix =18
    Left =495
    Top =225
    Right =14490
    Bottom =5505
    DatasheetForeColor =33554432
    DatasheetGridlinesColor =12632256
    MaxRecords =1000000
    RecSrcDt = Begin
        0x7324e55bef90e440
    End
    RecordSource ="SELECT STUF_DEF.CODE, STUF_DEF.NAME, STUF_DEF.N_FANI, STUF_DEF.TOZIH, STUF_DEF.B"
        "_SEF, STUF_DEF.N_SEF, STUF_DEF.MIN_M, STUF_DEF.MAX_M, STUF_DEF.MABL_F, MOGUDI_KO"
        "L_ANBARHA.MANDAH FROM STUF_DEF INNER JOIN MOGUDI_KOL_ANBARHA ON STUF_DEF.CODE = "
        "MOGUDI_KOL_ANBARHA.CODE WHERE (NOT (STUF_DEF.CODE IN (SELECT dbo.INVO_LST.CODE F"
        "ROM dbo.INVO_LST INNER JOIN dbo.HEAD_LST ON dbo.INVO_LST.NUMBER = dbo.HEAD_LST.N"
        "UMBER AND dbo.INVO_LST.TAG = dbo.HEAD_LST.TAG WHERE (dbo.INVO_LST.TAG = 2) AND ("
        "dbo.HEAD_LST.CUST_NO = N'124-1-1') GROUP BY dbo.INVO_LST.CODE)))"
    Caption ="ليست كالاهاي فروش نرفته (راكد)"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Tahoma"
    OnError ="[Event Procedure]"
    AllowFormView =0
    RecordSourceQualifier ="dbo"
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
            Height =3359
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =1
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1927
                    Top =113
                    Width =2490
                    ColumnWidth =705
                    ColumnOrder =0
                    Name ="CODE"
                    ControlSource ="CODE"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =113
                            Width =510
                            Height =240
                            Name ="Label1"
                            Caption ="كد كالا"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =1
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1927
                    Top =453
                    Width =2490
                    ColumnWidth =6165
                    ColumnOrder =1
                    TabIndex =1
                    Name ="NAME"
                    ControlSource ="NAME"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =450
                            Width =540
                            Height =240
                            Name ="Label3"
                            Caption ="نام كالا"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =1
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1927
                    Top =793
                    Width =2490
                    ColumnWidth =1890
                    ColumnOrder =8
                    TabIndex =2
                    Name ="N_FANI"
                    ControlSource ="N_FANI"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =793
                            Width =630
                            Height =240
                            Name ="Label5"
                            Caption ="باركد"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =1
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1927
                    Top =1133
                    Width =2490
                    ColumnWidth =1620
                    ColumnOrder =9
                    TabIndex =3
                    Name ="TOZIH"
                    ControlSource ="TOZIH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =1140
                            Width =675
                            Height =240
                            Name ="Label7"
                            Caption ="توضيحات"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =1
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1927
                    Top =1474
                    ColumnWidth =1110
                    ColumnOrder =4
                    TabIndex =4
                    Name ="B_SEF"
                    ControlSource ="B_SEF"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =1470
                            Width =690
                            Height =240
                            Name ="Label9"
                            Caption ="في خرده"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =1
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1927
                    Top =1814
                    ColumnOrder =5
                    TabIndex =5
                    Name ="N_SEF"
                    ControlSource ="N_SEF"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =1815
                            Width =1020
                            Height =240
                            Name ="Label11"
                            Caption ="نقطه سفارش"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =1
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1927
                    Top =2154
                    ColumnOrder =6
                    TabIndex =6
                    Name ="MIN_M"
                    ControlSource ="MIN_M"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =2160
                            Width =1125
                            Height =240
                            Name ="Label13"
                            Caption ="حداقل موجودي"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =1
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1927
                    Top =2494
                    ColumnWidth =1680
                    ColumnOrder =7
                    TabIndex =7
                    Name ="MAX_M"
                    ControlSource ="MAX_M"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =2490
                            Width =1365
                            Height =240
                            Name ="Label15"
                            Caption ="قيمت مصرف كننده"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =1
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1927
                    Top =2834
                    ColumnWidth =1305
                    ColumnOrder =3
                    TabIndex =8
                    Name ="MABL_F"
                    ControlSource ="MABL_F"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =2835
                            Width =1200
                            Height =240
                            Name ="Label17"
                            Caption ="في عمده فروش"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =1
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1920
                    Top =3118
                    ColumnWidth =915
                    ColumnOrder =2
                    TabIndex =9
                    Name ="MANDAH"
                    ControlSource ="MANDAH"
                    ConditionalFormat = Begin
                        0x0100000066000000010000000000000006000000000000000200000001000000 ,
                        0x00000000ccffcc00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x300000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =3119
                            Width =1200
                            Height =240
                            Name ="Label19"
                            Caption ="موجودي"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "froosh_naraftah_PERS.cls"
