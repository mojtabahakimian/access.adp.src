Version =20
VersionRequired =20
Begin Form
    AllowDeletions = NotDefault
    AllowAdditions = NotDefault
    KeyPreview = NotDefault
    AllowEdits = NotDefault
    DefaultView =2
    AllowUpdating =4
    ViewsAllowed =2
    TabularCharSet =178
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =6270
    RowHeight =345
    ItemSuffix =18
    Left =345
    Top =4200
    Right =6885
    Bottom =9360
    DatasheetForeColor =33554432
    DatasheetGridlinesColor =12632256
    MaxRecords =100000
    RecSrcDt = Begin
        0x8574648516ebe340
    End
    RecordSource ="SELECT HEAD_LST.NUMBER, HEAD_LST.TAG, HEAD_LST.DATE_N, HEAD_LST.TAH, INVO_LST.AN"
        "BAR, INVO_LST.CODE, INVO_LST.MEGHk, MEGH_RAIS.NUMBER1, ISNULL(MEGH_RAIS.MEGHkS, "
        "0) AS MEGHKS, INVO_LST.VAHED_K, STUF_DEF.NAME AS KALA, TCOD_ANBAR.NAMES AS ANBAR"
        "N, INVO_LST.MEGHk - ISNULL(MEGH_RAIS.MEGHkS, 0) AS MAND FROM HEAD_LST INNER JOIN"
        " INVO_LST ON HEAD_LST.NUMBER = INVO_LST.NUMBER AND HEAD_LST.TAG = INVO_LST.TAG I"
        "NNER JOIN STUF_DEF ON INVO_LST.CODE = STUF_DEF.CODE INNER JOIN TCOD_ANBAR ON INV"
        "O_LST.ANBAR = TCOD_ANBAR.CODE LEFT OUTER JOIN MEGH_RAIS ON INVO_LST.id = MEGH_RA"
        "IS.RADAH AND HEAD_LST.NUMBER = MEGH_RAIS.NUMBER1 WHERE (HEAD_LST.TAG = 23)"
    DatasheetFontName ="Tahoma"
    OnError ="[Event Procedure]"
    AllowFormView =0
    Begin
        Begin Label
            BackStyle =0
            TextFontCharSet =178
            FontName ="Tahoma"
        End
        Begin CheckBox
            SpecialEffect =2
            LabelX =230
            LabelY =-30
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
            Height =5172
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    ColumnHidden = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1190
                    Top =56
                    ColumnWidth =1350
                    ColumnOrder =0
                    Name ="NUMBER"
                    ControlSource ="NUMBER"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3105
                            Top =60
                            Width =1245
                            Height =240
                            Name ="Label0"
                            Caption ="شماره درخواست"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1245
                    Top =736
                    ColumnWidth =840
                    ColumnOrder =1
                    TabIndex =1
                    Name ="DATE_N"
                    ControlSource ="DATE_N"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3174
                            Top =736
                            Width =525
                            Height =240
                            Name ="Label2"
                            Caption ="تاريخ "
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1245
                    Top =1076
                    ColumnWidth =2070
                    ColumnOrder =2
                    TabIndex =2
                    Name ="TAH"
                    ControlSource ="TAH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3172
                            Top =1081
                            Width =885
                            Height =240
                            Name ="Label3"
                            Caption ="در خواست كننده "
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1247
                    Top =1364
                    ColumnWidth =1350
                    ColumnOrder =4
                    TabIndex =3
                    Name ="MEGHk"
                    ControlSource ="MEGHk"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3248
                            Top =1360
                            Width =1065
                            Height =405
                            Name ="Label13"
                            Caption ="مقدار درخواستي"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1303
                    Top =1817
                    ColumnWidth =1590
                    ColumnOrder =5
                    TabIndex =4
                    Name ="MEGHKS"
                    ControlSource ="MEGHKS"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3976
                            Top =1814
                            Width =1440
                            Height =405
                            Name ="Label14"
                            Caption ="مقدار  خريداري شده"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1247
                    Top =2267
                    ColumnWidth =3585
                    ColumnOrder =3
                    TabIndex =5
                    Name ="KALA"
                    ControlSource ="KALA"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4139
                            Top =2267
                            Width =510
                            Height =240
                            Name ="Label15"
                            Caption ="كالا"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1303
                    Top =3004
                    ColumnWidth =3195
                    ColumnOrder =7
                    TabIndex =6
                    Name ="ANBARN"
                    ControlSource ="ANBARN"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3955
                            Top =3004
                            Width =750
                            Height =240
                            Name ="Label16"
                            Caption ="انبار"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1303
                    Top =2664
                    ColumnWidth =750
                    ColumnOrder =6
                    TabIndex =7
                    Name ="MAND"
                    ControlSource ="MAND"
                    ConditionalFormat = Begin
                        0x0100000066000000010000000000000004000000000000000200000001000000 ,
                        0x00000000ffcc9900000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x300000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4135
                            Top =2664
                            Width =570
                            Height =240
                            Name ="Label17"
                            Caption ="مانده"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "HAVLISTDAR_SUB.cls"
