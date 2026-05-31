Version =20
VersionRequired =20
Begin Form
    DefaultView =2
    AllowUpdating =4
    TabularCharSet =178
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =5669
    RowHeight =330
    ItemSuffix =8
    Left =270
    Top =600
    Right =14850
    Bottom =8235
    DatasheetForeColor =33554432
    DatasheetGridlinesColor =12632256
    MaxRecords =100000
    RecSrcDt = Begin
        0x50d498db8ca7e340
    End
    RecordSource ="SELECT HEAD_LST_1.TAG, HEAD_LST.TAG AS Expr1, HEAD_LST.NUMBER, HEAD_LST_1.NUMBER"
        " AS Expr2, HEAD_LST_1.CUST_NO, HEAD_LST.CUST_NO AS Expr3, CUST_HESAB.NAME AS n1,"
        " CUST_HESAB_1.NAME AS n2 FROM HEAD_LST INNER JOIN HEAD_LST HEAD_LST_1 ON HEAD_LS"
        "T.NUMBER = HEAD_LST_1.NUMBER AND HEAD_LST.CUST_NO <> HEAD_LST_1.CUST_NO LEFT OUT"
        "ER JOIN CUST_HESAB CUST_HESAB_1 ON HEAD_LST.CUST_NO = CUST_HESAB_1.hes LEFT OUTE"
        "R JOIN CUST_HESAB ON HEAD_LST_1.CUST_NO = CUST_HESAB.hes WHERE (HEAD_LST_1.TAG ="
        " 12) AND (HEAD_LST.TAG = 1)"
    Caption ="ليست رسيد و فاكتورهاي داراي مغايرت"
    DatasheetFontName ="Tahoma"
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
        End
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            OldBorderStyle =0
            Width =1701
            LabelX =-1701
            FontName ="Tahoma"
            AsianLineBreak =255
        End
        Begin Section
            Height =3471
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =3628
                    Top =1530
                    ColumnWidth =1065
                    Name ="NUMBER"
                    ControlSource ="NUMBER"
                    OnDblClick ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1927
                            Top =1530
                            Width =1020
                            Height =240
                            Name ="Label2"
                            Caption ="شماره فاكتور"
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
                    Left =3628
                    Top =1870
                    ColumnWidth =1050
                    TabIndex =1
                    Name ="Expr2"
                    ControlSource ="Expr2"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1927
                            Top =1870
                            Width =555
                            Height =240
                            Name ="Label3"
                            Caption ="شماره رسيد"
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
                    Left =3628
                    Top =2211
                    ColumnWidth =1665
                    TabIndex =2
                    Name ="CUST_NO"
                    ControlSource ="CUST_NO"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1927
                            Top =2211
                            Width =1185
                            Height =240
                            Name ="Label4"
                            Caption ="شماره مشتري فاكتور"
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
                    Left =3628
                    Top =2551
                    ColumnWidth =1650
                    TabIndex =3
                    Name ="Expr3"
                    ControlSource ="Expr3"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1927
                            Top =2551
                            Width =555
                            Height =240
                            Name ="Label5"
                            Caption ="شماره مشتري رسيد"
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
                    Left =3628
                    Top =2891
                    ColumnWidth =2700
                    TabIndex =4
                    Name ="n1"
                    ControlSource ="n1"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1927
                            Top =2891
                            Width =315
                            Height =240
                            Name ="Label6"
                            Caption ="نام مشتري در فاكتور"
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
                    Left =3628
                    Top =3231
                    ColumnWidth =2580
                    TabIndex =5
                    Name ="n2"
                    ControlSource ="n2"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1927
                            Top =3231
                            Width =315
                            Height =240
                            Name ="Label7"
                            Caption ="نام مشتري در رسيد"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "R_F_MO.cls"
