Version =20
VersionRequired =20
Begin Form
    DefaultView =0
    AllowUpdating =4
    ViewsAllowed =2
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =5669
    RowHeight =360
    ItemSuffix =7
    Left =270
    Top =600
    Right =7380
    Bottom =3765
    DatasheetForeColor =33554432
    DatasheetGridlinesColor =12632256
    MaxRecords =10000000
    RecSrcDt = Begin
        0x74ada9f4182be440
    End
    RecordSource ="SELECT HEAD_LST.NUMBER, HEAD_LST.TAG, CUST_HESAB.NAME, HEAD_LST.DATE_N, HEAD_LST"
        ".USER_NAME, TAGCOD.BARGAH, CUST_HESAB.hes FROM HEAD_LST INNER JOIN CUST_HESAB ON"
        " HEAD_LST.CUST_NO = CUST_HESAB.hes INNER JOIN TAGCOD ON HEAD_LST.TAG = TAGCOD.CO"
        "DE LEFT OUTER JOIN OTHER_DTL ON HEAD_LST.NUMBER = OTHER_DTL.NUMBER AND HEAD_LST."
        "TAG = OTHER_DTL.TAG WHERE (OTHER_DTL.NUMBER IS NULL) AND (HEAD_LST.DATE_N BETWEE"
        "N 1 AND 2)"
    Caption ="ليست برگه هاي بدون قبض باسكول"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
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
                    Left =3344
                    Top =737
                    ColumnWidth =1155
                    ColumnOrder =0
                    Name ="NUMBER"
                    ControlSource ="NUMBER"
                    OnDblClick ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1643
                            Top =737
                            Width =1020
                            Height =240
                            Name ="Label0"
                            Caption ="*شماره فاكتور"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =3344
                    Top =1077
                    ColumnWidth =735
                    ColumnOrder =1
                    TabIndex =1
                    Name ="TAG"
                    ControlSource ="TAG"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1643
                            Top =1077
                            Width =690
                            Height =240
                            Name ="Label1"
                            Caption ="برچسب"
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
                    Left =3344
                    Top =1417
                    ColumnWidth =5115
                    ColumnOrder =3
                    TabIndex =2
                    Name ="NAME"
                    ControlSource ="NAME"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1650
                            Top =1410
                            Width =870
                            Height =240
                            Name ="Label2"
                            Caption ="نام مشتري"
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
                    Left =3344
                    Top =1757
                    ColumnWidth =930
                    ColumnOrder =4
                    TabIndex =3
                    Name ="DATE_N"
                    ControlSource ="DATE_N"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1643
                            Top =1757
                            Width =885
                            Height =240
                            Name ="Label3"
                            Caption ="تاريخ فاكتور"
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
                    Left =3344
                    Top =2097
                    ColumnOrder =6
                    TabIndex =4
                    Name ="USER_NAME"
                    ControlSource ="USER_NAME"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1643
                            Top =2097
                            Width =1035
                            Height =240
                            Name ="Label4"
                            Caption ="نام كاربر"
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
                    Left =3344
                    Top =2437
                    ColumnOrder =5
                    TabIndex =5
                    Name ="BARGAH"
                    ControlSource ="BARGAH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1643
                            Top =2437
                            Width =750
                            Height =240
                            Name ="Label5"
                            Caption ="نوع برگه"
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
                    Left =3798
                    Top =3231
                    ColumnOrder =2
                    TabIndex =6
                    Name ="hes"
                    ControlSource ="hes"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2100
                            Top =3225
                            Width =840
                            Height =240
                            Name ="Label6"
                            Caption ="كد مشتري"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "HAVALAH_WITHUT_BASKOOL.cls"
