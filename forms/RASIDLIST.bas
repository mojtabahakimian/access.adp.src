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
    Width =5669
    RowHeight =345
    ItemSuffix =12
    Left =270
    Top =600
    Right =7725
    Bottom =4350
    DatasheetForeColor =33554432
    DatasheetGridlinesColor =12632256
    MaxRecords =100000
    RecSrcDt = Begin
        0x7b09eded8dbce340
    End
    RecordSource ="SELECT HEAD_LST.NUMBER, HEAD_LST.TAG, HEAD_LST.DATE_N, HEAD_LST.TAH, HEAD_LST.CU"
        "ST_NO, CUST_HESAB.NAME, HEAD_LST.USER_NAME, HEAD_LST.FNUMCO, HEAD_LST.MOLAH FROM"
        " HEAD_LST LEFT OUTER JOIN CUST_HESAB ON HEAD_LST.CUST_NO = CUST_HESAB.hes WHERE "
        "(HEAD_LST.TAG = 1) AND (NOT (HEAD_LST.NUMBER IN (SELECT HEAD_LST.NUMBER FROM HEA"
        "D_LST WHERE (((HEAD_LST.TAG) = 12))))) ORDER BY HEAD_LST.NUMBER"
    Caption ="ليست رسيد هايي كه فاكتور آن صادر نشده است"
    OnCurrent ="[Event Procedure]"
    DatasheetFontName ="Tahoma"
    OnKeyPress ="[Event Procedure]"
    OnError ="[Event Procedure]"
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
            Height =4548
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
                    Left =3458
                    Top =850
                    ColumnWidth =1035
                    ColumnOrder =0
                    Name ="NUMBER"
                    ControlSource ="NUMBER"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1757
                            Top =850
                            Width =1020
                            Height =240
                            Name ="Label0"
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
                    Left =3458
                    Top =1530
                    ColumnWidth =840
                    ColumnOrder =3
                    TabIndex =1
                    Name ="DATE_N"
                    ControlSource ="DATE_N"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1757
                            Top =1530
                            Width =885
                            Height =240
                            Name ="Label2"
                            Caption ="تاريخ "
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
                    Left =3458
                    Top =1870
                    ColumnOrder =7
                    TabIndex =2
                    Name ="TAH"
                    ControlSource ="TAH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1757
                            Top =1870
                            Width =1065
                            Height =240
                            Name ="Label3"
                            Caption ="تحويل "
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
                    Left =3458
                    Top =2211
                    ColumnWidth =1230
                    ColumnOrder =5
                    TabIndex =3
                    Name ="CUST_NO"
                    ControlSource ="CUST_NO"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1757
                            Top =2211
                            Width =1185
                            Height =240
                            Name ="Label4"
                            Caption ="شماره فروشنده"
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
                    Left =3458
                    Top =2551
                    ColumnWidth =4410
                    ColumnOrder =6
                    TabIndex =4
                    Name ="NAME"
                    ControlSource ="NAME"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1755
                            Top =2550
                            Width =1170
                            Height =240
                            Name ="Label5"
                            Caption ="نام نام فروشنده"
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
                    Left =3458
                    Top =2912
                    ColumnWidth =1170
                    ColumnOrder =1
                    TabIndex =5
                    Name ="USER_NAME"
                    ControlSource ="USER_NAME"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1755
                            Top =2911
                            Width =870
                            Height =240
                            Name ="Label7"
                            Caption ="كاربر"
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
                    Left =3458
                    Top =3273
                    ColumnWidth =1125
                    ColumnOrder =2
                    TabIndex =6
                    Name ="FNUMCO"
                    ControlSource ="FNUMCO"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1755
                            Top =3272
                            Width =870
                            Height =240
                            Name ="Label9"
                            Caption ="شماره داخلي"
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
                    Left =3288
                    Top =4025
                    ColumnWidth =3210
                    ColumnOrder =4
                    TabIndex =7
                    Name ="Text10"
                    ControlSource ="MOLAH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1587
                            Top =4025
                            Width =1065
                            Height =240
                            Name ="Label11"
                            Caption ="تحويل  دهنده"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "RASIDLIST.cls"
