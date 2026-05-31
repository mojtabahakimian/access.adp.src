Version =20
VersionRequired =20
Begin Form
    AllowDeletions = NotDefault
    AllowAdditions = NotDefault
    KeyPreview = NotDefault
    OrderByOn = NotDefault
    AllowEdits = NotDefault
    DefaultView =2
    AllowUpdating =4
    ViewsAllowed =2
    TabularCharSet =178
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =5687
    RowHeight =345
    ItemSuffix =15
    Left =1125
    Top =1380
    Right =8850
    Bottom =5385
    DatasheetForeColor =33554432
    DatasheetGridlinesColor =12632256
    MaxRecords =100000
    RecSrcDt = Begin
        0x1653386c7285e540
    End
    RecordSource ="SELECT HEAD_LST.NUMBER, HEAD_LST.TAG, HEAD_LST.DATE_N, HEAD_LST.TAH, HEAD_LST.CU"
        "ST_NO, CUST_HESAB.NAME, HEAD_LST.FNUMCO, HEAD_LST.TAMIR, HEAD_LST.MOLAH, HEAD_LS"
        "T.SHARAYET, HEAD_LST.SADER, HEAD_LST.DEPATMAN, HEAD_LST.USER_NAME FROM HEAD_LST "
        "INNER JOIN CUST_HESAB ON HEAD_LST.CUST_NO = CUST_HESAB.hes WHERE (HEAD_LST.TAG ="
        " 2) AND (NOT (HEAD_LST.NUMBER IN (SELECT HEAD_LST.NUMBER FROM HEAD_LST WHERE ((("
        "HEAD_LST.TAG) = 13))))) ORDER BY HEAD_LST.NUMBER"
    Caption ="ليست حواله ها"
    OnCurrent ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
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
            Height =5625
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
                    OnDblClick ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1757
                            Top =850
                            Width =1020
                            Height =240
                            Name ="Label0"
                            Caption ="شماره حواله"
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
                    ColumnWidth =1020
                    ColumnOrder =1
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
                    ColumnOrder =4
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
                    ColumnOrder =2
                    TabIndex =3
                    Name ="CUST_NO"
                    ControlSource ="CUST_NO"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =1757
                            Top =2211
                            Width =1185
                            Height =240
                            Name ="Label4"
                            Caption ="شماره مشتري"
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
                    ColumnWidth =4620
                    ColumnOrder =3
                    TabIndex =4
                    Name ="NAME"
                    ControlSource ="NAME"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1755
                            Top =2550
                            Width =870
                            Height =240
                            Name ="Label5"
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
                    Left =3458
                    Top =2912
                    ColumnWidth =1125
                    TabIndex =5
                    Name ="FNUMCO"
                    ControlSource ="FNUMCO"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1755
                            Top =2910
                            Width =1020
                            Height =240
                            Name ="Label7"
                            Caption ="شماره داخلي"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    ReadingOrder =1
                    Left =3203
                    Top =283
                    ColumnWidth =1065
                    TabIndex =6
                    Name ="TAMIR"
                    ControlSource ="TAMIR"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1665
                            Top =255
                            Width =960
                            Height =240
                            Name ="Label9"
                            Caption ="بارگيري شده"
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
                    Top =3274
                    ColumnWidth =5880
                    TabIndex =7
                    Name ="SHARAYET"
                    ControlSource ="SHARAYET"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1755
                            Top =3272
                            Width =1020
                            Height =240
                            Name ="Label11"
                            Caption ="ملاحظات"
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
                    Left =2154
                    Top =3798
                    ColumnWidth =1515
                    TabIndex =8
                    Name ="SADER"
                    ControlSource ="SADER"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4922
                            Top =3802
                            Width =765
                            Height =240
                            Name ="Label12"
                            Caption ="نوع فروش"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1984
                    Top =4478
                    TabIndex =9
                    Name ="DEPATMAN"
                    ControlSource ="DEPATMAN"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4786
                            Top =4478
                            Width =600
                            Height =240
                            Name ="Label13"
                            Caption ="واحد"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1360
                    Top =5385
                    ColumnWidth =1605
                    TabIndex =10
                    Name ="USER_NAME"
                    ControlSource ="USER_NAME"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3727
                            Top =5385
                            Width =1035
                            Height =240
                            Name ="Label14"
                            Caption ="کاربر"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "HAVLIST.cls"
