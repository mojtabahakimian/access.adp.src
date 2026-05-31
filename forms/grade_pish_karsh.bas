Version =20
VersionRequired =20
Begin Form
    AllowDeletions = NotDefault
    AllowAdditions = NotDefault
    AllowEdits = NotDefault
    DefaultView =2
    AllowUpdating =4
    ViewsAllowed =2
    TabularCharSet =178
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =7088
    RowHeight =330
    ItemSuffix =30
    Left =825
    Top =7065
    Right =13545
    Bottom =9555
    DatasheetForeColor =33554432
    DatasheetGridlinesColor =12632256
    MaxRecords =100000
    RecSrcDt = Begin
        0x5c1a7064fdbae540
    End
    RecordSource ="SELECT SUM(INVO_LST.MABL_K - INVO_LST.N_MOIN + INVO_LST.IMBAA) AS GMAB, HEAD_LST"
        ".NUMBER, HEAD_LST.DATE_N, HEAD_LST.CUST_NO, HEAD_LST.MOLAH, HEAD_LST.DEPATMAN, H"
        "EAD_LST.SHIFT, HEAD_LST.CUST_KIND, HEAD_LST.USER_NAME, HEAD_LST.TAMIR, CUST_HESA"
        "B.NAME, Visit_route.HES FROM HEAD_LST INNER JOIN INVO_LST ON HEAD_LST.NUMBER = I"
        "NVO_LST.NUMBER AND HEAD_LST.TAG = INVO_LST.TAG INNER JOIN CUST_HESAB ON HEAD_LST"
        ".CUST_NO = CUST_HESAB.hes INNER JOIN Visit_route ON CUST_HESAB.ROUTE_NAME = Visi"
        "t_route.ROUTE_NAME WHERE (HEAD_LST.TAG = 20) AND (HEAD_LST.TAMIR <> 2) AND (HEAD"
        "_LST.TAMIR <> 3) AND (HEAD_LST.SGN1 = 1) GROUP BY HEAD_LST.NUMBER, HEAD_LST.DATE"
        "_N, HEAD_LST.CUST_NO, HEAD_LST.MOLAH, HEAD_LST.DEPATMAN, HEAD_LST.SHIFT, HEAD_LS"
        "T.CUST_KIND, HEAD_LST.USER_NAME, HEAD_LST.TAMIR, CUST_HESAB.NAME, Visit_route.HE"
        "S HAVING (Visit_route.HES = N'212-1-206')"
    Caption ="ليست پيش فاکتورهاي تاييد شده"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Tahoma"
    AllowFormView =0
    RecordSourceQualifier ="dbo"
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
        Begin ComboBox
            SpecialEffect =2
            TextFontCharSet =178
            Width =1701
            LabelX =-1701
            FontName ="Tahoma"
        End
        Begin Section
            Height =4208
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =2834
                    Top =1247
                    ColumnWidth =1830
                    ColumnOrder =3
                    Name ="GMAB"
                    ControlSource ="GMAB"
                    Format ="#,###"
                    ConditionalFormat = Begin
                        0x0100000066000000010000000000000004000000000000000200000001000000 ,
                        0x00000000ff99cc00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x300000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5681
                            Top =1247
                            Width =555
                            Height =240
                            Name ="Label20"
                            Caption ="مبلغ"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =2834
                    Top =1587
                    ColumnOrder =0
                    TabIndex =1
                    Name ="NUMBER"
                    ControlSource ="NUMBER"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5220
                            Top =1590
                            Width =1320
                            Height =240
                            Name ="Label21"
                            Caption ="شماره پيش فاكتور"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =2834
                    Top =1927
                    ColumnOrder =1
                    TabIndex =2
                    Name ="DATE_N"
                    ControlSource ="DATE_N"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5351
                            Top =1927
                            Width =885
                            Height =240
                            Name ="Label22"
                            Caption ="تاريخ "
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =2834
                    Top =2267
                    ColumnOrder =2
                    TabIndex =3
                    Name ="CUST_NO"
                    ControlSource ="CUST_NO"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5051
                            Top =2267
                            Width =1185
                            Height =240
                            Name ="Label23"
                            Caption ="شماره مشتري"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =2834
                    Top =2607
                    ColumnWidth =3675
                    ColumnOrder =9
                    TabIndex =4
                    Name ="MOLAH"
                    ControlSource ="MOLAH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5486
                            Top =2607
                            Width =750
                            Height =240
                            Name ="Label24"
                            Caption ="ملاحظات"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =2834
                    Top =3968
                    ColumnWidth =2130
                    ColumnOrder =8
                    TabIndex =8
                    Name ="USER_NAME"
                    ControlSource ="USER_NAME"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5201
                            Top =3968
                            Width =1035
                            Height =240
                            Name ="Label28"
                            Caption ="نام کاربر"
                        End
                    End
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =2834
                    Top =2948
                    ColumnWidth =1605
                    ColumnOrder =5
                    TabIndex =5
                    Name ="DEPATMAN"
                    ControlSource ="DEPATMAN"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT DEPATMAN, DEPNAME FROM DEPART"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5636
                            Top =2948
                            Width =600
                            Height =240
                            Name ="Label25"
                            Caption ="دپاتمان"
                        End
                    End
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =2834
                    Top =3288
                    ColumnOrder =6
                    TabIndex =6
                    Name ="SHIFT"
                    ControlSource ="SHIFT"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT SHIFT_ID, SHNAME FROM SHIFT"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5636
                            Top =3288
                            Width =600
                            Height =240
                            Name ="Label26"
                            Caption ="شيفت"
                        End
                    End
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =2834
                    Top =3628
                    ColumnOrder =7
                    TabIndex =7
                    Name ="CUST_KIND"
                    ControlSource ="CUST_KIND"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT CUST_COD, CUSTKNAME FROM CUSTKIND"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5276
                            Top =3628
                            Width =960
                            Height =240
                            Name ="Label27"
                            Caption ="نوع مشتري"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =2437
                    Top =170
                    Width =1986
                    Height =405
                    ColumnOrder =4
                    TabIndex =9
                    BackColor =65535
                    Name ="TAMIR"
                    ControlSource ="TAMIR"
                    RowSourceType ="Value List"
                    RowSource ="0;\"عادي\";1;\"رزرو شده\";2;\"تبديل به حواله\";3;\"لغو شد\";4;\"رزو قطعي\""
                    ColumnWidths ="0"
                    DefaultValue ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1035
                            Top =120
                            Width =585
                            Height =240
                            Name ="Label29"
                            Caption ="وضعيت"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "grade_pish_karsh.cls"
