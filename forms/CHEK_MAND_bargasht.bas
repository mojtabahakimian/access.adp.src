Version =20
VersionRequired =20
Begin Form
    DefaultView =0
    AllowUpdating =4
    ViewsAllowed =2
    TabularCharSet =178
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =5669
    RowHeight =375
    ItemSuffix =5
    DatasheetForeColor =33554432
    DatasheetGridlinesColor =12632256
    MaxRecords =1000000
    RecSrcDt = Begin
        0xb903131b701ae440
    End
    RecordSource ="SELECT SUM(PAY_GETD.MABL) AS MABLs, PAY_GETD.CUST_NO, CUST_HESAB.NAME, COUNT(PAY"
        "_GETD.CUST_NO) AS TEDAD FROM PAY_GETD INNER JOIN CUST_HESAB ON PAY_GETD.CUST_NO "
        "= CUST_HESAB.hes WHERE (NOT (PAY_GETD.N_KOL2 IS NULL)) GROUP BY PAY_GETD.CUST_NO"
        ", CUST_HESAB.NAME ORDER BY SUM(PAY_GETD.MABL) DESC"
    Caption ="جمع چكهاي برگشتي"
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
            Height =1770
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2834
                    Top =453
                    ColumnWidth =2625
                    ColumnOrder =2
                    Name ="MABLs"
                    ControlSource ="MABLs"
                    Format ="#,###"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1360
                            Top =453
                            Width =600
                            Height =240
                            Name ="Label0"
                            Caption ="مبلغ"
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
                    Left =2834
                    Top =793
                    ColumnWidth =1845
                    ColumnOrder =0
                    TabIndex =1
                    Name ="CUST_NO"
                    ControlSource ="CUST_NO"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1360
                            Top =793
                            Width =1185
                            Height =240
                            Name ="Label1"
                            Caption ="كد مشتري"
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
                    Left =2834
                    Top =1133
                    ColumnWidth =4005
                    ColumnOrder =1
                    TabIndex =2
                    Name ="NAME"
                    ControlSource ="NAME"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1365
                            Top =1140
                            Width =855
                            Height =240
                            Name ="Label2"
                            Caption ="مانده اسناد"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2839
                    Top =1530
                    TabIndex =3
                    Name ="TEDAD"
                    ControlSource ="TEDAD"
                    Format ="#,###"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =2
                            Left =1365
                            Top =1530
                            Width =600
                            Height =240
                            Name ="Label4"
                            Caption ="تعداد"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "CHEK_MAND_bargasht.cls"
