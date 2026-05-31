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
    ItemSuffix =3
    DatasheetForeColor =33554432
    DatasheetGridlinesColor =12632256
    MaxRecords =1000000
    RecSrcDt = Begin
        0x21a0c1d9aedce340
    End
    RecordSource ="SELECT SUM(PAY_GETD.MABL) AS MABLs, PAY_GETD.CUST_NO, CUST_HESAB.NAME FROM PAY_G"
        "ETD INNER JOIN CUST_HESAB ON PAY_GETD.CUST_NO = CUST_HESAB.hes WHERE (PAY_GETD.N"
        "_KOL2 IS NULL) AND (PAY_GETD.N_KOL3 IS NULL) AND (PAY_GETD.N_KOL = 112) GROUP BY"
        " PAY_GETD.CUST_NO, CUST_HESAB.NAME ORDER BY SUM(PAY_GETD.MABL) DESC"
    Caption ="مانده اسناد دريافتني نزد بانك و صندوق"
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
            Height =1701
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
            End
        End
    End
End
CodeBehindForm
' See "CHEK_MAND.cls"
