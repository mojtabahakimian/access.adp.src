Version =20
VersionRequired =20
Begin Form
    AllowAdditions = NotDefault
    OrderByOn = NotDefault
    DefaultView =2
    AllowUpdating =4
    ViewsAllowed =2
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =6990
    RowHeight =330
    ItemSuffix =32
    Left =570
    Top =3690
    Right =15345
    Bottom =10935
    DatasheetGridlinesColor =12632256
    MaxRecords =100000000
    RecSrcDt = Begin
        0x655d8e3232bfe540
    End
    RecordSource ="SELECT SCAID, CUST_NO, SCADATE, SCAJAMF, SCARASFR, SCARASTAV, SCARASPAY, SCATAFA"
        "VOT, SCAMABNADATE, SCAMABLA, SCAMABP, CTIM, USERCO, MONTH, dbo.UIIF(SCAMABLA, '>"
        "=', 0, SCAMABLA, 0) AS TAKHIR, dbo.UIIF(SCAMABLA, '<', 0, ABS(SCAMABLA), 0) AS T"
        "AGIL FROM SSM_CUST_ANALYS ORDER BY SCAJAMF DESC"
    Caption ="ليست محاسبات تاخير و تاجيل"
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
            Width =1701
            LabelX =-1701
            FontName ="Tahoma"
        End
        Begin Section
            Height =8617
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1927
                    Top =113
                    ColumnWidth =765
                    ColumnOrder =0
                    Name ="SCAID"
                    ControlSource ="SCAID"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =113
                            Width =540
                            Height =240
                            Name ="Label1"
                            Caption ="شماره"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1927
                    Top =453
                    Width =2475
                    ColumnWidth =1035
                    ColumnOrder =1
                    TabIndex =1
                    Name ="CUST_NO"
                    ControlSource ="CUST_NO"
                    OnDblClick ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =450
                            Width =930
                            Height =240
                            Name ="Label3"
                            Caption ="کد مشتري*"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1927
                    Top =793
                    ColumnWidth =1020
                    ColumnOrder =3
                    TabIndex =2
                    Name ="SCADATE"
                    ControlSource ="SCADATE"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =793
                            Width =765
                            Height =240
                            Name ="Label5"
                            Caption ="تاريخ"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1927
                    Top =1133
                    Width =2475
                    ColumnWidth =1425
                    ColumnOrder =4
                    TabIndex =3
                    Name ="SCAJAMF"
                    ControlSource ="SCAJAMF"
                    Format ="#,##0.00;-#,##0.00"
                    ConditionalFormat = Begin
                        0x0100000066000000010000000000000003000000000000000200000001000000 ,
                        0x00000000ff99cc00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x300000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =1140
                            Width =810
                            Height =240
                            Name ="Label7"
                            Caption ="جمع فاکتور"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1927
                    Top =1474
                    ColumnWidth =1080
                    ColumnOrder =8
                    TabIndex =4
                    Name ="SCARASFR"
                    ControlSource ="SCARASFR"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =1470
                            Width =975
                            Height =240
                            Name ="Label9"
                            Caption ="راس فاکتورها"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1927
                    Top =1814
                    ColumnWidth =915
                    ColumnOrder =9
                    TabIndex =5
                    Name ="SCARASTAV"
                    ControlSource ="SCARASTAV"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =1814
                            Width =960
                            Height =240
                            Name ="Label11"
                            Caption ="راس توافق"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1927
                    Top =2154
                    ColumnWidth =1020
                    ColumnOrder =10
                    TabIndex =6
                    Name ="SCARASPAY"
                    ControlSource ="SCARASPAY"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =2154
                            Width =960
                            Height =240
                            Name ="Label13"
                            Caption ="راس پرداخت"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1927
                    Top =2494
                    ColumnWidth =510
                    ColumnOrder =7
                    TabIndex =7
                    Name ="SCATAFAVOT"
                    ControlSource ="SCATAFAVOT"
                    ConditionalFormat = Begin
                        0x0100000066000000010000000000000003000000000000000200000001000000 ,
                        0x00000000ff660000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x300000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =2494
                            Width =1065
                            Height =240
                            Name ="Label15"
                            Caption ="تاخير"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1927
                    Top =2834
                    ColumnWidth =1020
                    ColumnOrder =11
                    TabIndex =8
                    Name ="SCAMABNADATE"
                    ControlSource ="SCAMABNADATE"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =2834
                            Width =1290
                            Height =240
                            Name ="Label17"
                            Caption ="تاريخ مبنا"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1927
                    Top =3174
                    Width =2475
                    ColumnWidth =1350
                    ColumnOrder =5
                    TabIndex =9
                    Name ="TAKHIR"
                    ControlSource ="TAKHIR"
                    Format ="#,##0.00;-#,##0.00"
                    ConditionalFormat = Begin
                        0x0100000066000000010000000000000003000000000000000200000001000000 ,
                        0x00000000ccffcc00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x300000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =3174
                            Width =870
                            Height =240
                            Name ="Label19"
                            Caption ="مبلغ تاخير"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1927
                    Top =3514
                    ColumnOrder =12
                    TabIndex =10
                    Name ="SCAMABP"
                    ControlSource ="SCAMABP"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =3510
                            Width =1365
                            Height =240
                            Name ="Label21"
                            Caption ="درصد کسر و اضافه"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1927
                    Top =3855
                    ColumnWidth =2070
                    ColumnOrder =15
                    TabIndex =11
                    Name ="CTIM"
                    ControlSource ="CTIM"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =3855
                            Width =450
                            Height =240
                            Name ="Label23"
                            Caption ="تاريخ"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1927
                    Top =4195
                    ColumnWidth =720
                    ColumnOrder =13
                    TabIndex =12
                    Name ="USERCO"
                    ControlSource ="USERCO"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =4195
                            Width =690
                            Height =240
                            Name ="Label25"
                            Caption ="کد کاربر"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1927
                    Top =4535
                    ColumnWidth =360
                    ColumnOrder =14
                    TabIndex =13
                    Name ="MONTH"
                    ControlSource ="MONTH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =4535
                            Width =615
                            Height =240
                            Name ="Label27"
                            Caption ="ماه"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =3
                    Top =4875
                    Width =6171
                    Height =315
                    ColumnWidth =3570
                    ColumnOrder =2
                    TabIndex =14
                    BackColor =13434828
                    ForeColor =8388608
                    Name ="Combo28"
                    ControlSource ="CUST_NO"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT     hes, NAME AS nam, hes AS Expr1\015\012FROM         CUST_HESAB"
                    ColumnWidths ="0;4536;567"
                    Tag ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6255
                            Top =4875
                            Width =735
                            Height =300
                            ForeColor =8388608
                            Name ="Label29"
                            Caption =" مشتري:"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =2324
                    Top =6009
                    Width =2475
                    ColumnOrder =6
                    TabIndex =15
                    Name ="TAGIL"
                    ControlSource ="TAGIL"
                    Format ="#,##0.00;-#,##0.00"
                    ConditionalFormat = Begin
                        0x0100000066000000010000000000000003000000000000000200000001000000 ,
                        0x00000000ccffcc00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x300000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =510
                            Top =6009
                            Width =870
                            Height =240
                            Name ="Label31"
                            Caption ="مبلغ تاجيل"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "SSM_CUST_ANALYS_form.cls"
