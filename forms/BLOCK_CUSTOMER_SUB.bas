Version =20
VersionRequired =20
Begin Form
    AllowDeletions = NotDefault
    DefaultView =2
    AllowUpdating =4
    ViewsAllowed =2
    TabularCharSet =178
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =6165
    RowHeight =330
    ItemSuffix =12
    Left =465
    Top =1110
    Right =10380
    Bottom =7260
    DatasheetForeColor =33554432
    DatasheetGridlinesColor =12632256
    MaxRecords =100000
    RecSrcDt = Begin
        0x889e0a648e70e340
    End
    ServerFilter ="ENDBLK = 0"
    RecordSource ="BLOCK_CUSTOMER"
    Caption ="ليست مشتريان مسدود شده"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Tahoma"
    OnError ="[Event Procedure]"
    AllowFormView =0
    RecordSourceQualifier ="dbo"
    Begin
        Begin Label
            BackStyle =0
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
            Height =2730
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
                    Left =1920
                    Top =343
                    ColumnOrder =3
                    TabIndex =1
                    Name ="STBLKDT"
                    ControlSource ="STBLKDT"
                    DefaultValue ="=FARSIDATE(Now())"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =340
                            Width =1350
                            Height =240
                            Name ="Label3"
                            Caption ="تاريخ مسدود شدن"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1927
                    Top =1076
                    ColumnWidth =1095
                    ColumnOrder =4
                    TabIndex =2
                    Name ="ENDBLKDT"
                    ControlSource ="ENDBLKDT"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =120
                            Top =1078
                            Width =990
                            Height =240
                            Name ="Label5"
                            Caption ="تاريخ باز شدن"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =255
                    Left =1927
                    Top =1133
                    ColumnWidth =930
                    ColumnOrder =2
                    TabIndex =3
                    Name ="ENDBLK"
                    ControlSource ="ENDBLK"
                    OnClick ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =120
                            Top =1140
                            Width =1320
                            Height =240
                            Name ="Label7"
                            Caption ="پايان مسدود شدن"
                        End
                    End
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =215
                    TextFontCharSet =178
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =226
                    Top =680
                    Width =3126
                    ColumnWidth =4515
                    ColumnOrder =1
                    TabIndex =4
                    Name ="HEST"
                    ControlSource ="HES"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT RTRIM(CAST(N_KOL AS nvarchar)) + '-' + RTRIM(CAST(NUMBER AS nvarchar)) + "
                        "'-' + RTRIM(CAST(TNUMBER AS nvarchar)) AS hes, ISNULL(NAME, N' ') AS Expr1 FROM "
                        "TDETA_HES ORDER BY ISNULL(NAME, N' ')"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =226
                            Top =680
                            Width =1020
                            Height =240
                            Name ="Label8"
                            Caption ="نام حساب"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =3677
                    Width =2475
                    ColumnWidth =1455
                    ColumnOrder =0
                    Name ="HES"
                    ControlSource ="HES"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT RTRIM(CAST(N_KOL AS nvarchar)) + '-' + RTRIM(CAST(NUMBER AS nvarchar)) + "
                        "'-' + RTRIM(CAST(TNUMBER AS nvarchar)) AS hes FROM TDETA_HES"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1870
                            Top =7
                            Width =1020
                            Height =240
                            Name ="Label1"
                            Caption ="كد حساب"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =247
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1643
                    Top =1303
                    ColumnWidth =3330
                    TabIndex =5
                    Name ="COMM"
                    ControlSource ="COMM"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4440
                            Top =1305
                            Width =675
                            Height =240
                            Name ="Label9"
                            Caption ="توضيحات"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2607
                    Top =2040
                    TabIndex =6
                    Name ="CUSER"
                    ControlSource ="CUSER"
                    DefaultValue ="=UCurrentUser()"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =5385
                            Top =2040
                            Width =780
                            Height =405
                            Name ="Label10"
                            Caption ="كاربر ايجاد كننده"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2324
                    Top =2324
                    TabIndex =7
                    Name ="EUSER"
                    ControlSource ="EUSER"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =5115
                            Top =2325
                            Width =810
                            Height =405
                            Name ="Label11"
                            Caption ="كاربر رفع كننده"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "BLOCK_CUSTOMER_SUB.cls"
