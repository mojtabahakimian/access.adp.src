Version =20
VersionRequired =20
Begin Form
    AllowDesignChanges = NotDefault
    DefaultView =2
    AllowUpdating =4
    ViewsAllowed =2
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =5160
    RowHeight =495
    DatasheetFontHeight =10
    ItemSuffix =19
    Left =1380
    Top =6285
    Right =2805
    Bottom =7725
    DatasheetGridlinesColor =12632256
    MaxRecords =100000000
    RecSrcDt = Begin
        0xbe46f497f65de640
    End
    RecordSource ="PRICE_ELAMIETF_EXCEPTION"
    Caption ="استثناء"
    DatasheetFontName ="Tahoma"
    OnError ="[Event Procedure]"
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
            Height =3194
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1927
                    Top =113
                    ColumnOrder =0
                    Name ="EXCEPTION_ID"
                    ControlSource ="EXCEPTION_ID"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =113
                            Width =1170
                            Height =240
                            Name ="Label1"
                            Caption ="شماره"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1927
                    Top =453
                    ColumnWidth =645
                    ColumnOrder =2
                    TabIndex =1
                    Name ="PETID"
                    ControlSource ="PETID"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =453
                            Width =510
                            Height =240
                            Name ="Label3"
                            Caption ="کد  تخفيف"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1927
                    Top =1133
                    ColumnWidth =1830
                    ColumnOrder =3
                    TabIndex =2
                    Name ="EXCEPTION_TF1"
                    ControlSource ="EXCEPTION_TF1"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =1133
                            Width =1275
                            Height =240
                            Name ="Label7"
                            Caption ="تخفيف اول"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1927
                    Top =1474
                    ColumnWidth =1830
                    ColumnOrder =4
                    TabIndex =3
                    Name ="EXCEPTION_TF2"
                    ControlSource ="EXCEPTION_TF2"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =1474
                            Width =1275
                            Height =240
                            Name ="Label9"
                            Caption ="تخفيف دوم"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =93
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1927
                    Top =1814
                    ColumnOrder =8
                    TabIndex =4
                    Name ="TR_DATE"
                    ControlSource ="TR_DATE"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =1814
                            Width =750
                            Height =240
                            Name ="Label11"
                            Caption ="تاريخ ايجاد"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1927
                    Top =2154
                    Width =2490
                    ColumnWidth =1680
                    ColumnOrder =5
                    TabIndex =5
                    Name ="USERNAME"
                    ControlSource ="USERNAME"
                    DefaultValue ="UCurrentUser()"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =2154
                            Width =885
                            Height =240
                            Name ="Label13"
                            Caption ="کاربر"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1927
                    Top =2494
                    Height =255
                    ColumnWidth =2430
                    ColumnOrder =6
                    TabIndex =6
                    Name ="CRT"
                    ControlSource ="CRT"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =2494
                            Width =375
                            Height =240
                            Name ="Label15"
                            Caption ="ايجاد"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1927
                    Top =2834
                    ColumnOrder =7
                    TabIndex =7
                    Name ="UID"
                    ControlSource ="UID"
                    DefaultValue ="=UCurrentUserIDD()"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =2834
                            Width =345
                            Height =240
                            Name ="Label17"
                            Caption ="کاربر"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =3
                    Left =1530
                    Top =850
                    Width =1680
                    Height =255
                    ColumnWidth =7395
                    ColumnOrder =1
                    TabIndex =8
                    BackColor =52479
                    ConditionalFormat = Begin
                        0x0100000066000000010000000000000004000000000000000200000001000000 ,
                        0x00000000ffcc0000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x300000000000
                    End
                    Name ="CODE"
                    ControlSource ="CODE"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT CODE, NAME, CODE AS Expr1 FROM STUF_DEF WHERE (NOT (MENUIT IS NULL)) ORDE"
                        "R BY NAME"
                    ColumnWidths ="0;5670"
                    Tag ="51"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextFontCharSet =178
                            Left =3465
                            Top =855
                            Width =1695
                            Height =990
                            Name ="Label18"
                            Caption ="کالاي مشمول تخفيف استثنا"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "PRICE_ELAMIETF_EXCEPTION_FRM.cls"
