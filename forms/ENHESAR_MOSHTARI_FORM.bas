Version =20
VersionRequired =20
Begin Form
    AutoCenter = NotDefault
    DividingLines = NotDefault
    DefaultView =2
    AllowUpdating =4
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =9418
    RowHeight =375
    ItemSuffix =15
    Left =270
    Top =225
    Right =7380
    Bottom =3360
    DatasheetGridlinesColor =12632256
    MaxRecords =100000000
    RecSrcDt = Begin
        0x57ccf47b76c1e540
    End
    RecordSource ="ENHESAR_MOSHTARI"
    Caption ="تعريف انحصار"
    DatasheetFontName ="Tahoma"
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
            Height =3619
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1927
                    Top =793
                    ColumnWidth =1440
                    ColumnOrder =3
                    TabIndex =1
                    Name ="EN_STDATE"
                    ControlSource ="EN_STDATE"
                    DefaultValue ="=farsidate(Date())"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =793
                            Width =930
                            Height =240
                            Name ="Label5"
                            Caption ="تاريخ شروع انحصار"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1927
                    Top =1133
                    ColumnWidth =1305
                    ColumnOrder =4
                    TabIndex =2
                    Name ="EN_ENDEN"
                    ControlSource ="EN_ENDEN"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =1133
                            Width =855
                            Height =240
                            Name ="Label7"
                            Caption ="تاريخ پايان انحصار"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1927
                    Top =1474
                    ColumnWidth =2070
                    ColumnOrder =5
                    TabIndex =3
                    Name ="EN_CDATE"
                    ControlSource ="EN_CDATE"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =1474
                            Width =855
                            Height =240
                            Name ="Label9"
                            Caption ="تاريخ ايجاد"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1927
                    Top =1814
                    ColumnWidth =690
                    ColumnOrder =6
                    TabIndex =4
                    Name ="USERID"
                    ControlSource ="USERID"
                    DefaultValue ="[Forms]![baseknow]![USERCOD]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =1814
                            Width =630
                            Height =240
                            Name ="Label11"
                            Caption ="کد کاربر"
                        End
                    End
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =1927
                    Top =113
                    ColumnWidth =1530
                    ColumnOrder =0
                    Name ="EN_CITY"
                    ControlSource ="EN_CITY"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT CITYCODE, CITYNAME FROM TCOD_CITY ORDER BY CITYNAME"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =113
                            Width =705
                            Height =240
                            Name ="Label1"
                            Caption ="شهر"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =3
                    Left =2428
                    Top =2850
                    Width =6171
                    Height =315
                    ColumnWidth =5340
                    ColumnOrder =2
                    TabIndex =5
                    BackColor =13434828
                    ForeColor =8388608
                    Name ="EN_CUST_CO"
                    ControlSource ="EN_CUST_CO"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT     hes, NAME AS nam, hes AS Expr1\015\012FROM         CUST_HESAB"
                    ColumnWidths ="0;4536;567"
                    AfterUpdate ="[Event Procedure]"
                    Tag ="0"
                    OnNotInList ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =8683
                            Top =2850
                            Width =735
                            Height =300
                            ForeColor =8388608
                            Name ="Label19"
                            Caption =" مشتري"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    AutoExpand = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =215
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =283
                    Top =2834
                    Width =2091
                    Height =315
                    ColumnWidth =1665
                    ColumnOrder =1
                    TabIndex =6
                    BackColor =13434828
                    ForeColor =8388608
                    Name ="EN_CUST_CO22"
                    ControlSource ="EN_CUST_CO"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT hes FROM CUST_HESAB"
                    OnNotInList ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Top =2834
                            Width =735
                            Height =300
                            ForeColor =8388608
                            Name ="Label14"
                            Caption ="کد مشتري"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "ENHESAR_MOSHTARI_FORM.cls"
