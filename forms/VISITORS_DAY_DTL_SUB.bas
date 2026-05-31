Version =20
VersionRequired =20
Begin Form
    AllowEdits = NotDefault
    DefaultView =2
    AllowUpdating =4
    ViewsAllowed =2
    TabularCharSet =178
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =8361
    RowHeight =330
    ItemSuffix =13
    Left =405
    Top =1650
    Right =14985
    Bottom =9210
    DatasheetGridlinesColor =12632256
    MaxRecords =1000000
    RecSrcDt = Begin
        0xf9cc1df6d075e440
    End
    RecordSource ="VISITORS_DAY_DTL"
    DatasheetFontName ="Tahoma"
    OnError ="[Event Procedure]"
    AllowFormView =0
    RecordSourceQualifier ="dbo"
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
        Begin ComboBox
            SpecialEffect =2
            TextFontCharSet =178
            Width =1701
            LabelX =-1701
            FontName ="Tahoma"
        End
        Begin Section
            Height =4025
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1927
                    Top =113
                    Width =2490
                    ColumnWidth =1417
                    ColumnOrder =0
                    Name ="HES"
                    ControlSource ="HES"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =113
                            Width =1095
                            Height =240
                            Name ="Label1"
                            Caption ="مسير"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    ReadingOrder =2
                    Left =1247
                    Top =767
                    ColumnWidth =495
                    ColumnOrder =4
                    TabIndex =1
                    Name ="RACTIVE"
                    ControlSource ="RACTIVE"
                    DefaultValue ="True"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =1477
                            Top =737
                            Width =690
                            Height =240
                            Name ="Label17"
                            Caption ="فعال"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =1927
                    Top =2040
                    Width =5556
                    Height =317
                    ColumnWidth =8610
                    ColumnOrder =2
                    TabIndex =2
                    BackColor =13434828
                    ForeColor =8388608
                    Name ="COUST_NO"
                    ControlSource ="COUST_NO"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT hes, NAME AS nam, hes AS Expr1 FROM CUST_HESAB"
                    ColumnWidths ="0;3969"
                    OnNotInList ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =7561
                            Top =2044
                            Width =735
                            Height =405
                            ForeColor =8388608
                            Name ="Label19"
                            Caption ="نام مشتري"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    AutoExpand = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2211
                    Top =3688
                    Width =1926
                    Height =317
                    ColumnWidth =2070
                    ColumnOrder =3
                    TabIndex =3
                    BackColor =13434828
                    ForeColor =8388608
                    Name ="COUST_NO2"
                    ControlSource ="COUST_NO"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT RTRIM(CAST(N_KOL AS nvarchar)) + '-' + RTRIM(CAST(NUMBER AS nvarchar)) + "
                        "'-' + RTRIM(CAST(TNUMBER AS nvarchar)) AS Expr1 FROM TDETA_HES"
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =1314
                            Top =3118
                            Width =840
                            Height =240
                            ForeColor =8388608
                            Name ="Label18"
                            Caption ="كد مشتري"
                        End
                    End
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2948
                    Top =3347
                    Width =2946
                    Height =317
                    ColumnWidth =1680
                    TabIndex =4
                    BackColor =65535
                    ForeColor =8388608
                    Name ="CLASS"
                    ControlSource ="CLASS"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT CLASS FROM Visit_route_dtl GROUP BY CLASS ORDER BY CLASS"
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =4490
                            Top =2948
                            Width =1065
                            Height =240
                            ForeColor =8388608
                            Name ="Label9"
                            Caption ="كلاس مشتري"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2948
                    Top =963
                    Width =1266
                    Height =317
                    ColumnWidth =1417
                    TabIndex =5
                    Name ="VDATE"
                    ControlSource ="VDATE"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =85
                            Left =4308
                            Top =970
                            Width =870
                            Height =240
                            Name ="Label22"
                            Caption ="تاريخ ويزيت"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2550
                    Top =1587
                    TabIndex =6
                    Name ="CDATE"
                    ControlSource ="CDATE"
                    DefaultValue ="=Now()"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5325
                            Top =1590
                            Width =810
                            Height =240
                            Name ="Label10"
                            Caption ="تاريخ ايجاد"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    ReadingOrder =2
                    Left =907
                    Top =1504
                    TabIndex =7
                    Name ="TOPLACE"
                    ControlSource ="TOPLACE"
                    OnDblClick ="[Event Procedure]"
                    DefaultValue ="True"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =1140
                            Top =1470
                            Width =1365
                            Height =240
                            Name ="Label12"
                            Caption ="ويزيت فقط حضوري"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "VISITORS_DAY_DTL_SUB.cls"
