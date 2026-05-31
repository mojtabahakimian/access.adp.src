Version =20
VersionRequired =20
Begin Form
    AllowEdits = NotDefault
    DefaultView =2
    AllowUpdating =4
    ViewsAllowed =2
    TabularCharSet =178
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =8361
    RowHeight =330
    ItemSuffix =10
    Left =1560
    Top =2775
    Right =16380
    Bottom =9930
    DatasheetForeColor =33554432
    DatasheetGridlinesColor =12632256
    MaxRecords =1000000
    RecSrcDt = Begin
        0x967350541353e440
    End
    RecordSource ="Visit_route_dtl"
    Caption ="Visit_route_dtl"
    AfterUpdate ="[Event Procedure]"
    DatasheetFontName ="Tahoma"
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
            Height =4899
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1927
                    Top =113
                    Width =2490
                    ColumnWidth =1417
                    ColumnOrder =0
                    Name ="ROUTE_NAME"
                    ControlSource ="ROUTE_NAME"
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
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1927
                    Top =1133
                    ColumnWidth =885
                    ColumnOrder =1
                    TabIndex =1
                    Name ="IDR"
                    ControlSource ="IDR"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =1133
                            Width =345
                            Height =240
                            Name ="Label7"
                            Caption ="IDR"
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
                    TabIndex =2
                    Name ="RACTIVE"
                    ControlSource ="RACTIVE"
                    OnDblClick ="[Event Procedure]"
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
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =1927
                    Top =2040
                    Width =5556
                    Height =317
                    ColumnWidth =8520
                    ColumnOrder =2
                    TabIndex =3
                    BackColor =13434828
                    ForeColor =8388608
                    Name ="COUST_NO"
                    ControlSource ="COUST_NO"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT hes, NAME AS nam, hes AS Expr1 FROM CUST_HESAB"
                    ColumnWidths ="0;3969"
                    BeforeUpdate ="[Event Procedure]"
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
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =793
                    Top =3291
                    Width =1926
                    Height =317
                    ColumnWidth =2625
                    ColumnOrder =3
                    TabIndex =4
                    BackColor =13434828
                    ForeColor =8388608
                    Name ="COUST_NO2"
                    ControlSource ="COUST_NO"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT hes  FROM CUST_HESAB"
                    BeforeUpdate ="[Event Procedure]"
                    OnNotInList ="[Event Procedure]"
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =1022
                            Top =2437
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
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3855
                    Top =3688
                    Width =1926
                    Height =317
                    ColumnWidth =1680
                    TabIndex =5
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
                            Left =4080
                            Top =2835
                            Width =1065
                            Height =240
                            ForeColor =8388608
                            Name ="Label9"
                            Caption ="كلاس مشتري"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "Visit_route_dtl_FORM.cls"
