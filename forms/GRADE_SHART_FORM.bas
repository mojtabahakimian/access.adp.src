Version =20
VersionRequired =20
Begin Form
    OrderByOn = NotDefault
    DefaultView =2
    AllowUpdating =4
    TabularCharSet =178
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =4940
    RowHeight =345
    ItemSuffix =14
    Left =525
    Top =3450
    Right =15105
    Bottom =9405
    DatasheetForeColor =33554432
    DatasheetGridlinesColor =12632256
    MaxRecords =1000000
    RecSrcDt = Begin
        0xd1b413271290e540
    End
    RecordSource ="GRADE_SHART"
    Caption ="GRADE_SHART"
    BeforeUpdate ="[Event Procedure]"
    DatasheetFontName ="Tahoma"
    OnError ="[Event Procedure]"
    RecordSourceQualifier ="dbo"
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
        Begin ComboBox
            SpecialEffect =2
            TextFontCharSet =178
            Width =1701
            LabelX =-1701
            FontName ="Tahoma"
        End
        Begin Section
            Height =2437
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    IMESentenceMode =3
                    Left =1927
                    Top =113
                    ColumnWidth =555
                    ColumnOrder =0
                    Name ="GSID"
                    ControlSource ="GSID"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =120
                            Width =525
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
                    ReadingOrder =2
                    ScrollBarAlign =2
                    IMESentenceMode =3
                    Left =1927
                    Top =453
                    ColumnWidth =4260
                    ColumnOrder =1
                    TabIndex =1
                    Name ="GID"
                    ControlSource ="GID"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =450
                            Width =855
                            Height =240
                            Name ="Label3"
                            Caption ="شماره گريد"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    IMESentenceMode =3
                    Left =1927
                    Top =1133
                    ColumnWidth =1920
                    ColumnOrder =3
                    TabIndex =3
                    Name ="GSVALUE"
                    ControlSource ="GSVALUE"
                    Format ="Standard"
                    OnKeyPress ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =1140
                            Width =750
                            Height =240
                            Name ="Label7"
                            Caption ="ارزش لازم"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    IMESentenceMode =3
                    Left =1927
                    Top =1474
                    Width =2490
                    ColumnWidth =2985
                    ColumnOrder =6
                    TabIndex =4
                    Name ="TOZI"
                    ControlSource ="TOZI"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =1470
                            Width =675
                            Height =240
                            Name ="Label9"
                            Caption ="توضيحات"
                        End
                    End
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =1927
                    Top =793
                    ColumnWidth =8745
                    ColumnOrder =2
                    TabIndex =2
                    Name ="GSHARTID"
                    ControlSource ="GSHARTID"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT GSHARTID, GSHNAME FROM GRADE_SHART_FUNC"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =795
                            Width =915
                            Height =240
                            Name ="Label5"
                            Caption ="شرط"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    IMESentenceMode =3
                    Left =1927
                    Top =1815
                    ColumnOrder =4
                    TabIndex =5
                    Name ="GSVALUE2"
                    ControlSource ="GSVALUE2"
                    Format ="Standard"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =1822
                            Width =750
                            Height =240
                            Name ="Label11"
                            Caption ="ارزش دوم"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    IMESentenceMode =3
                    Left =1927
                    Top =2182
                    ColumnWidth =1695
                    ColumnOrder =5
                    TabIndex =6
                    Name ="GSVALUE3"
                    ControlSource ="GSVALUE3"
                    Format ="Standard"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =2190
                            Width =825
                            Height =240
                            Name ="Label13"
                            Caption ="ارزش سوم"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "GRADE_SHART_FORM.cls"
