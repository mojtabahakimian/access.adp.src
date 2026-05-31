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
    RowHeight =420
    ItemSuffix =17
    Left =645
    Top =2280
    Right =15240
    Bottom =8220
    DatasheetForeColor =33554432
    DatasheetGridlinesColor =12632256
    MaxRecords =1000000
    Filter ="((CUSTCODE=\"104\"))"
    RecSrcDt = Begin
        0x9150aa9eb297e540
    End
    RecordSource ="PRICE_ELAMIETF_DTL"
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
        Begin Subform
            SpecialEffect =2
            Width =1701
            Height =1701
        End
        Begin Section
            CanGrow = NotDefault
            Height =4195
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    IMESentenceMode =3
                    Left =1927
                    Top =113
                    ColumnWidth =555
                    ColumnOrder =0
                    Name ="PEID"
                    ControlSource ="PEID"
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
                    OverlapFlags =85
                    TextAlign =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    IMESentenceMode =3
                    Left =1927
                    Top =1133
                    ColumnWidth =1920
                    ColumnOrder =3
                    TabIndex =3
                    Name ="TF1"
                    ControlSource ="TF1"
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
                            Caption ="تخفيف اول"
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
                    ColumnWidth =3765
                    ColumnOrder =2
                    TabIndex =2
                    Name ="PPID"
                    ControlSource ="PPID"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT PPID, PPAME FROM PRICE_PAYNO"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =795
                            Width =915
                            Height =240
                            Name ="Label5"
                            Caption ="نوع پرداخت"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    IMESentenceMode =3
                    Left =1927
                    Top =1815
                    ColumnWidth =1710
                    ColumnOrder =4
                    TabIndex =4
                    Name ="TF2"
                    ControlSource ="TF2"
                    Format ="Standard"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =1822
                            Width =750
                            Height =240
                            Name ="Label11"
                            Caption ="تخفيف دوم"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    ColumnHidden = NotDefault
                    OverlapFlags =93
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    IMESentenceMode =3
                    Left =1927
                    Top =2182
                    ColumnWidth =1695
                    ColumnOrder =5
                    TabIndex =5
                    Name ="PETID"
                    ControlSource ="PETID"
                    Format ="Standard"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =2190
                            Width =825
                            Height =240
                            Name ="Label13"
                            Caption ="شماره سطر"
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
                    Top =453
                    ColumnWidth =4260
                    ColumnOrder =1
                    TabIndex =1
                    Name ="CUSTCODE"
                    ControlSource ="CUSTCODE"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT CUST_COD, CUSTKNAME FROM CUSTKIND"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =450
                            Width =855
                            Height =240
                            Name ="Label3"
                            Caption ="نوع مشتري"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =93
                    TextAlign =3
                    IMESentenceMode =3
                    Left =396
                    Top =3004
                    Width =2961
                    Height =340
                    TabIndex =6
                    BackColor =10079487
                    Name ="USERNAME"
                    ControlSource ="USERNAME"
                    DefaultValue ="UCurrentUser()"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3412
                            Top =3004
                            Width =930
                            Height =240
                            Name ="Label14"
                            Caption ="نام کاربر"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    ColumnHidden = NotDefault
                    OverlapFlags =93
                    TextAlign =3
                    IMESentenceMode =3
                    Left =396
                    Top =3464
                    Width =2961
                    Height =340
                    ColumnWidth =2070
                    TabIndex =7
                    BackColor =10079487
                    Name ="TR_DATE"
                    ControlSource ="TR_DATE"
                    DefaultValue ="Now()"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3412
                            Top =3464
                            Width =930
                            Height =240
                            Name ="Label21"
                            Caption ="تاريخ ايجاد"
                        End
                    End
                End
                Begin Subform
                    OverlapFlags =247
                    Left =510
                    Top =2494
                    TabIndex =8
                    Name ="PRICE_ELAMIETF_EXCEPTION_FRM"
                    SourceObject ="Form.PRICE_ELAMIETF_EXCEPTION_FRM"
                    LinkChildFields ="PETID"
                    LinkMasterFields ="PETID"
                    Begin
                        Begin Label
                            OverlapFlags =255
                            Left =1551
                            Top =2254
                            Width =660
                            Height =240
                            Name ="Label16"
                            Caption ="Child15:"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "PRICE_ELAMIETF_DTL_FORM.cls"
