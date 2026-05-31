Version =20
VersionRequired =20
Begin Form
    DefaultView =2
    AllowUpdating =4
    TabularCharSet =178
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =4940
    RowHeight =420
    ItemSuffix =19
    Left =2205
    Top =3645
    Right =16770
    Bottom =9585
    DatasheetForeColor =33554432
    DatasheetGridlinesColor =12632256
    MaxRecords =1000000
    RecSrcDt = Begin
        0xbd5e11cdce99e540
    End
    RecordSource ="PRICE_ELAMIE_DTL"
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
            Height =3804
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
                    Name ="PEPID"
                    ControlSource ="PEPID"
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
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    IMESentenceMode =3
                    Left =1927
                    Top =1133
                    ColumnWidth =1920
                    ColumnOrder =3
                    TabIndex =1
                    Name ="PRICE1"
                    ControlSource ="PRICE1"
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
                            Caption ="قيمت"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    IMESentenceMode =3
                    Left =1927
                    Top =2182
                    ColumnWidth =1695
                    ColumnOrder =4
                    TabIndex =2
                    Name ="PETID"
                    ControlSource ="PERID"
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
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =396
                    Top =3004
                    Width =2961
                    Height =340
                    ColumnWidth =1800
                    ColumnOrder =5
                    TabIndex =3
                    BackColor =13408767
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
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =396
                    Top =3464
                    Width =2961
                    Height =340
                    ColumnWidth =2820
                    ColumnOrder =6
                    TabIndex =4
                    BackColor =13408767
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
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =1927
                    Top =480
                    ColumnWidth =9135
                    ColumnOrder =1
                    TabIndex =5
                    Name ="PGID"
                    ControlSource ="PGID"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT PGID, PGNAME FROM PRICE_GRP"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =480
                            Width =885
                            Height =240
                            Name ="Label16"
                            Caption ="گروه قيمتي"
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
                    Top =840
                    ColumnWidth =1200
                    ColumnOrder =2
                    TabIndex =6
                    Name ="PGID2"
                    ControlSource ="PGID"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =840
                            Width =1095
                            Height =240
                            Name ="Label18"
                            Caption ="کد گروه قيمتي"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "PRICE_ELAMIE_DTL_FORM.cls"
