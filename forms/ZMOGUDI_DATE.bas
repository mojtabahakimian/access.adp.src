Version =20
VersionRequired =20
Begin Form
    AutoCenter = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    OrderByOn = NotDefault
    DefaultView =2
    AllowUpdating =4
    ViewsAllowed =2
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =5159
    RowHeight =420
    ItemSuffix =17
    DatasheetGridlinesColor =12632256
    MaxRecords =100000000
    RecSrcDt = Begin
        0x1cd9bda4f214e640
    End
    Caption ="ليست موجودي کالاها با تاريخ ايجاد مانده کالا"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Tahoma"
    OnLoad ="[Event Procedure]"
    AllowFormView =0
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
            Height =3458
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1927
                    Top =113
                    Width =2475
                    ColumnOrder =0
                    Name ="CODE"
                    ControlSource ="CODE"
                    OnDblClick ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =120
                            Width =510
                            Height =240
                            Name ="Label1"
                            Caption ="کد کالا"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1927
                    Top =453
                    ColumnWidth =675
                    ColumnOrder =1
                    TabIndex =1
                    Name ="ANBAR"
                    ControlSource ="ANBAR"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =453
                            Width =585
                            Height =240
                            Name ="Label3"
                            Caption ="کد انبار"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1927
                    Top =793
                    ColumnWidth =1065
                    ColumnOrder =4
                    TabIndex =2
                    Name ="MAND"
                    ControlSource ="MAND"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =795
                            Width =645
                            Height =240
                            Name ="Label5"
                            Caption ="موجودي"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1927
                    Top =1133
                    ColumnWidth =1275
                    ColumnOrder =5
                    TabIndex =3
                    Name ="FII"
                    ControlSource ="FII"
                    Format ="#,###"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =1140
                            Width =375
                            Height =240
                            Name ="Label7"
                            Caption ="مبلغ"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1927
                    Top =1474
                    ColumnWidth =1905
                    ColumnOrder =6
                    TabIndex =4
                    Name ="MABLK"
                    ControlSource ="MABLK"
                    Format ="#,###"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =1470
                            Width =615
                            Height =240
                            Name ="Label9"
                            Caption ="مبلغ کل"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1927
                    Top =2494
                    Width =2475
                    ColumnWidth =1275
                    ColumnOrder =3
                    TabIndex =7
                    Name ="EGDATE"
                    ControlSource ="EGDATE"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =2490
                            Width =1335
                            Height =240
                            Name ="Label15"
                            Caption ="تاريخ ايجاد مانده"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1871
                    Top =2948
                    ColumnWidth =4890
                    ColumnOrder =2
                    TabIndex =8
                    Name ="NAME"
                    ControlSource ="NAME"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =170
                            Top =2948
                            Width =555
                            Height =240
                            Name ="Label16"
                            Caption ="نام کالا"
                        End
                    End
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =1927
                    Top =2154
                    ColumnWidth =1515
                    ColumnOrder =7
                    TabIndex =6
                    Name ="GRCOD"
                    ControlSource ="GRCOD"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT CODE, NAMES FROM TCOD_STUFGROUP"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =2154
                            Width =615
                            Height =240
                            Name ="Label13"
                            Caption ="گروه"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1927
                    Top =1870
                    Width =2475
                    ColumnWidth =1215
                    ColumnOrder =8
                    TabIndex =5
                    Name ="NAMES"
                    ControlSource ="NAMES"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =1870
                            Width =495
                            Height =240
                            Name ="Label11"
                            Caption ="واحد"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "ZMOGUDI_DATE.cls"
