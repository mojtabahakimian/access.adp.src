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
    Width =4940
    RowHeight =330
    ItemSuffix =10
    Left =465
    Top =1080
    Right =6360
    Bottom =7650
    DatasheetForeColor =16711680
    DatasheetGridlinesColor =12632256
    MaxRecords =100000
    RecSrcDt = Begin
        0xbef51289231ce340
    End
    RecordSource ="SURAT_MALI"
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
            Height =2130
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
                    ColumnWidth =1417
                    Name ="CODG"
                    ControlSource ="CODG"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =120
                            Width =615
                            Height =240
                            Name ="Label1"
                            Caption ="كد رديف"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1927
                    Top =453
                    Width =2490
                    TabIndex =1
                    Name ="HESAB"
                    ControlSource ="HESAB"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =453
                            Width =555
                            Height =240
                            Name ="Label3"
                            Caption ="حساب"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =283
                    Top =1303
                    TabIndex =3
                    Name ="prg"
                    ControlSource ="prg"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3315
                            Top =1305
                            Width =525
                            Height =240
                            Name ="Label9"
                            Caption ="شماره"
                        End
                    End
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =1927
                    Top =793
                    Width =2490
                    Height =345
                    ColumnWidth =4095
                    TabIndex =2
                    Name ="TOSIF"
                    ControlSource ="TOSIF"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT IDD, TITEL FROM SQLSTATE ORDER BY TITEL"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =795
                            Width =1605
                            Height =240
                            Name ="Label5"
                            Caption ="نقل از گزارش پيشرفته"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "SURAT_MALI_SUB.cls"
