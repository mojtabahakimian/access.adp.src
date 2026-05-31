Version =20
VersionRequired =20
Begin Form
    AllowAdditions = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =2
    AllowUpdating =4
    ViewsAllowed =2
    TabularCharSet =178
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =5478
    RowHeight =375
    ItemSuffix =18
    Left =8940
    Top =4860
    Right =19695
    Bottom =7365
    DatasheetForeColor =33554432
    DatasheetGridlinesColor =12632256
    MaxRecords =10000000
    RecSrcDt = Begin
        0x572e41db3c2be440
    End
    RecordSource ="SELECT OTHER_DTL_SUB.* FROM OTHER_DTL_SUB ORDER BY RADIF"
    Caption ="توزين"
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
            Height =3244
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
                    ColumnOrder =0
                    Name ="NUMBER"
                    ControlSource ="NUMBER"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =113
                            Width =690
                            Height =240
                            Name ="Label1"
                            Caption ="NUMBER"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1927
                    Top =453
                    ColumnWidth =1417
                    ColumnOrder =1
                    TabIndex =1
                    Name ="TAGG"
                    ControlSource ="TAGG"
                    DefaultValue ="2"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =453
                            Width =480
                            Height =240
                            Name ="Label3"
                            Caption ="TAGG"
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
                    Top =1133
                    ColumnWidth =1230
                    ColumnOrder =3
                    TabIndex =2
                    Name ="CAM_KHALY"
                    ControlSource ="CAM_KHALY"
                    AfterUpdate ="[Event Procedure]"
                    OnKeyPress ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =1133
                            Width =960
                            Height =240
                            Name ="Label7"
                            Caption ="وزن خالي"
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
                    Top =1474
                    ColumnWidth =1200
                    ColumnOrder =4
                    TabIndex =3
                    Name ="CAM_POOR"
                    ControlSource ="CAM_POOR"
                    AfterUpdate ="[Event Procedure]"
                    OnKeyPress ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =1474
                            Width =930
                            Height =240
                            Name ="Label9"
                            Caption ="وزن پر"
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
                    Left =1927
                    Top =1814
                    ColumnWidth =1185
                    ColumnOrder =5
                    TabIndex =4
                    Name ="MEGHk"
                    ControlSource ="MEGHk"
                    OnDblClick ="[Event Procedure]"
                    OnKeyPress ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =1814
                            Width =960
                            Height =240
                            Name ="Label11"
                            Caption ="وزن كالا"
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
                    Top =2154
                    Width =2490
                    ColumnOrder =7
                    TabIndex =5
                    Name ="TOZIH"
                    ControlSource ="TOZIH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =2160
                            Width =675
                            Height =240
                            Name ="Label13"
                            Caption ="توضيحات"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =3
                    Left =1927
                    Top =737
                    Width =1680
                    Height =255
                    ColumnWidth =3855
                    ColumnOrder =2
                    TabIndex =6
                    Name ="CODE"
                    ControlSource ="CODE"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT STUF_DEF.CODE, STUF_DEF.NAME, STUF_DEF.CODE FROM STUF_DEF ORDER BY STUF_D"
                        "EF.NAME"
                    ColumnWidths ="0;2268;567"
                    ValidationRule ="Is Not Null"
                    Tag ="2607"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =737
                            Width =510
                            Height =240
                            Name ="Label14"
                            Caption =" كالا"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1927
                    Top =2721
                    ColumnWidth =1417
                    ColumnOrder =8
                    TabIndex =7
                    Name ="RADIF"
                    ControlSource ="RADIF"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4729
                            Top =2721
                            Width =600
                            Height =240
                            Name ="Label15"
                            Caption ="رديف"
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
                    Left =1934
                    Top =2430
                    ColumnWidth =1417
                    ColumnOrder =6
                    TabIndex =8
                    Name ="VAZNH"
                    ControlSource ="VAZNH"
                    OnDblClick ="[Event Procedure]"
                    OnKeyPress ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =2430
                            Width =960
                            Height =240
                            Name ="Label17"
                            Caption ="وزن حواله"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "OTHER_DTL_SUB_SUB.cls"
