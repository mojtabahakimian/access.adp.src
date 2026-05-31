Version =20
VersionRequired =20
Begin Form
    OrderByOn = NotDefault
    DefaultView =0
    AllowUpdating =4
    ViewsAllowed =2
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =4940
    RowHeight =315
    ItemSuffix =8
    Left =600
    Top =1320
    Right =7560
    Bottom =5010
    DatasheetGridlinesColor =12632256
    MaxRecords =10000000
    RecSrcDt = Begin
        0x49904b49c0dee340
    End
    RecordSource ="TCOD_MAP_GRP"
    Caption ="كدينگ مپ"
    OnClose ="[Event Procedure]"
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
        Begin Section
            Height =1395
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =1
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1927
                    Top =113
                    Name ="MPP"
                    ControlSource ="MPP"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =113
                            Width =375
                            Height =240
                            Name ="Label1"
                            Caption ="كد"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =1
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1927
                    Top =453
                    Width =2490
                    ColumnWidth =3360
                    TabIndex =1
                    Name ="MPNAME"
                    ControlSource ="MPNAME"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =453
                            Width =705
                            Height =240
                            Name ="Label3"
                            Caption ="نام مپ"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =1
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1927
                    Top =793
                    ColumnWidth =1395
                    TabIndex =2
                    Name ="SIZEF"
                    ControlSource ="SIZEF"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =793
                            Width =840
                            Height =240
                            Name ="Label5"
                            Caption ="سايز كد در مپ"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =1
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1927
                    Top =1153
                    ColumnWidth =1545
                    TabIndex =3
                    Name ="STARTF"
                    ControlSource ="STARTF"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =1155
                            Width =1440
                            Height =240
                            Name ="Label7"
                            Caption ="شروع كد از موقعيت"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "TCOD_MAP_GRP.cls"
