Version =20
VersionRequired =20
Begin Form
    AutoCenter = NotDefault
    DefaultView =0
    AllowUpdating =4
    ViewsAllowed =2
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =7937
    RowHeight =285
    ItemSuffix =6
    DatasheetGridlinesColor =12632256
    MaxRecords =10000000
    Filter ="((MPCODE=284))"
    OrderBy ="MPCODE"
    RecSrcDt = Begin
        0xfa46399630ffe340
    End
    RecordSource ="SELECT TCOD_MAP.* FROM TCOD_MAP WHERE (MPP < 100)"
    Caption ="كدينگ مپ شماره فني"
    OnOpen ="[Event Procedure]"
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
            Height =1153
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =6021
                    Top =453
                    Width =1866
                    TabIndex =1
                    Name ="MPCODE"
                    ControlSource ="MPCODE"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =907
                            Top =453
                            Width =705
                            Height =240
                            Name ="Label3"
                            Caption ="كد "
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
                    Top =793
                    Width =5955
                    ColumnWidth =4740
                    TabIndex =2
                    Name ="MPNAME"
                    ControlSource ="MPNAME"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =793
                            Width =705
                            Height =240
                            Name ="Label5"
                            Caption ="نام"
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
                    Left =6021
                    Top =113
                    Width =1866
                    Name ="MPP"
                    ControlSource ="MPP"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="TCOD_MAP_GRP"
                    ColumnWidths ="0"
                    OnDblClick ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =900
                            Top =120
                            Width =585
                            Height =240
                            Name ="Label1"
                            Caption ="كد گروه"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "TCOD_MAPF.cls"
