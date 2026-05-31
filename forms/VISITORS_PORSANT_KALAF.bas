Version =20
VersionRequired =20
Begin Form
    DefaultView =2
    AllowUpdating =4
    ViewsAllowed =2
    TabularCharSet =178
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =4940
    RowHeight =330
    ItemSuffix =8
    Left =465
    Top =2220
    Right =14880
    Bottom =9090
    DatasheetForeColor =33554432
    DatasheetGridlinesColor =12632256
    MaxRecords =1000000
    RecSrcDt = Begin
        0x15f2cc54f391e440
    End
    RecordSource ="VISITORS_PORSANT_KALA"
    Caption ="ليست پورسانت كالاها"
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
            Height =1615
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =1
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1927
                    Top =453
                    ColumnWidth =765
                    ColumnOrder =2
                    TabIndex =1
                    Name ="PORID"
                    ControlSource ="PORID"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =450
                            Width =1230
                            Height =240
                            Name ="Label3"
                            Caption ="شناسه پورسانت"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =1
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1927
                    Top =793
                    ColumnOrder =3
                    TabIndex =2
                    Name ="PORSANT"
                    ControlSource ="PORSANT"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =793
                            Width =780
                            Height =240
                            Name ="Label5"
                            Caption ="پورسانت"
                        End
                    End
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =1927
                    Top =113
                    Width =2490
                    Height =255
                    ColumnWidth =10770
                    ColumnOrder =1
                    Name ="CODE"
                    ControlSource ="CODE"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT CODE, NAME, CODE AS Expr1 FROM STUF_DEF WHERE (NOT (MENUIT IS NULL)) ORDE"
                        "R BY NAME"
                    ColumnWidths ="0"
                    OnNotInList ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =113
                            Width =495
                            Height =240
                            Name ="Label1"
                            Caption ="كالا"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1870
                    Top =1360
                    Width =2490
                    ColumnOrder =0
                    TabIndex =3
                    Name ="CODE1"
                    ControlSource ="CODE"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT CODE FROM STUF_DEF WHERE (NOT (MENUIT IS NULL))"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =1365
                            Width =555
                            Height =240
                            Name ="Label7"
                            Caption =" كد كالا"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "VISITORS_PORSANT_KALAF.cls"
