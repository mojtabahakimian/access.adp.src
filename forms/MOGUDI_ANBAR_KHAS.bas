Version =20
VersionRequired =20
Begin Form
    OrderByOn = NotDefault
    DefaultView =2
    AllowUpdating =4
    ViewsAllowed =2
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =5730
    RowHeight =285
    ItemSuffix =18
    Left =600
    Top =720
    Right =8115
    Bottom =5970
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xeb69ae75b565e240
    End
    RecordSource ="MOGUDI_KOL_ANBARHA"
    Caption ="موجودي كلي كالاها"
    DatasheetFontName ="Tahoma"
    OnError ="[Event Procedure]"
    RecordSourceQualifier ="dbo"
    Begin
        Begin Label
            BackStyle =0
            TextFontCharSet =178
            FontName ="Tahoma"
        End
        Begin CommandButton
            TextFontCharSet =178
            Width =1701
            Height =283
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Tahoma"
        End
        Begin OptionButton
            SpecialEffect =2
            LabelX =230
            LabelY =-30
        End
        Begin CheckBox
            SpecialEffect =2
            LabelX =230
            LabelY =-30
        End
        Begin OptionGroup
            SpecialEffect =3
            Width =1701
            Height =1701
        End
        Begin BoundObjectFrame
            SpecialEffect =2
            Width =4536
            Height =2835
            LabelX =-1701
        End
        Begin TextBox
            SpecialEffect =2
            OldBorderStyle =0
            TextFontCharSet =178
            Width =1701
            LabelX =-1701
            FontName ="Tahoma"
        End
        Begin ListBox
            SpecialEffect =2
            TextFontCharSet =178
            Width =1701
            Height =1417
            LabelX =-1701
            FontName ="Tahoma"
        End
        Begin ComboBox
            SpecialEffect =2
            TextFontCharSet =178
            Width =1701
            LabelX =-1701
            FontName ="Tahoma"
        End
        Begin Section
            Height =3194
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1104
                    Top =113
                    Width =2490
                    ColumnWidth =615
                    ColumnOrder =0
                    Name ="CODE"
                    ControlSource ="CODE"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4965
                            Top =120
                            Width =510
                            Height =240
                            Name ="Label1"
                            Caption ="كد كالا"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =2
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1893
                    Top =453
                    ColumnOrder =3
                    TabIndex =1
                    Name ="MAND"
                    ControlSource ="MAND"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4950
                            Top =450
                            Width =645
                            Height =240
                            Name ="Label3"
                            Caption ="موجودي"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1893
                    Top =1474
                    ColumnWidth =1785
                    ColumnOrder =4
                    TabIndex =2
                    Name ="FI"
                    ControlSource ="FI"
                    Format ="Standard"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4115
                            Top =1474
                            Width =1350
                            Height =240
                            Name ="Label9"
                            Caption ="مبلغ ميانگين"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1893
                    Top =1814
                    ColumnWidth =1680
                    ColumnOrder =5
                    TabIndex =3
                    Name ="MABLK"
                    ControlSource ="MABLK"
                    Format ="Standard"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4185
                            Top =1815
                            Width =1545
                            Height =225
                            Name ="Label11"
                            Caption ="مبلغ موجودي"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =144
                    Top =2154
                    Width =3450
                    ColumnWidth =3645
                    ColumnOrder =1
                    TabIndex =4
                    Name ="NAME"
                    ControlSource ="NAME"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4965
                            Top =2160
                            Width =540
                            Height =240
                            Name ="Label13"
                            Caption ="نام كالا"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =804
                    Top =2494
                    Width =2790
                    ColumnWidth =1185
                    ColumnOrder =2
                    TabIndex =5
                    Name ="NAMES"
                    ControlSource ="NAMES"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4835
                            Top =2494
                            Width =630
                            Height =240
                            Name ="Label15"
                            Caption ="نام واحد"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "MOGUDI_ANBAR_KHAS.cls"
