Version =20
VersionRequired =20
Begin Form
    OrderByOn = NotDefault
    DefaultView =2
    AllowUpdating =4
    ViewsAllowed =2
    TabularCharSet =178
    TabularFamily =10
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =4762
    ItemSuffix =23
    Left =600
    Top =720
    Right =7140
    Bottom =6555
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x694ac5893067e240
    End
    RecordSource ="KART_KALA"
    Caption ="كارت  انبار كالا"
    DatasheetFontName ="Arabic Transparent"
    OnError ="[Event Procedure]"
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
            Height =5228
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =933
                    Top =113
                    ColumnWidth =480
                    ColumnOrder =4
                    Name ="ANBAR"
                    ControlSource ="ANBAR"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3920
                            Top =113
                            Width =585
                            Height =240
                            Name ="Label1"
                            Caption ="انبار"
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
                    Top =453
                    Width =2490
                    ColumnWidth =615
                    ColumnOrder =5
                    TabIndex =1
                    Name ="CODE"
                    ControlSource ="CODE"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4005
                            Top =450
                            Width =510
                            Height =240
                            Name ="Label3"
                            Caption ="كد كالا"
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
                    Top =793
                    Width =2490
                    ColumnOrder =6
                    TabIndex =2
                    Name ="DATE_NN"
                    ControlSource ="=Format([DATE_N],\"yyyy/mm/dd\")"
                    Format ="yyyy/mm/dd"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3845
                            Top =793
                            Width =660
                            Height =240
                            Name ="Label5"
                            Caption ="تاريخ"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =933
                    Top =1474
                    ColumnOrder =8
                    TabIndex =3
                    Name ="MABL_A"
                    ControlSource ="MABL_A"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3845
                            Top =1474
                            Width =660
                            Height =240
                            Name ="Label9"
                            Caption ="مبلغ "
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =933
                    Top =1814
                    ColumnOrder =9
                    TabIndex =4
                    Name ="MABLK"
                    ControlSource ="MABLK"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3945
                            Top =1815
                            Width =690
                            Height =240
                            Name ="Label11"
                            Caption ="مبلغ  كل"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =933
                    Top =2494
                    ColumnOrder =7
                    TabIndex =5
                    Name ="MEGK"
                    ControlSource ="MEGK"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4020
                            Top =2490
                            Width =690
                            Height =240
                            Name ="Label15"
                            Caption ="مقدار كل"
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
                    Top =2834
                    Width =2490
                    ColumnOrder =0
                    TabIndex =6
                    Name ="BARGAH"
                    ControlSource ="BARGAH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3815
                            Top =2834
                            Width =690
                            Height =240
                            Name ="Label17"
                            Caption ="نوع برگه"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =850
                    Top =3344
                    ColumnWidth =975
                    ColumnOrder =1
                    TabIndex =7
                    Name ="NUMBER"
                    ControlSource ="NUMBER"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3495
                            Top =3345
                            Width =855
                            Height =240
                            Name ="Label18"
                            Caption ="شماره برگه"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1303
                    Top =4081
                    ColumnOrder =2
                    TabIndex =8
                    Name ="NAME"
                    ControlSource ="NAME"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4150
                            Top =4081
                            Width =555
                            Height =240
                            Name ="Label19"
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
                    Left =1190
                    Top =4535
                    ColumnWidth =1635
                    ColumnOrder =3
                    TabIndex =9
                    Name ="NAMES"
                    ControlSource ="NAMES"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3962
                            Top =4535
                            Width =630
                            Height =240
                            Name ="Label20"
                            Caption ="نام انبار"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =453
                    Top =1133
                    TabIndex =10
                    Name ="DATE_N"
                    ControlSource ="DATE_N"
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =85
                            Left =3210
                            Top =1133
                            Width =645
                            Height =240
                            Name ="Label22"
                            Caption ="تاريخ"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "KART_KALA.cls"
