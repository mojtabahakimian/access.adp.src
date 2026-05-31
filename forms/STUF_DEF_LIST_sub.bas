Version =20
VersionRequired =20
Begin Form
    AllowDeletions = NotDefault
    AllowAdditions = NotDefault
    DefaultView =2
    AllowUpdating =4
    ViewsAllowed =2
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =6973
    RowHeight =375
    ItemSuffix =42
    Left =465
    Top =855
    Right =14730
    Bottom =7260
    DatasheetGridlinesColor =12632256
    MaxRecords =10000000
    RecSrcDt = Begin
        0x5448475e9aa9e340
    End
    RecordSource ="SELECT CODE, NAME, N_FANI, TOZIH, VAHED, B_SEF, N_SEF, MIN_M, MAX_M, RADAH, KIND"
        "K, MABL_F FROM STUF_DEF"
    Caption ="مشخصات كالاها"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Tahoma"
    OnError ="[Event Procedure]"
    AllowFormView =0
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
            Height =1927
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4283
                    Top =56
                    Width =1605
                    Height =255
                    ColumnOrder =0
                    FontWeight =700
                    Name ="CODE"
                    ControlSource ="CODE"
                    OnDblClick ="[Event Procedure]"
                    FontName ="Arabic Transparent"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =6051
                            Top =63
                            Width =600
                            Height =300
                            Name ="Label1"
                            Caption ="كد كالا"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4373
                    Top =736
                    Width =1515
                    Height =315
                    ColumnWidth =1890
                    ColumnOrder =1
                    FontWeight =700
                    TabIndex =1
                    Name ="N_FANI"
                    ControlSource ="N_FANI"
                    FontName ="Arabic Transparent"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =6051
                            Top =736
                            Width =915
                            Height =300
                            Name ="Label5"
                            Caption ="شماره فني"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3803
                    Top =1076
                    Width =2085
                    Height =315
                    ColumnWidth =1620
                    ColumnOrder =6
                    FontWeight =700
                    TabIndex =2
                    Name ="TOZIH"
                    ControlSource ="TOZIH"
                    FontName ="Arabic Transparent"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =6051
                            Top =1076
                            Width =615
                            Height =300
                            Name ="Label7"
                            Caption ="توضيح"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =683
                    Top =56
                    Width =1296
                    Height =315
                    ColumnWidth =1417
                    ColumnOrder =7
                    FontWeight =700
                    TabIndex =3
                    Name ="B_SEF"
                    ControlSource ="B_SEF"
                    FontName ="Arabic Transparent"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =2142
                            Top =56
                            Width =1080
                            Height =300
                            Name ="Label11"
                            Caption ="بهينه سفارش"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =684
                    Top =396
                    Width =1296
                    Height =315
                    ColumnWidth =1417
                    ColumnOrder =8
                    FontWeight =700
                    TabIndex =4
                    Name ="N_SEF"
                    ControlSource ="N_SEF"
                    FontName ="Arabic Transparent"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =2143
                            Top =396
                            Width =1080
                            Height =300
                            Name ="Label13"
                            Caption ="نقطه سفارش"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =683
                    Top =1088
                    Width =1296
                    Height =315
                    ColumnWidth =1417
                    ColumnOrder =9
                    FontWeight =700
                    TabIndex =5
                    Name ="MIN_M"
                    ControlSource ="MIN_M"
                    FontName ="Arabic Transparent"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =2142
                            Top =1088
                            Width =1260
                            Height =300
                            Name ="Label15"
                            Caption ="حداقل موجودي"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =684
                    Top =736
                    Width =1296
                    Height =315
                    ColumnWidth =1417
                    ColumnOrder =10
                    FontWeight =700
                    TabIndex =6
                    Name ="MAX_M"
                    ControlSource ="MAX_M"
                    FontName ="Arabic Transparent"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =2143
                            Top =736
                            Width =1320
                            Height =300
                            Name ="Label17"
                            Caption ="حد اكثر موجودي"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =680
                    Top =1474
                    Width =1296
                    Height =315
                    ColumnWidth =1995
                    ColumnOrder =3
                    FontWeight =700
                    TabIndex =7
                    Name ="MABL_F"
                    ControlSource ="MABL_F"
                    Format ="Standard"
                    OnDblClick ="[Event Procedure]"
                    FontName ="Arabic Transparent"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =2139
                            Top =1474
                            Width =1380
                            Height =300
                            Name ="Label23"
                            Caption ="في عمده فروش "
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3803
                    Top =396
                    Width =2085
                    Height =315
                    ColumnWidth =5685
                    ColumnOrder =2
                    FontWeight =700
                    TabIndex =8
                    Name ="NAME"
                    ControlSource ="NAME"
                    OnDblClick ="[Event Procedure]"
                    FontName ="Arabic Transparent"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =6051
                            Top =396
                            Width =660
                            Height =300
                            Name ="Label3"
                            Caption ="نام كالا"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "STUF_DEF_LIST_sub.cls"
