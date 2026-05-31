Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    DividingLines = NotDefault
    DefaultView =2
    AllowUpdating =4
    ViewsAllowed =2
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =5722
    RowHeight =300
    ItemSuffix =22
    Left =600
    Top =585
    Right =10920
    Bottom =5760
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x3ca7f710d20de340
    End
    RecordSource ="TARAZ_ANBAR_KOL"
    Caption ="تراز موجودي انبار ها"
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
            Height =3780
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1854
                    Top =113
                    Width =1740
                    ColumnWidth =1365
                    ColumnOrder =0
                    Name ="CODE"
                    ControlSource ="CODE"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4970
                            Top =113
                            Width =495
                            Height =240
                            Name ="Label1"
                            Caption ="كد كالا"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =3
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1863
                    Top =453
                    Width =1731
                    ColumnWidth =1305
                    ColumnOrder =2
                    TabIndex =1
                    Name ="MEG"
                    ControlSource ="MEG"
                    Format ="Fixed"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5075
                            Top =453
                            Width =390
                            Height =240
                            Name ="Label3"
                            Caption ="مقدار اوليه"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1863
                    Top =793
                    Width =1731
                    ColumnWidth =1455
                    ColumnOrder =3
                    TabIndex =2
                    Name ="SumOfMABL_A"
                    ControlSource ="SumOfMABL_A"
                    Format ="#,###"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4325
                            Top =793
                            Width =1140
                            Height =240
                            Name ="Label5"
                            Caption ="مبلغ اوليه"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =3
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1863
                    Top =1133
                    Width =1731
                    ColumnWidth =1155
                    ColumnOrder =4
                    TabIndex =3
                    Name ="MEGHKH"
                    ControlSource ="MEGHKH"
                    Format ="Fixed"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4325
                            Top =1133
                            Width =1140
                            Height =240
                            Name ="Label7"
                            Caption ="مقدار افزايش"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1863
                    Top =1474
                    Width =1731
                    ColumnWidth =1350
                    ColumnOrder =5
                    TabIndex =4
                    Name ="MABKH"
                    ControlSource ="MABKH"
                    Format ="#,###"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4430
                            Top =1474
                            Width =1035
                            Height =240
                            Name ="Label9"
                            Caption ="مبلغ افزايش"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =3
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1863
                    Top =1814
                    Width =1731
                    ColumnWidth =1275
                    ColumnOrder =6
                    TabIndex =5
                    Name ="MEGFR"
                    ControlSource ="MEGFR"
                    Format ="Fixed"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4430
                            Top =1814
                            Width =1035
                            Height =240
                            Name ="Label11"
                            Caption ="مقدار كاهش"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1863
                    Top =2154
                    Width =1731
                    ColumnWidth =1560
                    ColumnOrder =7
                    TabIndex =6
                    Name ="MABFR"
                    ControlSource ="MABFR"
                    Format ="#,###"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4430
                            Top =2154
                            Width =1035
                            Height =240
                            Name ="Label13"
                            Caption ="مبلغ كاهش"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =3
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1863
                    Top =2494
                    Width =1731
                    ColumnWidth =1230
                    ColumnOrder =8
                    TabIndex =7
                    Name ="MEGMA"
                    ControlSource ="MEGMA"
                    Format ="Fixed"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4400
                            Top =2494
                            Width =1065
                            Height =240
                            Name ="Label15"
                            Caption ="مقدار موجودي"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1863
                    Top =2834
                    Width =1731
                    ColumnWidth =1560
                    ColumnOrder =9
                    TabIndex =8
                    Name ="MABMA"
                    ControlSource ="MABMA"
                    Format ="#,###"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4400
                            Top =2834
                            Width =1065
                            Height =240
                            Name ="Label17"
                            Caption ="مبلغ موجودي"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1854
                    Top =3174
                    Width =1740
                    ColumnWidth =3570
                    ColumnOrder =1
                    TabIndex =9
                    Name ="NAME"
                    ControlSource ="NAME"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4520
                            Top =3174
                            Width =945
                            Height =240
                            Name ="Label19"
                            Caption ="نام كالا"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1860
                    Top =3465
                    Width =1740
                    ColumnWidth =945
                    TabIndex =10
                    Name ="Text20"
                    ControlSource ="VCOD"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4976
                            Top =3465
                            Width =495
                            Height =240
                            Name ="Label21"
                            Caption ="كد كالا"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "TARAZ_ANBAR_KOL.cls"
