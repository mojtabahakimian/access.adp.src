Version =20
VersionRequired =20
Begin Form
    DefaultView =2
    AllowUpdating =4
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =4065
    RowHeight =390
    ItemSuffix =37
    Left =240
    Top =1320
    Right =14880
    Bottom =8565
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xdc291d3cd9c0e240
    End
    RecordSource ="PRGANBAR"
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
        Begin FormHeader
            Height =0
            BackColor =-2147483633
            Name ="FormHeader"
        End
        Begin Section
            Height =3641
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnOrder =0
                    Name ="PRGID"
                    ControlSource ="PRGID"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3032
                            Width =540
                            Height =240
                            Name ="Label1"
                            Caption ="PRGID"
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
                    Top =1361
                    ColumnWidth =960
                    ColumnOrder =3
                    TabIndex =1
                    Name ="MEGH"
                    ControlSource ="MANDAH"
                    Format ="Fixed"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3120
                            Top =1365
                            Width =645
                            Height =240
                            Name ="Label9"
                            Caption ="موجودي"
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
                    Top =1701
                    ColumnWidth =1725
                    ColumnOrder =5
                    TabIndex =2
                    Name ="MEGHK"
                    ControlSource ="MOGMOG"
                    Format ="Fixed"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2430
                            Top =1695
                            Width =1515
                            Height =240
                            Name ="Label11"
                            Caption ="موجودي  قابل مصرف"
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
                    Top =2041
                    ColumnWidth =1275
                    ColumnOrder =4
                    TabIndex =3
                    Name ="PERT"
                    ControlSource ="MIN_M"
                    Format ="Fixed"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2595
                            Top =2040
                            Width =1125
                            Height =240
                            Name ="Label13"
                            Caption ="حداقل موجودي"
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
                    Top =2381
                    ColumnWidth =1110
                    ColumnOrder =6
                    TabIndex =4
                    Name ="KOLMAV"
                    ControlSource ="kolm"
                    Format ="Fixed"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2970
                            Top =2385
                            Width =900
                            Height =240
                            Name ="Label15"
                            Caption ="مواد مود نياز"
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
                    Top =2670
                    ColumnWidth =1305
                    ColumnOrder =7
                    TabIndex =5
                    Name ="SEF"
                    ControlSource ="=IIf([kolm]-[MOGMOG]>0,[kolm]-[MOGMOG],0)"
                    Format ="Fixed"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3195
                            Top =2664
                            Width =870
                            Height =240
                            Name ="Label17"
                            Caption ="كسري مواد"
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
                    Left =56
                    Top =3061
                    ColumnWidth =960
                    ColumnOrder =8
                    TabIndex =6
                    Name ="EZAF"
                    ControlSource ="=IIf([MOGMOG]-[kolm]>0,[MOGMOG]-[kolm],0)"
                    Format ="Fixed"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3015
                            Top =3060
                            Width =795
                            Height =240
                            Name ="Label19"
                            Caption ="مواد اضافه"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Top =963
                    ColumnWidth =1155
                    ColumnOrder =2
                    TabIndex =8
                    Name ="VAHED"
                    ControlSource ="VAHED"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TCOD_VAHEDS.CODE, TCOD_VAHEDS.NAMES FROM TCOD_VAHEDS;"
                    ColumnWidths ="0"
                    Format ="General Number"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3182
                            Top =963
                            Width =390
                            Height =240
                            Name ="Label7"
                            Caption ="واحد"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =21
                    Top =680
                    Width =1680
                    ColumnWidth =5385
                    ColumnOrder =1
                    TabIndex =7
                    Name ="CODB"
                    ControlSource ="NAME"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2087
                            Top =680
                            Width =1485
                            Height =240
                            Name ="Label5"
                            Caption ="مواد مصرفي"
                        End
                    End
                End
            End
        End
        Begin FormFooter
            Height =1134
            BackColor =-2147483633
            Name ="FormFooter"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Name ="SKOLMAV"
                    ControlSource ="=Sum([SEF])"
                    Format ="General Number"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2972
                            Width =600
                            Height =240
                            Name ="Label34"
                            Caption ="كل مواد"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "QPROGPAS3_SUB.cls"
