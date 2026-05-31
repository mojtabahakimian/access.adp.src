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
    RowHeight =450
    ItemSuffix =37
    Left =210
    Top =1320
    Right =14880
    Bottom =8565
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xd798af13522ae340
    End
    RecordSource ="SELECT QPROGPAS2.PRGID, QPROGPAS2.CODB, QPROGPAS2.VAHED, QPROGPAS2.SumOfMEGH AS "
        "MEGH, QPROGPAS2.SumOfMEGHK AS MEGHK, QPROGPAS2.SumOfPERT AS PERT, QPROGPAS2.SumO"
        "fKOLMAV AS KOLMAV, QPROGPAS2.SumOfKOLMAV * mabstandard.price AS MABLK, QPROGPAS2"
        ".PASED, QPROGPAS2.RADAH, mabstandard.price AS mabl FROM QPROGPAS2 INNER JOIN mab"
        "standard ON QPROGPAS2.CODB = mabstandard.CODE"
    DatasheetFontName ="Tahoma"
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
        Begin FormHeader
            Height =0
            BackColor =-2147483633
            Name ="FormHeader"
        End
        Begin Section
            Height =2777
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
                    ColumnWidth =1417
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
                    ColumnHidden = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Top =885
                    ColumnWidth =870
                    ColumnOrder =2
                    TabIndex =1
                    Name ="MEGH1"
                    ControlSource ="MEGH"
                    Format ="Fixed"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3120
                            Top =889
                            Width =645
                            Height =240
                            Name ="Label9"
                            Caption ="مقدار"
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
                    Top =1191
                    ColumnWidth =1230
                    ColumnOrder =4
                    TabIndex =2
                    Name ="MEGHK1"
                    ControlSource ="MEGHK"
                    Format ="Fixed"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2430
                            Top =1185
                            Width =1515
                            Height =240
                            Name ="Label11"
                            Caption ="مقدار كل"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    DecimalPlaces =2
                    OverlapFlags =93
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =56
                    Top =1531
                    ColumnWidth =630
                    ColumnOrder =3
                    TabIndex =3
                    Name ="PERT1"
                    ControlSource ="PERT"
                    Format ="Fixed"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2651
                            Top =1530
                            Width =1125
                            Height =240
                            Name ="Label13"
                            Caption ="پرت"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    DecimalPlaces =2
                    OverlapFlags =87
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Top =1785
                    ColumnWidth =1110
                    ColumnOrder =5
                    TabIndex =4
                    Name ="KOLMAV1"
                    ControlSource ="KOLMAV"
                    Format ="Fixed"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2970
                            Top =1789
                            Width =900
                            Height =240
                            Name ="Label15"
                            Caption ="كل مواد"
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
                    Top =2076
                    ColumnWidth =1260
                    ColumnOrder =7
                    TabIndex =5
                    Name ="MABL1"
                    ControlSource ="MABL"
                    Format ="#,###"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3195
                            Top =2070
                            Width =870
                            Height =240
                            Name ="Label17"
                            Caption ="في"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Top =2371
                    ColumnWidth =2010
                    ColumnOrder =8
                    TabIndex =6
                    Name ="MABLK1"
                    ControlSource ="MABLK"
                    Format ="#,###"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2959
                            Top =2370
                            Width =795
                            Height =240
                            Name ="Label19"
                            Caption ="مبلغ كل"
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
                    Top =583
                    ColumnWidth =1170
                    ColumnOrder =6
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
                            Top =583
                            Width =390
                            Height =240
                            Name ="Label7"
                            Caption ="واحد"
                        End
                    End
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =21
                    Top =300
                    Width =1680
                    ColumnWidth =6855
                    ColumnOrder =1
                    TabIndex =7
                    Name ="CODB"
                    ControlSource ="CODB"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT STUF_DEF.CODE, STUF_DEF.NAME FROM STUF_DEF;"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2087
                            Top =300
                            Width =1485
                            Height =240
                            Name ="Label5"
                            Caption ="كالا"
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
                    ControlSource ="=Sum([KOLMAV])"
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
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Top =510
                    TabIndex =1
                    Name ="SMABLK"
                    ControlSource ="=Sum([MABLK])"
                    Format ="General Number"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2957
                            Top =510
                            Width =615
                            Height =240
                            Name ="Label36"
                            Caption ="مبلغ كل"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "QPROGPAS2_SUB.cls"
