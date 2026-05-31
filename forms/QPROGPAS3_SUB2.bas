Version =20
VersionRequired =20
Begin Form
    DefaultView =2
    AllowUpdating =4
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =4633
    RowHeight =390
    ItemSuffix =40
    Left =450
    Top =1320
    Right =14880
    Bottom =8550
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x45df7bdb43e6e540
    End
    RecordSource ="SELECT QPROGPASS2.CODB, QPROGPASS2.VAHED, QPROGPASS2.SumOfKOLMAV AS kolm, QPROGP"
        "ASS2.GRP, QPROGPASS2.SumOfMABLK / QPROGPASS2.SumOfKOLMAV AS mab, QPROGPASS2.SumO"
        "fMABLK AS mablk, STUF_DEF.NAME, STUF_DEF.RADAH FROM STUF_DEF INNER JOIN QPROGPAS"
        "S2 ON STUF_DEF.CODE = QPROGPASS2.CODB WHERE (QPROGPASS2.SumOfKOLMAV <> 0)"
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
        Begin Subform
            SpecialEffect =2
            Width =1701
            Height =1701
        End
        Begin FormHeader
            Height =0
            BackColor =-2147483633
            Name ="FormHeader"
        End
        Begin Section
            Height =4422
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
                    Left =568
                    ColumnOrder =0
                    Name ="GRP"
                    ControlSource ="GRP"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3600
                            Width =540
                            Height =240
                            Name ="Label1"
                            Caption ="PRGID"
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
                    Left =568
                    Top =2381
                    ColumnWidth =1470
                    ColumnOrder =3
                    TabIndex =2
                    Name ="KOLMAV"
                    ControlSource ="kolm"
                    Format ="Fixed"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3538
                            Top =2385
                            Width =900
                            Height =240
                            Name ="Label15"
                            Caption ="مواد مود نياز"
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
                    Left =568
                    Top =963
                    ColumnWidth =1230
                    ColumnOrder =2
                    TabIndex =3
                    Name ="VAHED"
                    ControlSource ="VAHED"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TCOD_VAHEDS.CODE, TCOD_VAHEDS.NAMES FROM TCOD_VAHEDS;"
                    ColumnWidths ="0"
                    Format ="General Number"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3750
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
                    IMESentenceMode =3
                    Left =566
                    Top =3458
                    ColumnWidth =1530
                    ColumnOrder =4
                    TabIndex =4
                    Name ="mab"
                    ControlSource ="mab"
                    Format ="#,###"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3533
                            Top =3458
                            Width =435
                            Height =240
                            Name ="Label37"
                            Caption ="في"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =566
                    Top =3798
                    ColumnWidth =2190
                    ColumnOrder =5
                    TabIndex =5
                    Name ="mablk"
                    ControlSource ="mablk"
                    Format ="#,###"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3428
                            Top =3798
                            Width =540
                            Height =240
                            Name ="Label38"
                            Caption ="مبلغ"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =566
                    Top =1474
                    ColumnWidth =5670
                    ColumnOrder =1
                    TabIndex =6
                    Name ="NAME"
                    ControlSource ="NAME"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3413
                            Top =1474
                            Width =555
                            Height =240
                            Name ="Label39"
                            Caption ="نام"
                        End
                    End
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    DecimalPlaces =0
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =568
                    Top =2041
                    ColumnWidth =1785
                    ColumnOrder =6
                    TabIndex =1
                    Name ="RADAH"
                    ControlSource ="RADAH"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT CODE, NAMES FROM TCOD_STUFGROUP"
                    ColumnWidths ="0"
                    Format ="Fixed"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3163
                            Top =2040
                            Width =1125
                            Height =240
                            Name ="Label13"
                            Caption ="گروه"
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
                    Left =568
                    Name ="SKOLMAV"
                    ControlSource ="=Sum([SEF])"
                    Format ="General Number"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3540
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
' See "QPROGPAS3_SUB2.cls"
