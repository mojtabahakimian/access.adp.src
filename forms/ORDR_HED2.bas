Version =20
VersionRequired =20
Begin Form
    AutoCenter = NotDefault
    AllowDeletions = NotDefault
    AllowAdditions = NotDefault
    AllowEdits = NotDefault
    DefaultView =2
    AllowUpdating =4
    ViewsAllowed =2
    TabularCharSet =178
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =8524
    RowHeight =315
    ItemSuffix =39
    Left =600
    Top =420
    Right =14520
    Bottom =7815
    DatasheetForeColor =33554432
    DatasheetGridlinesColor =12632256
    MaxRecords =100000
    RecSrcDt = Begin
        0xf7eb13e6afdde340
    End
    RecordSource ="SELECT ORDR_HED.DATE AS ORDR_HED_DATE, ORDR_LST.CODE, SUM(ORDR_LST.MEGHK) AS MEG"
        "HKk, SUM(ORDR_LST.MEGH) AS meghj, ORDR_LST.VAHED_K FROM ORDR_HED INNER JOIN ORDR"
        "_LST ON ORDR_HED.id = ORDR_LST.ID GROUP BY ORDR_HED.DATE, ORDR_LST.CODE, ORDR_LS"
        "T.VAHED_K"
    Caption ="گزارش سفارشات"
    DatasheetFontName ="Tahoma"
    AllowFormView =0
    Begin
        Begin Label
            BackStyle =0
            TextFontCharSet =178
            BackColor =-2147483633
            ForeColor =-2147483630
        End
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
            Width =850
            Height =850
        End
        Begin Line
            Width =1701
        End
        Begin Image
            BackStyle =0
            OldBorderStyle =0
            PictureAlignment =2
            Width =1701
            Height =1701
        End
        Begin CommandButton
            TextFontCharSet =178
            Width =1701
            Height =283
            FontSize =8
            FontWeight =400
            FontName ="MS Sans Serif"
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
            OldBorderStyle =0
            BackStyle =0
            Width =4536
            Height =2835
            LabelX =-1701
        End
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            TextFontCharSet =178
            Width =1701
            LabelX =-1701
            BackColor =-2147483643
            ForeColor =-2147483640
            AsianLineBreak =255
        End
        Begin ListBox
            SpecialEffect =2
            TextFontCharSet =178
            Width =1701
            Height =1417
            LabelX =-1701
            BackColor =-2147483643
            ForeColor =-2147483640
        End
        Begin ComboBox
            SpecialEffect =2
            TextFontCharSet =178
            Width =1701
            LabelX =-1701
            BackColor =-2147483643
            ForeColor =-2147483640
        End
        Begin Subform
            SpecialEffect =2
            Width =1701
            Height =1701
        End
        Begin UnboundObjectFrame
            SpecialEffect =2
            OldBorderStyle =1
            Width =4536
            Height =2835
        End
        Begin ToggleButton
            TextFontCharSet =178
            Width =283
            Height =283
            FontSize =8
            FontWeight =400
            FontName ="MS Sans Serif"
        End
        Begin Tab
            TextFontCharSet =178
            BackStyle =0
            Width =5103
            Height =3402
        End
        Begin FormHeader
            Height =0
            BackColor =-2147483633
            Name ="FormHeader"
        End
        Begin Section
            Height =4848
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    FELineBreak = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1814
                    Top =226
                    Width =1386
                    Height =315
                    ColumnWidth =1200
                    ColumnOrder =0
                    BackColor =16777215
                    ForeColor =0
                    Name ="ORDR_HED_DATE"
                    ControlSource ="ORDR_HED_DATE"
                    DefaultValue ="FARSIDATE(Date())"
                    FontName ="Tahoma"
                    InputMask ="####/##/##"
                    AsianLineBreak =0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =3301
                            Top =226
                            Width =540
                            Height =300
                            BackColor =16777215
                            ForeColor =0
                            Name ="Label1"
                            Caption ="تاريخ"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =3
                    Left =1186
                    Top =2551
                    Width =1680
                    Height =255
                    ColumnWidth =7335
                    ColumnOrder =1
                    TabIndex =1
                    BackColor =16777215
                    ForeColor =0
                    Name ="CODE"
                    ControlSource ="CODE"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT STUF_DEF.CODE, STUF_DEF.NAME, STUF_DEF.CODE FROM STUF_DEF ORDER BY STUF_D"
                        "EF.NAME;"
                    ColumnWidths ="0;2268;567"
                    FontName ="Tahoma"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4227
                            Top =2551
                            Width =510
                            Height =240
                            BackColor =16777215
                            ForeColor =0
                            Name ="Label9"
                            Caption =" كالا"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =1186
                    Top =4251
                    Width =1821
                    Height =255
                    ColumnWidth =1170
                    ColumnOrder =3
                    TabIndex =2
                    BackColor =16777215
                    ForeColor =0
                    Name ="VAHED_K"
                    ControlSource ="VAHED_K"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT VAHEDS.VAHED, TCOD_VAHEDS.NAMES FROM TCOD_VAHEDS INNER JOIN VAHEDS ON TCO"
                        "D_VAHEDS.CODE = VAHEDS.VAHED GROUP BY VAHEDS.VAHED, TCOD_VAHEDS.NAMES ORDER BY T"
                        "COD_VAHEDS.NAMES"
                    ColumnWidths ="0"
                    FontName ="Tahoma"
                    Tag ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3693
                            Top =4251
                            Width =615
                            Height =240
                            BackColor =16777215
                            ForeColor =0
                            Name ="Label35"
                            Caption ="واحدكالا"
                            FontName ="Tahoma"
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
                    Left =1417
                    Top =963
                    ColumnOrder =4
                    TabIndex =3
                    Name ="MEGHKk"
                    ControlSource ="MEGHKk"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4024
                            Top =963
                            Width =795
                            Height =240
                            Name ="Label37"
                            Caption ="مقدار كل"
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
                    Left =1133
                    Top =1417
                    ColumnOrder =2
                    TabIndex =4
                    Name ="meghj"
                    ControlSource ="meghj"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3995
                            Top =1417
                            Width =540
                            Height =240
                            Name ="Label38"
                            Caption ="مقدار"
                        End
                    End
                End
            End
        End
        Begin FormFooter
            Height =0
            BackColor =-2147483633
            Name ="FormFooter"
        End
    End
End
