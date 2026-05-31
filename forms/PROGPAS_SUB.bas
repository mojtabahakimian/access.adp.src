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
    Width =5312
    RowHeight =480
    ItemSuffix =32
    Left =540
    Top =1950
    Right =14955
    Bottom =8310
    DatasheetForeColor =33554432
    DatasheetGridlinesColor =12632256
    MaxRecords =100000
    RecSrcDt = Begin
        0xc39e764846e6e540
    End
    RecordSource ="SELECT PROGPAS1.PRGID, PROGPAS1.CODA, PROGPAS1.CODB, PROGPAS1.VAHED, PROGPAS1.ME"
        "GH * PRGHEAD.NUM AS PMEGH, PROGPAS1.MEGHK * PRGHEAD.NUM AS PMEGHK, PROGPAS1.PERT"
        " * PRGHEAD.NUM AS PPER, PROGPAS1.KOLMAV * PRGHEAD.NUM AS PKOLMAV, PROGPAS1.MABL "
        "AS PMABL, PROGPAS1.MABLK * PRGHEAD.NUM AS PMABLK, PROGPAS1.PASED, PROGPAS1.RADAH"
        ", PROGPAS1.FR, PROGPAS1.MEGH, PROGPAS1.MEGHK, PROGPAS1.KOLMAV FROM PROGPAS1 LEFT"
        " OUTER JOIN PRGHEAD ON PROGPAS1.PRGID = PRGHEAD.PRGID"
    Caption ="PROGPAS1"
    OnCurrent ="[Event Procedure]"
    DatasheetFontName ="Tahoma"
    OnLoad ="[Event Procedure]"
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
            Height =5910
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
                    Left =144
                    Top =113
                    ColumnWidth =1417
                    ColumnOrder =0
                    Name ="PRGID"
                    ControlSource ="PRGID"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3176
                            Top =113
                            Width =540
                            Height =240
                            Name ="Label1"
                            Caption ="PRGID"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =144
                    Top =1474
                    ColumnWidth =750
                    ColumnOrder =6
                    TabIndex =4
                    Name ="MEGH"
                    ControlSource ="PMEGH"
                    Format ="General Number"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =3266
                            Top =1474
                            Width =450
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
                    Left =144
                    Top =1814
                    ColumnWidth =960
                    ColumnOrder =7
                    TabIndex =5
                    Name ="MEGHK"
                    ControlSource ="PMEGHK"
                    Format ="General Number"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3026
                            Top =1814
                            Width =690
                            Height =240
                            Name ="Label11"
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
                    Top =2154
                    ColumnWidth =435
                    ColumnOrder =8
                    TabIndex =6
                    Name ="PERT"
                    ControlSource ="PPER"
                    Format ="General Number"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3386
                            Top =2154
                            Width =330
                            Height =240
                            Name ="Label13"
                            Caption ="پرت"
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
                    Left =144
                    Top =2494
                    ColumnWidth =1170
                    ColumnOrder =9
                    TabIndex =7
                    Name ="KOLMAV"
                    ControlSource ="PKOLMAV"
                    Format ="Fixed"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3116
                            Top =2494
                            Width =600
                            Height =240
                            Name ="Label15"
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
                    Left =144
                    Top =2834
                    ColumnWidth =930
                    ColumnOrder =10
                    TabIndex =8
                    Name ="MABL"
                    ControlSource ="PMABL"
                    Format ="#,###"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3341
                            Top =2834
                            Width =375
                            Height =240
                            Name ="Label17"
                            Caption ="مبلغ"
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
                    Top =3174
                    ColumnWidth =1350
                    ColumnOrder =11
                    TabIndex =9
                    Name ="MABLK"
                    ControlSource ="PMABLK"
                    Format ="#,###"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3101
                            Top =3174
                            Width =615
                            Height =240
                            Name ="Label19"
                            Caption ="مبلغ كل"
                        End
                    End
                End
                Begin CheckBox
                    ColumnHidden = NotDefault
                    OverlapFlags =255
                    ReadingOrder =1
                    Left =1587
                    Top =1360
                    ColumnWidth =1417
                    ColumnOrder =1
                    TabIndex =11
                    Name ="PASED"
                    ControlSource ="PASED"
                    Begin
                        Begin Label
                            OverlapFlags =223
                            Left =3163
                            Top =1360
                            Width =555
                            Height =240
                            Name ="Label23"
                            Caption ="PASED"
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
                    Left =165
                    Top =453
                    Width =1680
                    ColumnWidth =2985
                    ColumnOrder =2
                    TabIndex =1
                    Name ="CODA"
                    ControlSource ="CODA"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT STUF_DEF.CODE, STUF_DEF.NAME FROM STUF_DEF;"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2310
                            Top =450
                            Width =1260
                            Height =210
                            Name ="Label3"
                            Caption ="كالاي ساخته "
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
                    Left =165
                    Top =793
                    Width =1680
                    ColumnWidth =3345
                    ColumnOrder =3
                    TabIndex =2
                    Name ="CODB"
                    ControlSource ="CODB"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT STUF_DEF.CODE, STUF_DEF.NAME FROM STUF_DEF;"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3146
                            Top =793
                            Width =570
                            Height =240
                            Name ="Label5"
                            Caption ="كالاي نيمه ساخته يا مواد مصرفي"
                        End
                    End
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =247
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =144
                    Top =1133
                    ColumnWidth =990
                    ColumnOrder =5
                    TabIndex =3
                    Name ="VAHED"
                    ControlSource ="VAHED"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TCOD_VAHEDS.CODE, TCOD_VAHEDS.NAMES FROM TCOD_VAHEDS;"
                    ColumnWidths ="0"
                    Format ="General Number"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =3326
                            Top =1133
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
                    Left =144
                    Top =3514
                    ColumnWidth =1485
                    ColumnOrder =12
                    TabIndex =10
                    Name ="RADAH"
                    ControlSource ="RADAH"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TCOD_STUFGROUP.CODE, TCOD_STUFGROUP.NAMES FROM TCOD_STUFGROUP;"
                    ColumnWidths ="0"
                    Format ="General Number"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3431
                            Top =3514
                            Width =285
                            Height =240
                            Name ="Label21"
                            Caption ="رده"
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
                    ColumnCount =3
                    Top =3911
                    Width =4311
                    Height =317
                    ColumnWidth =3930
                    ColumnOrder =4
                    TabIndex =12
                    BackColor =16777164
                    Name ="FR"
                    ControlSource ="FR"
                    RowSourceType ="Table/View/StoredProc"
                    ColumnWidths ="0;3402;567"
                    ValidationRule ="Not Is Null"
                    OnEnter ="[Event Procedure]"
                    OnExit ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4367
                            Top =3911
                            Width =945
                            Height =300
                            Name ="Label57"
                            Caption ="فرمول"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1020
                    Top =4422
                    ColumnWidth =840
                    ColumnOrder =13
                    TabIndex =13
                    ForeColor =16777215
                    Name ="Text24"
                    ControlSource ="PASED"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3780
                            Top =4425
                            Width =735
                            Height =240
                            Name ="Label25"
                            Caption ="تجزيه شد"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =850
                    Top =4988
                    TabIndex =14
                    Name ="Text26"
                    ControlSource ="MEGH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3600
                            Top =4995
                            Width =735
                            Height =240
                            Name ="Label27"
                            Caption ="مقدار جزء"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =850
                    Top =5329
                    TabIndex =15
                    Name ="Text28"
                    ControlSource ="MEGHK"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3600
                            Top =5325
                            Width =1080
                            Height =240
                            Name ="Label29"
                            Caption ="جمع مقدار جزء"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1190
                    Top =5669
                    TabIndex =16
                    Name ="Text30"
                    ControlSource ="KOLMAV"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3945
                            Top =5670
                            Width =870
                            Height =240
                            Name ="Label31"
                            Caption ="جز کل  مواد"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "PROGPAS_SUB.cls"
