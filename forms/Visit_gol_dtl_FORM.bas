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
    Width =8361
    RowHeight =330
    ItemSuffix =25
    Left =405
    Top =1605
    Right =14985
    Bottom =9210
    DatasheetForeColor =33554432
    DatasheetGridlinesColor =12632256
    MaxRecords =1000000
    RecSrcDt = Begin
        0xa12c2d492f62e440
    End
    RecordSource ="visitgol_dtl"
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
        Begin CheckBox
            SpecialEffect =2
            LabelX =230
            LabelY =-30
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
            Height =5442
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =1
                    IMESentenceMode =3
                    Left =3458
                    Top =170
                    Width =2490
                    ColumnWidth =1417
                    ColumnOrder =0
                    Name ="MAH"
                    ControlSource ="MAH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1644
                            Top =170
                            Width =1095
                            Height =240
                            Name ="Label1"
                            Caption ="ماه"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =1
                    IMESentenceMode =3
                    Left =4024
                    Top =2211
                    ColumnOrder =1
                    TabIndex =1
                    Name ="HES"
                    ControlSource ="HES"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =7005
                            Top =2205
                            Width =510
                            Height =240
                            Name ="Label8"
                            Caption ="ويزيتور"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =1
                    IMESentenceMode =3
                    Left =4024
                    Top =3911
                    ColumnWidth =1125
                    ColumnOrder =3
                    TabIndex =2
                    Name ="MEGH"
                    ControlSource ="MEGH"
                    ValidationRule =">0"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6691
                            Top =3911
                            Width =735
                            Height =240
                            Name ="Label15"
                            Caption ="مقداركالا"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =1
                    IMESentenceMode =3
                    Left =4024
                    Top =4251
                    ColumnOrder =5
                    TabIndex =3
                    Name ="MEGHk"
                    ControlSource ="MEGHk"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6406
                            Top =4251
                            Width =1020
                            Height =240
                            Name ="Label16"
                            Caption ="هدف"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =1
                    IMESentenceMode =3
                    Left =4024
                    Top =4932
                    ColumnOrder =8
                    TabIndex =4
                    Name ="USERNAME"
                    ControlSource ="USERNAME"
                    DefaultValue ="UCurrentUser()"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6481
                            Top =4932
                            Width =945
                            Height =240
                            Name ="Label18"
                            Caption ="نام كاربر"
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
                    NumeralShapes =1
                    ColumnCount =3
                    Left =4195
                    Top =963
                    Width =1680
                    Height =255
                    ColumnWidth =8055
                    ColumnOrder =2
                    TabIndex =5
                    Name ="CODE"
                    ControlSource ="CODE"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT STUF_DEF.CODE, STUF_DEF.NAME, STUF_DEF.CODE FROM STUF_DEF ORDER BY STUF_D"
                        "EF.NAME"
                    ColumnWidths ="0;6237;567"
                    ValidationRule ="Is Not Null"
                    Tag ="1"
                    OnNotInList ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =7236
                            Top =963
                            Width =510
                            Height =240
                            Name ="Label9"
                            Caption =" كالا"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    FELineBreak = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =1
                    IMESentenceMode =3
                    Left =3968
                    Top =1417
                    Width =1680
                    Height =255
                    ColumnOrder =9
                    TabIndex =6
                    Name ="CODEh"
                    ControlSource ="CODE"
                    ValidationRule ="Is Not Null"
                    Tag ="1"
                    AsianLineBreak =0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =7009
                            Top =1417
                            Width =510
                            Height =240
                            Name ="Label54"
                            Caption ="كد كالا"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =1
                    ColumnCount =2
                    Left =3970
                    Top =3231
                    Height =255
                    ColumnOrder =4
                    TabIndex =7
                    Name ="VAHED_K"
                    ControlSource ="VAHED_K"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT VAHEDS.VAHED, TCOD_VAHEDS.NAMES FROM TCOD_VAHEDS INNER JOIN VAHEDS ON TCO"
                        "D_VAHEDS.CODE = VAHEDS.VAHED GROUP BY VAHEDS.VAHED, TCOD_VAHEDS.NAMES ORDER BY T"
                        "COD_VAHEDS.NAMES"
                    ColumnWidths ="0"
                    ValidationRule ="Not Is Null"
                    AfterUpdate ="[Event Procedure]"
                    OnEnter ="[Event Procedure]"
                    OnExit ="[Event Procedure]"
                    Tag ="4"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6687
                            Top =3231
                            Width =855
                            Height =240
                            Name ="Label35"
                            Caption ="واحدكالا"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    FELineBreak = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =1
                    Left =3968
                    Top =3514
                    ColumnOrder =6
                    TabIndex =8
                    Name ="RADIF"
                    ControlSource ="RADIF"
                    DefaultValue ="[RecordsetClone].[RecordCount]+1"
                    AsianLineBreak =0
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    FELineBreak = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4251
                    Top =1814
                    Width =1371
                    Height =255
                    ColumnWidth =1980
                    ColumnOrder =7
                    TabIndex =9
                    ForeColor =8388608
                    Name ="CDATE"
                    ControlSource ="CDATE"
                    DefaultValue ="Now()"
                    Tag ="910922"
                    AsianLineBreak =0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =5661
                            Top =1814
                            Width =585
                            Height =300
                            ForeColor =8388608
                            Name ="Label24"
                            Caption ="تاريخ "
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "Visit_gol_dtl_FORM.cls"
