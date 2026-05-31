Version =20
VersionRequired =20
Begin Form
    AutoResize = NotDefault
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    DefaultView =2
    AllowUpdating =4
    ScrollBars =0
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =3973
    RowHeight =315
    ItemSuffix =10
    Left =750
    Top =975
    Right =4620
    Bottom =3210
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x5e769d722ef7e340
    End
    RecordSource ="TR_MODULE_D"
    Caption ="MODULE_D"
    OnCurrent ="[Event Procedure]"
    DatasheetFontName ="Tahoma"
    OnLoad ="[Event Procedure]"
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
            Height =3358
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
                    Left =405
                    Top =113
                    Width =1440
                    ColumnWidth =1417
                    Name ="CODE"
                    ControlSource ="CODE"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3251
                            Top =113
                            Width =465
                            Height =240
                            Name ="Label1"
                            Caption ="كدكالا"
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
                    Top =1133
                    ColumnWidth =810
                    TabIndex =2
                    Name ="NESBAT"
                    ControlSource ="NESBAT"
                    OnKeyPress ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3206
                            Top =1133
                            Width =510
                            Height =240
                            Name ="Label7"
                            Caption ="نسبت"
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
                    Top =1474
                    TabIndex =3
                    Name ="MABL_F"
                    ControlSource ="MABL_F"
                    OnKeyPress ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3401
                            Top =1474
                            Width =315
                            Height =240
                            Name ="Label9"
                            Caption ="في"
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
                    Top =793
                    Width =816
                    ColumnWidth =1365
                    TabIndex =1
                    Name ="VAHED"
                    ControlSource ="VAHED"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TCOD_VAHEDS.CODE, TCOD_VAHEDS.NAMES FROM TCOD_VAHEDS ORDER BY TCOD_VAHEDS"
                        ".NAMES;"
                    ColumnWidths ="0"
                    OnDblClick ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3101
                            Top =793
                            Width =615
                            Height =240
                            Name ="Label5"
                            Caption ="واحدكالا"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =112
                    Top =2664
                    TabIndex =4
                    Name ="UP_DATE"
                    ControlSource ="UP_DATE"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2621
                            Top =2664
                            Width =810
                            Height =240
                            Name ="Label48"
                            Caption ="UP_DATE:"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =56
                    Top =3118
                    TabIndex =5
                    Name ="UP_TIME"
                    ControlSource ="UP_TIME"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2678
                            Top =3118
                            Width =780
                            Height =240
                            Name ="Label49"
                            Caption ="UP_TIME:"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "TR_MODULE_D_SUB.cls"
