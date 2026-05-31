Version =20
VersionRequired =20
Begin Form
    DefaultView =2
    AllowUpdating =4
    ViewsAllowed =2
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =4938
    RowHeight =375
    ItemSuffix =21
    Left =510
    Top =1860
    Right =14610
    Bottom =7770
    DatasheetGridlinesColor =12632256
    BeforeDelConfirm ="[Event Procedure]"
    RecSrcDt = Begin
        0x46d62ec53a17e540
    End
    RecordSource ="SELECT QPROGPAS2.PRGID, QPROGPAS2.CODB, QPROGPAS2.VAHED, QPROGPAS2.SumOfMEGH AS "
        "MEGH, QPROGPAS2.SumOfMEGHK AS MEGHK, QPROGPAS2.SumOfPERT AS PERT, QPROGPAS2.SumO"
        "fKOLMAV AS KOLMAV, QPROGPAS2.PASED, QPROGPAS2.RADAH, STUF_DEF.NAME, QPROGPAS2.Su"
        "mOfMABL, QPROGPAS2.SumOfMABLK FROM QPROGPAS2 INNER JOIN STUF_DEF ON QPROGPAS2.CO"
        "DB = STUF_DEF.CODE"
    Caption ="DTL_MANF"
    OnDelete ="[Event Procedure]"
    DatasheetFontName ="Tahoma"
    OnLoad ="[Event Procedure]"
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
        Begin FormHeader
            Height =0
            BackColor =-2147483633
            Name ="FormHeader"
        End
        Begin Section
            Height =3074
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1109
                    Top =1474
                    ColumnWidth =945
                    ColumnOrder =4
                    Name ="MEGH"
                    ControlSource ="MEGH"
                    ValidationRule ="Is Not Null"
                    AfterUpdate ="[Event Procedure]"
                    OnKeyPress ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3901
                            Top =1474
                            Width =780
                            Height =240
                            Name ="Label9"
                            Caption ="مقدار مواد"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1109
                    Top =1814
                    ColumnWidth =1215
                    ColumnOrder =5
                    TabIndex =1
                    Name ="MEGHk"
                    ControlSource ="MEGHk"
                    ValidationRule ="Is Not Null"
                    OnKeyPress ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3661
                            Top =1814
                            Width =1020
                            Height =240
                            Name ="Label11"
                            Caption ="مقدار كل مواد"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1109
                    Top =2154
                    ColumnWidth =435
                    ColumnOrder =6
                    TabIndex =2
                    Name ="PERT"
                    ControlSource ="PERT"
                    ValidationRule ="Is Not Null"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4351
                            Top =2154
                            Width =330
                            Height =240
                            Name ="Label13"
                            Caption ="پرت"
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
                    Left =1085
                    Top =1125
                    Width =1686
                    Height =255
                    ColumnWidth =1215
                    ColumnOrder =3
                    TabIndex =4
                    Name ="VAHED_K"
                    ControlSource ="VAHED"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT VAHEDS.VAHED, TCOD_VAHEDS.NAMES FROM TCOD_VAHEDS INNER JOIN VAHEDS ON TCO"
                        "D_VAHEDS.CODE = VAHEDS.VAHED GROUP BY VAHEDS.VAHED, TCOD_VAHEDS.NAMES ORDER BY T"
                        "COD_VAHEDS.NAMES"
                    ColumnWidths ="0"
                    AfterUpdate ="[Event Procedure]"
                    OnEnter ="[Event Procedure]"
                    OnExit ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4042
                            Top =1125
                            Width =615
                            Height =240
                            Name ="Label35"
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
                    Left =1077
                    Top =2437
                    ColumnWidth =1440
                    TabIndex =5
                    Name ="SMABL"
                    ControlSource ="SumOfMABL"
                    Format ="#,###"
                    ValidationRule ="Is Not Null"
                    AfterUpdate ="[Event Procedure]"
                    OnKeyPress ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3369
                            Top =2437
                            Width =1110
                            Height =240
                            Name ="Label17"
                            Caption ="مبلغ "
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =793
                    Top =2834
                    ColumnWidth =1950
                    TabIndex =6
                    Name ="MABLK"
                    ControlSource ="SumOfMABLK"
                    Format ="#,###"
                    ValidationRule ="Is Not Null"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3520
                            Top =2834
                            Width =675
                            Height =240
                            Name ="Label18"
                            Caption ="مبلغ كل"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1085
                    Top =480
                    Width =1680
                    Height =255
                    ColumnWidth =6045
                    ColumnOrder =2
                    TabIndex =3
                    Name ="CODE"
                    ControlSource ="NAME"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3731
                            Top =487
                            Width =1155
                            Height =240
                            Name ="Label14"
                            Caption =" ماده اوليه"
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
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1055
                    Top =90
                    Name ="SKOL"
                    ControlSource ="=Sum([MEGHk])"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3607
                            Top =90
                            Width =1020
                            Height =240
                            Name ="Label16"
                            Caption ="مقدار كل مواد"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1055
                    Top =450
                    TabIndex =1
                    Name ="SMABK"
                    ControlSource ="=Sum([MABLK])"
                    Format ="#,###"
                End
            End
        End
    End
End
CodeBehindForm
' See "DTL_MANF_SUB1_DAY.cls"
