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
    Width =4938
    RowHeight =375
    ItemSuffix =21
    Left =405
    Top =1860
    Right =15240
    Bottom =8445
    DatasheetForeColor =33554432
    DatasheetGridlinesColor =12632256
    BeforeDelConfirm ="[Event Procedure]"
    RecSrcDt = Begin
        0x66413bea4f76e240
    End
    RecordSource ="DTL_MANF"
    Caption ="DTL_MANF"
    OnCurrent ="[Event Procedure]"
    OnDelete ="[Event Procedure]"
    BeforeUpdate ="[Event Procedure]"
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
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1109
                    Top =113
                    ColumnWidth =270
                    ColumnOrder =1
                    Name ="FNUMB"
                    ControlSource ="FNUMB"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3691
                            Top =113
                            Width =990
                            Height =240
                            Name ="Label1"
                            Caption ="شماره فرمول"
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
                    Top =1474
                    ColumnWidth =945
                    ColumnOrder =4
                    TabIndex =1
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
                    TabIndex =2
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
                    TabIndex =3
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
                    ColumnCount =3
                    Left =1085
                    Top =480
                    Width =1680
                    Height =255
                    ColumnWidth =5640
                    ColumnOrder =2
                    TabIndex =4
                    Name ="CODE"
                    ControlSource ="CODE"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT     TOP 100 PERCENT dbo.STUF_DEF.CODE, dbo.STUF_DEF.NAME, dbo.STUF_DEF.CO"
                        "DE AS Expr1 FROM dbo.STUF_DEF INNER JOIN dbo.STUF_FSK INNER JOIN dbo.STUF_STK ON"
                        " dbo.STUF_FSK.CODE = dbo.STUF_STK.CODE AND dbo.STUF_FSK.ANBAR = dbo.STUF_STK.ANB"
                        "AR ON dbo.STUF_DEF.CODE = dbo.STUF_FSK.CODE WHERE (dbo.STUF_FSK.ANBAR = 1)ORDER "
                        "BY dbo.STUF_DEF.NAME"
                    ColumnWidths ="0;4536;567"
                    AfterUpdate ="[Event Procedure]"
                    OnEnter ="[Event Procedure]"
                    OnExit ="[Event Procedure]"
                    OnNotInList ="[Event Procedure]"
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
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =1106
                    Top =802
                    Width =1686
                    Height =255
                    ColumnWidth =1395
                    ColumnOrder =0
                    TabIndex =5
                    Name ="ANBAR"
                    ControlSource ="ANBAR"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TCOD_ANBAR.CODE, TCOD_ANBAR.NAMES FROM TCOD_ANBAR ORDER BY TCOD_ANBAR.COD"
                        "E;"
                    ColumnWidths ="0"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="1"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4333
                            Top =802
                            Width =330
                            Height =240
                            Name ="Label5"
                            Caption ="انبار"
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
                    TabIndex =6
                    Name ="VAHED_K"
                    ControlSource ="VAHED_K"
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
                    ColumnWidth =1155
                    TabIndex =7
                    Name ="SMABL"
                    ControlSource ="SMABL"
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
                            Caption ="مبلغ استاندارد"
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
                    TabIndex =8
                    Name ="MABLK"
                    ControlSource ="MABLK"
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
' See "DTL_MANF_SUB1.cls"
