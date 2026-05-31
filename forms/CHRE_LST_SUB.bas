Version =20
VersionRequired =20
Begin Form
    AllowAdditions = NotDefault
    KeyPreview = NotDefault
    AllowEdits = NotDefault
    DefaultView =2
    AllowUpdating =4
    ViewsAllowed =2
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =4314
    RowHeight =270
    ItemSuffix =34
    Left =330
    Top =1395
    Right =14745
    Bottom =7815
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x8add515e8e86e440
    End
    RecordSource ="CHRE_LIST_Q"
    Caption ="CHRE_LST"
    OnCurrent ="[Event Procedure]"
    OnDelete ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Tahoma"
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
            Height =4605
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
                    Left =485
                    Top =113
                    Height =255
                    ColumnOrder =0
                    Name ="DATE"
                    ControlSource ="DATE"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3667
                            Top =113
                            Width =390
                            Height =240
                            Name ="Label1"
                            Caption ="تاريخ"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =485
                    Top =453
                    Height =255
                    ColumnOrder =1
                    TabIndex =1
                    Name ="RADIF"
                    ControlSource ="RADIF"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3652
                            Top =453
                            Width =405
                            Height =240
                            Name ="Label3"
                            Caption ="رديف"
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
                    Left =485
                    Top =793
                    Width =1596
                    Height =255
                    ColumnWidth =1200
                    ColumnOrder =2
                    TabIndex =2
                    Name ="N_SERI"
                    ControlSource ="N_SERI"
                    RowSourceType ="Table/View/StoredProc"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3052
                            Top =793
                            Width =1005
                            Height =240
                            Name ="Label5"
                            Caption ="شماره سريال"
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
                    Left =485
                    Top =1133
                    Width =1596
                    Height =255
                    ColumnWidth =1680
                    ColumnOrder =3
                    TabIndex =3
                    Name ="BANK"
                    ControlSource ="BANK"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT CODE, NAMES FROM TCOD_BANKS"
                    ColumnWidths ="0"
                    OnEnter ="[Event Procedure]"
                    OnExit ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3682
                            Top =1133
                            Width =375
                            Height =240
                            Name ="Label7"
                            Caption ="بانك"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =87
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =226
                    Top =2891
                    Width =2106
                    Height =255
                    ColumnWidth =5970
                    ColumnOrder =8
                    TabIndex =4
                    Name ="HES1"
                    ControlSource ="HES1"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT hes, hes + N' : ' + NAME AS Expr1 FROM CUST_HESAB"
                    ColumnWidths ="0"
                    OnEnter ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =2328
                            Top =2891
                            Width =1005
                            Height =240
                            Name ="Label9"
                            Caption ="به حساب"
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
                    Left =510
                    Top =2154
                    ColumnWidth =1785
                    ColumnOrder =4
                    TabIndex =6
                    Name ="SHOBEH"
                    ControlSource ="SHOBEH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3177
                            Top =2154
                            Width =735
                            Height =240
                            Name ="Label22"
                            Caption ="شعبه"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =93
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =510
                    Top =2494
                    ColumnWidth =1560
                    ColumnOrder =6
                    TabIndex =7
                    Name ="MABL"
                    ControlSource ="MABL"
                    Format ="#,###"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3387
                            Top =2494
                            Width =525
                            Height =240
                            Name ="Label23"
                            Caption ="مبلغ"
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
                    Left =506
                    Top =1785
                    Width =1476
                    Height =255
                    ColumnWidth =1500
                    ColumnOrder =5
                    TabIndex =5
                    Name ="DATE_S"
                    ControlSource ="DATE_S"
                    RowSourceType ="Table/View/StoredProc"
                    OnEnter ="[Event Procedure]"
                    OnExit ="[Event Procedure]"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2798
                            Top =1785
                            Width =1110
                            Height =240
                            Name ="Label10"
                            Caption ="تاريخ سررسيد"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    Locked = NotDefault
                    ColumnHidden = NotDefault
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =510
                    Top =3194
                    Width =1881
                    Height =255
                    ColumnWidth =1965
                    TabIndex =8
                    Name ="N_MOIN"
                    ControlSource ="N_MOIN"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT DETA_HES.NUMBER, DETA_HES.NAME FROM DETA_HES WHERE (((DETA_HES.N_KOL)=[Fo"
                        "rms]![Baseknow]![BANKHA])) GROUP BY DETA_HES.NUMBER, DETA_HES.NAME ORDER BY DETA"
                        "_HES.NAME;"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2610
                            Top =3195
                            Width =1125
                            Height =240
                            Name ="Label26"
                            Caption ="به حساب معين"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =510
                    Top =3514
                    ColumnWidth =1417
                    TabIndex =9
                    Name ="N_KOL3"
                    ControlSource ="N_KOL3"
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =510
                    Top =3854
                    ColumnWidth =1417
                    TabIndex =10
                    Name ="N_MOIN3"
                    ControlSource ="N_MOIN3"
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =93
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =510
                    Top =4195
                    ColumnWidth =1417
                    TabIndex =11
                    Name ="N_TAF3"
                    ControlSource ="N_TAF3"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =2712
                            Top =4195
                            Width =1200
                            Height =240
                            Name ="Label31"
                            Caption ="حساب تفصيلي:"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =247
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =453
                    Top =2494
                    ColumnWidth =2430
                    TabIndex =12
                    Name ="N_S"
                    ControlSource ="N_S"
                End
                Begin TextBox
                    OverlapFlags =247
                    IMESentenceMode =3
                    Left =396
                    Top =4365
                    TabIndex =13
                    Name ="N_TAF"
                    ControlSource ="N_TAF"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =2598
                            Top =4365
                            Width =1200
                            Height =240
                            Name ="Label33"
                            Caption ="حساب تفصيلي:"
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
                    Left =284
                    Name ="smab"
                    ControlSource ="=Sum([MABL])"
                End
            End
        End
    End
End
CodeBehindForm
' See "CHRE_LST_SUB.cls"
