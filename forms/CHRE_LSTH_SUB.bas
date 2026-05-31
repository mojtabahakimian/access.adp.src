Version =20
VersionRequired =20
Begin Form
    KeyPreview = NotDefault
    DefaultView =2
    AllowUpdating =4
    ViewsAllowed =2
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =4314
    RowHeight =270
    ItemSuffix =39
    Left =435
    Top =1590
    Right =14625
    Bottom =8280
    DatasheetGridlinesColor =12632256
    MaxRecords =10000000
    RecSrcDt = Begin
        0x45c920a6b174e540
    End
    RecordSource ="SELECT PAY_GETD.SHOBEH, PAY_GETD.MABL, PAY_GETD.DATE_S AS Expr1, PAY_GETD.N_KOL,"
        " PAY_GETD.N_MOIN, PAY_GETD.N_TAF, PAY_GETD.N_S, CHRE_LSPH.BANK, CHRE_LSPH.DATE_S"
        ", CHRE_LSPH.N_SERI, CHRE_LSPH.DATE, CHRE_LSPH.RADIF, PAY_GETD.N_HESAB, PAY_GETD."
        "HES1 AS hes, CUST_HESAB.ECODE AS BKK, CHRE_LSPH.USER_NAME, PAY_GETD.HES1 FROM PA"
        "Y_GETD INNER JOIN CUST_HESAB ON PAY_GETD.HES1 = CUST_HESAB.hes RIGHT OUTER JOIN "
        "CHRE_LSPH ON PAY_GETD.N_SERI = CHRE_LSPH.N_SERI AND PAY_GETD.BANK = CHRE_LSPH.BA"
        "NK AND PAY_GETD.DATE_S = CHRE_LSPH.DATE_S"
    Caption ="CHRE_LST"
    OnCurrent ="[Event Procedure]"
    OnDelete ="[Event Procedure]"
    BeforeUpdate ="[Event Procedure]"
    AfterUpdate ="[Event Procedure]"
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
            Height =5045
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
                    ColumnWidth =1095
                    ColumnOrder =2
                    TabIndex =2
                    Name ="N_SERI"
                    ControlSource ="N_SERI"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT N_SERI FROM PAY_GETD WHERE (N_KOL IS NULL) AND (N_KOL2 IS NULL) AND (N_KO"
                        "L3 IS NULL) GROUP BY N_SERI"
                    AfterUpdate ="[Event Procedure]"
                    OnDblClick ="[Event Procedure]"
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
                    ColumnWidth =1335
                    ColumnOrder =3
                    TabIndex =3
                    Name ="BANK"
                    ControlSource ="BANK"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT PAY_GETD.BANK, TCOD_BANKS.NAMES FROM TCOD_BANKS INNER JOIN PAY_GETD ON TC"
                        "OD_BANKS.CODE = PAY_GETD.BANK GROUP BY PAY_GETD.BANK, TCOD_BANKS.NAMES"
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
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =340
                    Top =2834
                    Width =2106
                    Height =255
                    ColumnWidth =4380
                    ColumnOrder =8
                    TabIndex =4
                    Name ="hes"
                    ControlSource ="hes"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT hes, NAME FROM CUST_HESAB"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2652
                            Top =2834
                            Width =795
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
                    ColumnWidth =2820
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
                    ColumnWidth =1740
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
                    ColumnWidth =1230
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
                    ColumnWidths ="0"
                    DefaultValue ="1"
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
                    OverlapFlags =247
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =453
                    Top =2494
                    TabIndex =9
                    Name ="N_S"
                    ControlSource ="N_S"
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
                    Top =3569
                    Width =1881
                    Height =255
                    ColumnWidth =735
                    TabIndex =10
                    Name ="N_KOL"
                    ControlSource ="N_KOL"
                    RowSourceType ="Table/View/StoredProc"
                    ColumnWidths ="0"
                    DefaultValue ="=[Forms]![Baseknow]![BANKHA]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2610
                            Top =3570
                            Width =1125
                            Height =240
                            Name ="Label34"
                            Caption ="به حساب كل"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =169
                    Top =4025
                    ColumnWidth =1417
                    TabIndex =11
                    Name ="BKK"
                    ControlSource ="BKK"
                End
                Begin TextBox
                    OverlapFlags =247
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =170
                    Top =4251
                    Height =255
                    TabIndex =12
                    Name ="USER_NAME"
                    ControlSource ="USER_NAME"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3352
                            Top =4251
                            Width =390
                            Height =240
                            Name ="Label37"
                            Caption ="كاربر"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    ReadingOrder =2
                    ScrollBarAlign =1
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =283
                    Top =4705
                    TabIndex =13
                    Name ="N_TAF"
                    ControlSource ="N_TAF"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2485
                            Top =4705
                            Width =1200
                            Height =240
                            Name ="Label38"
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
' See "CHRE_LSTH_SUB.cls"
