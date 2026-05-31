Version =20
VersionRequired =20
Begin Form
    AllowAdditions = NotDefault
    KeyPreview = NotDefault
    AllowEdits = NotDefault
    DefaultView =2
    AllowUpdating =4
    ViewsAllowed =2
    TabularCharSet =178
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =4308
    RowHeight =360
    ItemSuffix =35
    Left =405
    Top =1545
    Right =14655
    Bottom =8670
    DatasheetForeColor =33554432
    DatasheetGridlinesColor =12632256
    BeforeDelConfirm ="[Event Procedure]"
    RecSrcDt = Begin
        0x059c076c8e86e440
    End
    RecordSource ="CHREC_LSP_Q"
    Caption ="CHRE_LST"
    OnCurrent ="[Event Procedure]"
    OnDelete ="[Event Procedure]"
    AfterUpdate ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
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
            Height =4151
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
                    Left =483
                    Top =113
                    Height =255
                    ColumnOrder =0
                    Name ="DATE"
                    ControlSource ="DATE"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3665
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
                    OverlapFlags =93
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =483
                    Top =453
                    Height =255
                    ColumnOrder =1
                    TabIndex =1
                    Name ="RADIF"
                    ControlSource ="RADIF"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3650
                            Top =453
                            Width =405
                            Height =240
                            Name ="Label3"
                            Caption ="رديف"
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
                    Left =483
                    Top =1133
                    Width =936
                    Height =255
                    ColumnWidth =1335
                    ColumnOrder =3
                    TabIndex =3
                    Name ="BANK"
                    ControlSource ="BANK"
                    RowSourceType ="Table/View/StoredProc"
                    ColumnWidths ="0"
                    OnEnter ="[Event Procedure]"
                    OnExit ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3680
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
                    ColumnHidden = NotDefault
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =513
                    Top =1474
                    Width =906
                    Height =255
                    ColumnWidth =0
                    ColumnOrder =8
                    TabIndex =4
                    Name ="N_TAF"
                    ControlSource ="N_TAF"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TDETA_HES.TNUMBER, TDETA_HES.NAME FROM DETA_HES INNER JOIN TDETA_HES ON ("
                        "DETA_HES.NUMBER = TDETA_HES.NUMBER) AND (DETA_HES.N_KOL = TDETA_HES.N_KOL) WHERE"
                        " (((TDETA_HES.NUMBER)=1) AND ((TDETA_HES.N_KOL)=[Forms]![Baseknow]![BANKHA])) GR"
                        "OUP BY TDETA_HES.TNUMBER, TDETA_HES.NAME;"
                    ColumnWidths ="0"
                    OnEnter ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2615
                            Top =1474
                            Width =1005
                            Height =240
                            Name ="Label9"
                            Caption ="از حساب"
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
                    Left =508
                    Top =2154
                    ColumnWidth =1845
                    ColumnOrder =4
                    TabIndex =6
                    Name ="SHOBEH"
                    ControlSource ="SHOBEH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3175
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
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =508
                    Top =2494
                    ColumnWidth =1575
                    ColumnOrder =6
                    TabIndex =7
                    Name ="MABL"
                    ControlSource ="MABL"
                    Format ="#,###"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3385
                            Top =2494
                            Width =525
                            Height =240
                            Name ="Label23"
                            Caption ="مبلغ"
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
                    Left =508
                    Top =2834
                    Height =255
                    ColumnOrder =7
                    TabIndex =8
                    Name ="PDTS"
                    ControlSource ="PDTS"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2800
                            Top =2834
                            Width =1110
                            Height =240
                            Name ="Label24"
                            Caption ="تاريخ سررسيد:"
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
                    Left =508
                    Top =3194
                    Width =906
                    Height =255
                    ColumnWidth =2760
                    TabIndex =9
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
                            Left =2612
                            Top =3195
                            Width =1230
                            Height =240
                            Name ="Label26"
                            Caption ="پرداخت از حساب"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =93
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Top =3458
                    TabIndex =10
                    Name ="N_KOL3"
                    ControlSource ="N_KOL3"
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Top =3798
                    TabIndex =11
                    Name ="N_MOIN3"
                    ControlSource ="N_MOIN3"
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =247
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =850
                    Top =3514
                    ColumnWidth =3555
                    TabIndex =12
                    Name ="N_TAF3"
                    ControlSource ="N_TAF3"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2938
                            Top =3514
                            Width =1200
                            Height =240
                            Name ="Label31"
                            Caption ="حساب تفصيلي"
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
                    IMESentenceMode =3
                    Left =504
                    Top =1785
                    Width =921
                    Height =255
                    ColumnWidth =1395
                    ColumnOrder =5
                    TabIndex =5
                    Name ="DATE_S"
                    ControlSource ="DATE_S"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2796
                            Top =1785
                            Width =1110
                            Height =240
                            Name ="Label10"
                            Caption ="تاريخ سررسيد"
                        End
                    End
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =483
                    Top =793
                    Width =936
                    ColumnWidth =1095
                    ColumnOrder =2
                    TabIndex =2
                    Name ="N_SERI"
                    ControlSource ="N_SERI"
                    RowSourceType ="Table/View/StoredProc"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =3050
                            Top =793
                            Width =1005
                            Height =240
                            Name ="Label5"
                            Caption ="شماره سريال"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =247
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =283
                    Top =623
                    Width =2106
                    Height =255
                    ColumnWidth =6540
                    TabIndex =13
                    Name ="HES1"
                    ControlSource ="HES1"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT hes, hes + N' : ' + NAME AS Expr1 FROM CUST_HESAB"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =223
                            Left =2385
                            Top =623
                            Width =1005
                            Height =240
                            Name ="Label34"
                            Caption ="از حساب"
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
                    Left =964
                    Name ="smab"
                    ControlSource ="=Sum([MABL])"
                End
            End
        End
    End
End
CodeBehindForm
' See "CHRE_LSP_SUB.cls"
