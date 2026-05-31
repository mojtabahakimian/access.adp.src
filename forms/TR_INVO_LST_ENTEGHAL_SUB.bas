Version =20
VersionRequired =20
Begin Form
    AllowDeletions = NotDefault
    AllowAdditions = NotDefault
    OrderByOn = NotDefault
    AllowEdits = NotDefault
    DefaultView =2
    AllowUpdating =4
    ViewsAllowed =2
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =4769
    RowHeight =315
    ItemSuffix =51
    Left =465
    Top =2490
    Right =14415
    Bottom =8460
    HelpContextId =7001
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x0615699251d7e340
    End
    RecordSource ="TR_INVO_LST"
    Caption ="INVO_LST"
    HelpFile ="dena.hlp"
    DatasheetFontName ="Tahoma"
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
            Height =6859
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
                    Left =953
                    Top =480
                    Height =255
                    ColumnWidth =1417
                    ColumnOrder =0
                    Name ="NUMBER"
                    ControlSource ="NUMBER"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3565
                            Top =480
                            Width =960
                            Height =240
                            Name ="Label1"
                            Caption ="شماره فاكتور"
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
                    Left =953
                    Top =820
                    Height =255
                    ColumnWidth =1417
                    ColumnOrder =1
                    TabIndex =1
                    Name ="TAG"
                    ControlSource ="TAG"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3895
                            Top =820
                            Width =630
                            Height =240
                            Name ="Label3"
                            Caption ="برچسب"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =940
                    Top =1814
                    Height =255
                    ColumnWidth =1260
                    ColumnOrder =6
                    TabIndex =4
                    Name ="MEGH"
                    ControlSource ="MEGH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3837
                            Top =1814
                            Width =675
                            Height =240
                            Name ="Label11"
                            Caption ="مقداركالا"
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
                    Left =966
                    Top =2437
                    Height =255
                    ColumnOrder =7
                    TabIndex =5
                    Name ="MEGH_MAR"
                    ControlSource ="MEGH_MAR"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3428
                            Top =2437
                            Width =1110
                            Height =240
                            Name ="Label13"
                            Caption ="مقدار مرجوعي"
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
                    Left =966
                    Top =2777
                    Height =255
                    ColumnOrder =8
                    TabIndex =6
                    Name ="MANDAH"
                    ControlSource ="MANDAH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4118
                            Top =2777
                            Width =420
                            Height =240
                            Name ="Label15"
                            Caption ="مانده"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    ColumnHidden = NotDefault
                    TabStop = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =966
                    Top =3117
                    Height =255
                    ColumnWidth =1440
                    ColumnOrder =10
                    TabIndex =7
                    Name ="MABL"
                    ControlSource ="MABL"
                    Format ="#,##0.00;-#,##0.00"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4163
                            Top =3117
                            Width =375
                            Height =240
                            Name ="Label17"
                            Caption ="مبلغ"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    ColumnHidden = NotDefault
                    TabStop = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =966
                    Top =3457
                    Height =255
                    ColumnWidth =1785
                    ColumnOrder =11
                    TabIndex =8
                    Name ="MABL_K"
                    ControlSource ="MABL_K"
                    Format ="#,##0.00;-#,##0.00"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3923
                            Top =3457
                            Width =615
                            Height =240
                            Name ="Label19"
                            Caption ="مبلغ كل"
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
                    Left =966
                    Top =3797
                    Height =255
                    ColumnOrder =12
                    TabIndex =9
                    Name ="FROM_A"
                    ControlSource ="FROM_A"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3833
                            Top =3797
                            Width =705
                            Height =240
                            Name ="Label21"
                            Caption ="FROM_A"
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
                    Left =966
                    Top =4138
                    Height =255
                    ColumnOrder =13
                    TabIndex =10
                    Name ="N_RASID"
                    ControlSource ="N_RASID"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =3593
                            Top =4138
                            Width =945
                            Height =240
                            Name ="Label23"
                            Caption ="شماره رسيد"
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
                    Left =966
                    Top =4478
                    Height =255
                    ColumnOrder =14
                    TabIndex =11
                    Name ="MEGH_R"
                    ControlSource ="MEGH_R"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =3668
                            Top =4478
                            Width =870
                            Height =240
                            Name ="Label25"
                            Caption ="مقدار رسيد"
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
                    Left =966
                    Top =4818
                    Height =255
                    ColumnOrder =15
                    TabIndex =12
                    Name ="RADAH"
                    ControlSource ="RADAH"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =4253
                            Top =4818
                            Width =285
                            Height =240
                            Name ="Label27"
                            Caption ="رده"
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
                    Left =966
                    Top =5158
                    Height =255
                    ColumnOrder =16
                    TabIndex =13
                    Name ="SANAD_NO"
                    ControlSource ="SANAD_NO"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3653
                            Top =5158
                            Width =885
                            Height =240
                            Name ="Label29"
                            Caption ="شماره سند"
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
                    Left =966
                    Top =5498
                    Height =255
                    ColumnOrder =17
                    TabIndex =14
                    Name ="CUST_NO"
                    ControlSource ="CUST_NO"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3413
                            Top =5498
                            Width =1125
                            Height =240
                            Name ="Label31"
                            Caption ="شماره مشتري"
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
                    Left =966
                    Top =5838
                    Height =255
                    ColumnWidth =945
                    ColumnOrder =2
                    TabIndex =15
                    Name ="ANBARF"
                    ControlSource ="ANBARF"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3743
                            Top =5838
                            Width =795
                            Height =240
                            Name ="Label33"
                            Caption ="انبار فرعي"
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
                    ColumnCount =3
                    Left =961
                    Top =1474
                    Width =1680
                    Height =255
                    ColumnWidth =9075
                    ColumnOrder =4
                    TabIndex =3
                    Name ="CODE"
                    ControlSource ="CODE"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT STUF_DEF.CODE, STUF_DEF.NAME, STUF_DEF.CODE AS Expr1 FROM STUF_DEF INNER "
                        "JOIN STUF_FSK ON STUF_DEF.CODE = STUF_FSK.CODE WHERE (STUF_FSK.ANBAR = 2) ORDER "
                        "BY STUF_DEF.NAME"
                    ColumnWidths ="0;6237;567"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4002
                            Top =1474
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
                    OverlapFlags =85
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =964
                    Top =2097
                    Height =255
                    ColumnWidth =1545
                    ColumnOrder =9
                    TabIndex =17
                    Name ="MEGHk"
                    ControlSource ="MEGHk"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3346
                            Top =2097
                            Width =1020
                            Height =240
                            Name ="Label38"
                            Caption ="مقدار كل "
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
                    Left =966
                    Top =6179
                    Width =696
                    Height =255
                    ColumnOrder =5
                    TabIndex =16
                    Name ="VAHED_K"
                    ControlSource ="VAHED_K"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT VAHEDS.VAHED, TCOD_VAHEDS.NAMES FROM TCOD_VAHEDS INNER JOIN VAHEDS ON TCO"
                        "D_VAHEDS.CODE = VAHEDS.VAHED GROUP BY VAHEDS.VAHED, TCOD_VAHEDS.NAMES ORDER BY T"
                        "COD_VAHEDS.NAMES"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3923
                            Top =6179
                            Width =615
                            Height =240
                            Name ="Label35"
                            Caption ="واحدكالا"
                        End
                    End
                End
                Begin ComboBox
                    ColumnHidden = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =968
                    Top =1160
                    Width =1686
                    Height =255
                    ColumnWidth =1417
                    ColumnOrder =3
                    TabIndex =2
                    Name ="ANBAR"
                    ControlSource ="ANBAR"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT CODE, NAMES FROM TCOD_ANBAR WHERE (CODE <> 0) ORDER BY CODE"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4195
                            Top =1160
                            Width =330
                            Height =240
                            Name ="Label5"
                            Caption ="انبار"
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
                    Left =964
                    Top =170
                    Height =255
                    ColumnWidth =1417
                    TabIndex =18
                    Name ="FLG"
                    DefaultValue ="False"
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1360
                    Top =6519
                    TabIndex =19
                    Name ="RADIF"
                    ControlSource ="RADIF"
                    DefaultValue ="=[RecordsetClone].[RecordCount]+1"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4297
                            Top =6519
                            Width =465
                            Height =240
                            Name ="Label47"
                            Caption ="رديف:"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =247
                    IMESentenceMode =3
                    Left =1133
                    Top =4365
                    TabIndex =20
                    Name ="UP_DATE"
                    ControlSource ="UP_DATE"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =3725
                            Top =4365
                            Width =810
                            Height =240
                            Name ="Label49"
                            Caption ="UP_DATE:"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =247
                    IMESentenceMode =3
                    Left =1133
                    Top =4705
                    TabIndex =21
                    Name ="UP_TIME"
                    ControlSource ="UP_TIME"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =3755
                            Top =4705
                            Width =780
                            Height =240
                            Name ="Label50"
                            Caption ="UP_TIME:"
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
                    TabStop = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1078
                    Top =56
                    Height =255
                    Name ="SMABLK"
                    ControlSource ="=Sum([MABL_K])"
                End
            End
        End
    End
End
CodeBehindForm
' See "TR_INVO_LST_ENTEGHAL_SUB.cls"
