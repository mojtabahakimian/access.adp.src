Version =20
VersionRequired =20
Begin Form
    KeyPreview = NotDefault
    DefaultView =2
    AllowUpdating =4
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =9810
    RowHeight =315
    ItemSuffix =55
    Left =60
    Top =1140
    Right =23775
    Bottom =8670
    HelpContextId =5003
    DatasheetBackColor =13434828
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xa6e51a016ba8e240
    End
    RecordSource ="PGET_LST"
    OnCurrent ="[Event Procedure]"
    BeforeInsert ="[Event Procedure]"
    OnDelete ="[Event Procedure]"
    BeforeUpdate ="[Event Procedure]"
    AfterUpdate ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    HelpFile ="dena.hlp"
    DatasheetFontName ="Tahoma"
    OnKeyPress ="[Event Procedure]"
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
        Begin FormHeader
            Height =0
            BackColor =-2147483633
            Name ="FormHeader"
        End
        Begin Section
            Height =7680
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5955
                    Top =30
                    Height =255
                    ColumnWidth =1417
                    ColumnOrder =0
                    Name ="DATE"
                    ControlSource ="DATE"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =9137
                            Top =30
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
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5955
                    Top =370
                    Height =255
                    ColumnWidth =1417
                    ColumnOrder =1
                    TabIndex =1
                    Name ="RADIF"
                    ControlSource ="RADIF"
                    DefaultValue ="=[RecordsetClone].[RecordCount]+1"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =9122
                            Top =370
                            Width =405
                            Height =240
                            Name ="Label3"
                            Caption ="رديف"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5955
                    Top =4081
                    Height =255
                    ColumnWidth =1620
                    ColumnOrder =15
                    TabIndex =7
                    Name ="MABL"
                    ControlSource ="MABL"
                    Format ="#,###"
                    ValidationRule ="<>0 And Is Not Null"
                    OnExit ="[Event Procedure]"
                    DefaultValue ="0"
                    OnKeyPress ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =9152
                            Top =4081
                            Width =375
                            Height =240
                            Name ="Label19"
                            Caption ="مبلغ"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5955
                    Top =4421
                    Height =255
                    ColumnWidth =0
                    ColumnOrder =16
                    TabIndex =14
                    Name ="N_SERI"
                    ControlSource ="N_SERI"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =8522
                            Top =4421
                            Width =1005
                            Height =240
                            Name ="Label21"
                            Caption ="شماره سريال"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5955
                    Top =4762
                    Height =255
                    ColumnWidth =0
                    ColumnOrder =17
                    TabIndex =15
                    Name ="BANK"
                    ControlSource ="BANK"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =9152
                            Top =4762
                            Width =375
                            Height =240
                            Name ="Label23"
                            Caption ="بانك"
                        End
                    End
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =5955
                    Top =710
                    Height =255
                    ColumnWidth =960
                    ColumnOrder =2
                    TabIndex =2
                    Name ="NO_AM"
                    ControlSource ="NO_AM"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TCOD_DPS.CODE, TCOD_DPS.NAMES FROM TCOD_DPS ORDER BY TCOD_DPS.CODE, TCOD_"
                        "DPS.NAMES;"
                    ColumnWidths ="0"
                    BeforeUpdate ="[Event Procedure]"
                    AfterUpdate ="[Event Procedure]"
                    OnDblClick ="[Event Procedure]"
                    OnNotInList ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =8670
                            Top =705
                            Width =945
                            Height =240
                            Name ="Label5"
                            Caption ="نوع عمليات*"
                        End
                    End
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =5955
                    Top =1050
                    Height =255
                    ColumnWidth =1440
                    ColumnOrder =3
                    TabIndex =3
                    Name ="NAHVA"
                    ControlSource ="NAHVA"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TCOD_DPSKIND.CODE, TCOD_DPSKIND.NAMES FROM TCOD_DPSKIND ORDER BY TCOD_DPS"
                        "KIND.CODE, TCOD_DPSKIND.NAMES;"
                    ColumnWidths ="0"
                    BeforeUpdate ="[Event Procedure]"
                    AfterUpdate ="[Event Procedure]"
                    OnDblClick ="[Event Procedure]"
                    OnNotInList ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =9150
                            Top =1050
                            Width =465
                            Height =240
                            Name ="Label7"
                            Caption ="نحوه*"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5717
                    Top =5695
                    Height =255
                    ColumnWidth =1417
                    ColumnOrder =18
                    TabIndex =16
                    Name ="Text28"
                    ControlSource ="N_SERI"
                End
                Begin TextBox
                    Visible = NotDefault
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5730
                    Top =6015
                    Height =255
                    ColumnWidth =1440
                    ColumnOrder =19
                    TabIndex =17
                    Name ="Text30"
                    ControlSource ="BANK"
                End
                Begin TextBox
                    Visible = NotDefault
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4696
                    Top =226
                    Width =397
                    Height =114
                    ColumnWidth =1417
                    ColumnOrder =20
                    TabIndex =18
                    Name ="upd"
                    DefaultValue ="False"
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =801
                    Top =3741
                    Width =7371
                    ColumnWidth =15802
                    ColumnOrder =14
                    TabIndex =6
                    Name ="SHARH"
                    ControlSource ="SHARH"
                    AfterUpdate ="[Event Procedure]"
                    OnDblClick ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =9107
                            Top =3741
                            Width =420
                            Height =240
                            Name ="Label17"
                            Caption ="شرح*"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =4470
                    Top =1391
                    Width =3216
                    Height =255
                    ColumnWidth =660
                    ColumnOrder =10
                    TabIndex =11
                    Name ="FHES_K"
                    ControlSource ="FHES_K"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =8597
                            Top =1391
                            Width =930
                            Height =240
                            Name ="Label9"
                            Caption ="از  كل"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =4470
                    Top =1731
                    Width =3186
                    Height =255
                    ColumnWidth =825
                    ColumnOrder =11
                    TabIndex =12
                    Name ="FHES_M"
                    ControlSource ="FHES_M"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =8387
                            Top =1731
                            Width =1140
                            Height =240
                            Name ="Label11"
                            Caption =" از  معين"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =4470
                    Top =2106
                    Width =3186
                    Height =255
                    ColumnWidth =780
                    ColumnOrder =12
                    TabIndex =13
                    Name ="FHES_T"
                    ControlSource ="FHES_T"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =8387
                            Top =2106
                            Width =1140
                            Height =240
                            Name ="Label35"
                            Caption =" ازتفضيلي"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =42
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =4526
                    Top =2494
                    Width =3186
                    Height =255
                    ColumnWidth =660
                    ColumnOrder =6
                    TabIndex =8
                    Name ="THES_K"
                    ControlSource ="THES_K"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =8623
                            Top =2494
                            Width =960
                            Height =240
                            Name ="Label13"
                            Caption ="به  كل"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =4586
                    Top =2947
                    Width =3126
                    Height =255
                    ColumnWidth =690
                    ColumnOrder =7
                    TabIndex =9
                    Name ="THES_M"
                    ControlSource ="THES_M"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =8458
                            Top =2947
                            Width =1125
                            Height =240
                            Name ="Label15"
                            Caption ="به  معين"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =215
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =4529
                    Top =6462
                    Width =5046
                    Height =255
                    ColumnWidth =1680
                    ColumnOrder =5
                    TabIndex =4
                    Name ="FHES"
                    ControlSource ="FHES"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT hes, NAME FROM CUST_HESAB"
                    ColumnWidths ="1418;2835"
                    AfterUpdate ="[Event Procedure]"
                    OnExit ="[Event Procedure]"
                    OnDblClick ="[Event Procedure]"
                    OnNotInList ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =9039
                            Top =6465
                            Width =690
                            Height =240
                            Name ="Label39"
                            Caption ="از حساب*"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =215
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =4529
                    Top =6803
                    Width =5046
                    Height =255
                    ColumnWidth =1470
                    ColumnOrder =4
                    TabIndex =5
                    Name ="THES"
                    ControlSource ="THES"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT hes, NAME FROM CUST_HESAB"
                    ColumnWidths ="1418;2835"
                    AfterUpdate ="[Event Procedure]"
                    OnExit ="[Event Procedure]"
                    OnDblClick ="[Event Procedure]"
                    OnNotInList ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =9039
                            Top =6810
                            Width =720
                            Height =240
                            Name ="Label41"
                            Caption ="به حساب*"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =5320
                    Top =5102
                    ColumnOrder =21
                    TabIndex =19
                    Name ="ID"
                    ControlSource ="ID"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =7102
                            Top =5102
                            Width =1620
                            Height =240
                            Name ="Label42"
                            Caption ="شماره دريافت پرداخت:"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4470
                    Top =7320
                    Height =255
                    ColumnWidth =0
                    ColumnOrder =22
                    TabIndex =20
                    Name ="ARZD"
                    ControlSource ="ARZD"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =7412
                            Top =7320
                            Width =630
                            Height =240
                            Name ="Label33"
                            Caption ="نرخ ارز"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1587
                    Top =963
                    Width =1146
                    Height =255
                    ColumnWidth =975
                    ColumnOrder =13
                    TabIndex =21
                    Name ="FHES_T2"
                    ControlSource ="FHES_T2"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3464
                            Top =963
                            Width =870
                            Height =240
                            Name ="Label44"
                            Caption =" ازتفضيلي2"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =4586
                    Top =3322
                    Width =3126
                    ColumnWidth =750
                    ColumnOrder =8
                    TabIndex =10
                    Name ="THES_T"
                    ControlSource ="THES_T"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =8458
                            Top =3322
                            Width =1125
                            Height =240
                            Name ="Label37"
                            Caption ="به تفضيلي"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =453
                    Top =4081
                    Width =2886
                    ColumnWidth =1005
                    ColumnOrder =9
                    TabIndex =22
                    Name ="THES_T2"
                    ControlSource ="THES_T2"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3800
                            Top =4081
                            Width =900
                            Height =240
                            Name ="Label46"
                            Caption ="به تفضيلي2"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =453
                    Top =4441
                    Width =2946
                    Height =255
                    ColumnWidth =1417
                    TabIndex =23
                    Name ="FHES_T3"
                    ControlSource ="FHES_T3"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3845
                            Top =4441
                            Width =870
                            Height =240
                            Name ="Label48"
                            Caption =" ازتفضيلي3"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =453
                    Top =4820
                    Width =2886
                    ColumnWidth =1417
                    TabIndex =24
                    Name ="THES_T3"
                    ControlSource ="THES_T3"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3800
                            Top =4820
                            Width =900
                            Height =240
                            Name ="Label50"
                            Caption ="به تفضيلي3"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =453
                    Top =5180
                    Width =2946
                    Height =255
                    ColumnWidth =1417
                    TabIndex =25
                    Name ="FHES_T4"
                    ControlSource ="FHES_T4"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3845
                            Top =5180
                            Width =870
                            Height =240
                            Name ="Label52"
                            Caption =" ازتفضيلي4"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =453
                    Top =5530
                    Width =2886
                    ColumnWidth =1417
                    TabIndex =26
                    Name ="THES_T4"
                    ControlSource ="THES_T4"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3800
                            Top =5530
                            Width =900
                            Height =240
                            Name ="Label54"
                            Caption ="به تفضيلي4"
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
                    IMESentenceMode =3
                    ColumnCount =2
                    Width =2481
                    ColumnWidth =3000
                    TabIndex =27
                    Name ="MHAZ_NO"
                    ControlSource ="MHAZ_NO"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT MHAZ_NO, MHAZNAME FROM TCOD_MARKAZHAZ"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2532
                            Width =870
                            Height =240
                            Name ="Label40"
                            Caption ="مركز هزينه"
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
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2097
                    Top =113
                    Name ="Text26"
                    ControlSource ="=Sum([MABL])"
                    Format ="#,###"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5294
                            Top =113
                            Width =375
                            Height =240
                            Name ="Label27"
                            Caption ="مبلغ"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "PGET_LST_SUB.cls"
