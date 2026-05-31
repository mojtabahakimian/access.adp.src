Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    AllowEdits = NotDefault
    DefaultView =0
    AllowUpdating =4
    ScrollBars =0
    ViewsAllowed =1
    TabularCharSet =178
    TabularFamily =42
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =10110
    DatasheetFontHeight =10
    ItemSuffix =66
    Left =3270
    Top =915
    Right =13125
    Bottom =6570
    HelpContextId =8002
    DatasheetGridlinesColor =12632256
    MaxRecords =10000000
    Filter ="((PAY_GETP.N_SERI=235378))"
    RecSrcDt = Begin
        0x7daeb612fb62e240
    End
    RecordSource ="PAY_GETP"
    Caption ="چكهاي  پرداختي"
    OnCurrent ="[Event Procedure]"
    OnDelete ="[Event Procedure]"
    BeforeUpdate ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    HelpFile ="dena.hlp"
    DatasheetFontName ="Arial (Arabic)"
    OnLoad ="[Event Procedure]"
    OnError ="[Event Procedure]"
    AllowDatasheetView =0
    Begin
        Begin Label
            BackStyle =0
            TextFontCharSet =178
            FontName ="Tahoma"
        End
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
            Width =850
            Height =850
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
        Begin UnboundObjectFrame
            SpecialEffect =2
            OldBorderStyle =1
            Width =4536
            Height =2835
        End
        Begin CustomControl
            SpecialEffect =2
            Width =4536
            Height =2835
        End
        Begin FormHeader
            Height =656
            BackColor =8421376
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =93
                    TextFontFamily =2
                    Left =3855
                    Top =11
                    Width =2415
                    Height =600
                    FontSize =16
                    FontWeight =700
                    ForeColor =16777215
                    Name ="Label60"
                    Caption ="دفتر چكهاي پرداختي"
                    FontName ="Titr"
                End
                Begin Label
                    OverlapFlags =215
                    TextFontFamily =2
                    Left =3860
                    Width =2415
                    Height =600
                    FontSize =16
                    FontWeight =700
                    Name ="Label61"
                    Caption ="دفتر چكهاي پرداختي"
                    FontName ="Titr"
                End
            End
        End
        Begin Section
            SpecialEffect =1
            Height =5782
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =6405
                    Top =540
                    Width =2076
                    Height =315
                    ColumnOrder =28
                    Name ="N_SERI"
                    ControlSource ="N_SERI"
                    Format ="Fixed"
                    OnDblClick ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextFontCharSet =0
                            Left =8633
                            Top =540
                            Width =1080
                            Height =300
                            Name ="Label1"
                            Caption ="شماره سريال:"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =6977
                    Top =1835
                    Width =1506
                    Height =315
                    ColumnOrder =25
                    TabIndex =3
                    Name ="DATE"
                    ControlSource ="DATE"
                    DefaultValue ="=FARSIDATE(Date())"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextFontCharSet =0
                            Left =8642
                            Top =1830
                            Width =1110
                            Height =300
                            Name ="Label5"
                            Caption ="تاريخ پرداخت:"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =6977
                    Top =2265
                    Width =1506
                    Height =315
                    ColumnOrder =24
                    TabIndex =4
                    Name ="DATE_S"
                    ControlSource ="DATE_S"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextFontCharSet =0
                            Left =8635
                            Top =2265
                            Width =1110
                            Height =300
                            Name ="Label7"
                            Caption ="تاريخ سررسيد:"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =6429
                    Top =2691
                    Width =2061
                    Height =315
                    ColumnOrder =23
                    TabIndex =5
                    Name ="MABL"
                    ControlSource ="MABL"
                    Format ="#,##0.00\" ريال\";-#,##0.00\" ريال\""
                    OnKeyPress ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextFontCharSet =0
                            Left =8642
                            Top =2685
                            Width =480
                            Height =300
                            Name ="Label11"
                            Caption ="مبلغ:"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =6120
                    Top =3525
                    Width =2376
                    Height =315
                    ColumnOrder =21
                    TabIndex =7
                    Name ="N_HESAB"
                    ControlSource ="N_HESAB"
                    Format ="Fixed"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextFontCharSet =0
                            Left =8648
                            Top =3529
                            Width =1035
                            Height =300
                            Name ="Label15"
                            Caption ="جاري چك:"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =0
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =6165
                    Top =5267
                    Width =1596
                    Height =300
                    ColumnOrder =1
                    TabIndex =27
                    BackColor =8421376
                    ForeColor =16777215
                    Name ="N_S"
                    ControlSource ="N_S"
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =7772
                            Top =5265
                            Width =1080
                            Height =300
                            BackColor =8421376
                            Name ="Label31"
                            Caption ="شماره سند:"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3585
                    Top =5265
                    Width =1461
                    Height =300
                    ColumnOrder =0
                    TabIndex =28
                    BackColor =8421376
                    ForeColor =16777215
                    Name ="NUMBER"
                    ControlSource ="NUMBER"
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =87
                            TextFontCharSet =0
                            Left =5055
                            Top =5265
                            Width =1110
                            Height =300
                            BackColor =8421376
                            Name ="Label35"
                            Caption ="شماره فاكتور:"
                        End
                    End
                End
                Begin ComboBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =5979
                    Top =960
                    Width =2511
                    Height =315
                    ColumnOrder =27
                    TabIndex =1
                    Name ="BANK"
                    ControlSource ="BANK"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TCOD_BANKS.CODE, TCOD_BANKS.NAMES FROM TCOD_BANKS ORDER BY TCOD_BANKS.NAM"
                        "ES;"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextFontCharSet =0
                            Left =8642
                            Top =965
                            Width =450
                            Height =300
                            Name ="Label3"
                            Caption ="بانك:"
                        End
                    End
                End
                Begin Rectangle
                    OverlapFlags =255
                    Left =5120
                    Top =56
                    Width =4876
                    Height =4814
                    Name ="Box42"
                End
                Begin ComboBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =247
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5537
                    Top =1395
                    Width =2955
                    Height =315
                    ColumnOrder =26
                    TabIndex =2
                    Name ="SHOBEH"
                    ControlSource ="SHOBEH"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT PAY_GETP.SHOBEH FROM PAY_GETP GROUP BY PAY_GETP.SHOBEH ORDER BY PAY_GETP."
                        "SHOBEH;"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextFontCharSet =0
                            Left =8644
                            Top =1399
                            Width =945
                            Height =300
                            Name ="Label9"
                            Caption ="نام شعبه:"
                        End
                    End
                End
                Begin Rectangle
                    BorderWidth =3
                    OverlapFlags =93
                    Left =118
                    Top =60
                    Width =4876
                    Height =4814
                    Name ="Box43"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =215
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2508
                    Top =505
                    Width =1206
                    Height =315
                    ColumnOrder =19
                    TabIndex =9
                    Name ="N_KOL"
                    ControlSource ="N_KOL"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TOTA_HES.NUMBER FROM TOTA_HES ORDER BY TOTA_HES.NUMBER;"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextFontCharSet =0
                            Left =3855
                            Top =510
                            Width =900
                            Height =300
                            Name ="Label17"
                            Caption ="حساب كل:"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =215
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2510
                    Top =865
                    Width =1206
                    Height =315
                    ColumnOrder =17
                    TabIndex =11
                    Name ="N_MOIN"
                    ControlSource ="N_MOIN"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT DETA_HES.NUMBER FROM DETA_HES WHERE (((DETA_HES.N_KOL)=[Forms]![PAY_GETP]"
                        "![N_KOL])) GROUP BY DETA_HES.NUMBER ORDER BY DETA_HES.NUMBER;"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextFontCharSet =0
                            Left =3857
                            Top =860
                            Width =1065
                            Height =300
                            Name ="Label19"
                            Caption ="حساب معين:"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =215
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2505
                    Top =1915
                    Width =1206
                    Height =315
                    ColumnOrder =13
                    TabIndex =15
                    Name ="N_KOL2"
                    ControlSource ="N_KOL2"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TOTA_HES.NUMBER FROM TOTA_HES ORDER BY TOTA_HES.NUMBER;"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextFontCharSet =0
                            Left =3858
                            Top =1915
                            Width =975
                            Height =300
                            Name ="Label21"
                            Caption ="حساب كل:"
                        End
                    End
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =215
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2505
                    Top =2275
                    Width =1206
                    Height =315
                    ColumnOrder =11
                    TabIndex =17
                    Name ="N_MOIN2"
                    ControlSource ="N_MOIN2"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT DETA_HES.NUMBER FROM DETA_HES WHERE (((DETA_HES.N_KOL)=[Forms]![PAY_GETP]"
                        "![N_KOL2])) GROUP BY DETA_HES.NUMBER ORDER BY DETA_HES.NUMBER;"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextFontCharSet =0
                            Left =3858
                            Top =2275
                            Width =1005
                            Height =300
                            Name ="Label23"
                            Caption ="حساب معين:"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =215
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2505
                    Top =3325
                    Width =1206
                    Height =315
                    ColumnOrder =7
                    TabIndex =21
                    Name ="N_KOL3"
                    ControlSource ="N_KOL3"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TOTA_HES.NUMBER FROM TOTA_HES ORDER BY TOTA_HES.NUMBER;"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextFontCharSet =0
                            Left =3858
                            Top =3325
                            Width =960
                            Height =300
                            Name ="Label25"
                            Caption ="حساب كل:"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =215
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2505
                    Top =3685
                    Width =1206
                    Height =315
                    ColumnOrder =5
                    TabIndex =23
                    Name ="N_MOIN3"
                    ControlSource ="N_MOIN3"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT DETA_HES.NUMBER FROM DETA_HES WHERE (((DETA_HES.N_KOL)=[Forms]![PAY_GETP]"
                        "![N_KOL3])) GROUP BY DETA_HES.NUMBER ORDER BY DETA_HES.NUMBER;"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextFontCharSet =0
                            Left =3858
                            Top =3685
                            Width =1065
                            Height =300
                            Name ="Label27"
                            Caption ="حساب معين :"
                        End
                    End
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =215
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =193
                    Top =495
                    Width =2226
                    Height =315
                    ColumnOrder =18
                    TabIndex =10
                    Name ="AN_KOL"
                    ControlSource ="N_KOL"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TOTA_HES.NUMBER, TOTA_HES.NAME FROM TOTA_HES ORDER BY TOTA_HES.NAME;"
                    ColumnWidths ="0"
                    AfterUpdate ="[Event Procedure]"
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =215
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =195
                    Top =855
                    Width =2226
                    Height =315
                    ColumnOrder =16
                    TabIndex =12
                    Name ="AN_MOIN"
                    ControlSource ="N_MOIN"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT DETA_HES.NUMBER, DETA_HES.NAME FROM DETA_HES WHERE (((DETA_HES.N_KOL)=[Fo"
                        "rms]![PAY_GETP]![N_KOL])) GROUP BY DETA_HES.NUMBER, DETA_HES.NAME ORDER BY DETA_"
                        "HES.NAME;"
                    ColumnWidths ="0"
                    AfterUpdate ="[Event Procedure]"
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =215
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =195
                    Top =1905
                    Width =2226
                    Height =315
                    ColumnOrder =12
                    TabIndex =16
                    Name ="AN_KOL2"
                    ControlSource ="N_KOL2"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TOTA_HES.NUMBER, TOTA_HES.NAME FROM TOTA_HES ORDER BY TOTA_HES.NAME;"
                    ColumnWidths ="0"
                    AfterUpdate ="[Event Procedure]"
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =215
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =195
                    Top =2265
                    Width =2226
                    Height =315
                    ColumnOrder =10
                    TabIndex =18
                    Name ="AN_MOIN2"
                    ControlSource ="N_MOIN2"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT DETA_HES.NUMBER, DETA_HES.NAME FROM DETA_HES WHERE (((DETA_HES.N_KOL)=[Fo"
                        "rms]![PAY_GETP]![N_KOL2])) GROUP BY DETA_HES.NUMBER, DETA_HES.NAME ORDER BY DETA"
                        "_HES.NAME;"
                    ColumnWidths ="0"
                    AfterUpdate ="[Event Procedure]"
                End
                Begin ComboBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =215
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =195
                    Top =3315
                    Width =2226
                    Height =315
                    ColumnOrder =6
                    TabIndex =22
                    Name ="AN_KOL3"
                    ControlSource ="N_KOL3"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TOTA_HES.NUMBER, TOTA_HES.NAME FROM TOTA_HES ORDER BY TOTA_HES.NAME;"
                    ColumnWidths ="0"
                    AfterUpdate ="[Event Procedure]"
                End
                Begin ComboBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =215
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =195
                    Top =3675
                    Width =2226
                    Height =315
                    ColumnOrder =4
                    TabIndex =24
                    Name ="AN_MOIN3"
                    ControlSource ="N_MOIN3"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT DETA_HES.NUMBER, DETA_HES.NAME FROM DETA_HES WHERE (((DETA_HES.N_KOL)=[Fo"
                        "rms]![PAY_GETP]![N_KOL3])) GROUP BY DETA_HES.NUMBER, DETA_HES.NAME ORDER BY DETA"
                        "_HES.NAME;"
                    ColumnWidths ="0"
                    AfterUpdate ="[Event Procedure]"
                End
                Begin Label
                    OverlapFlags =215
                    TextFontCharSet =0
                    Left =3240
                    Top =120
                    Width =1695
                    Height =300
                    Name ="Label56"
                    Caption ="پرداخت از محل حساب:"
                End
                Begin Label
                    OverlapFlags =215
                    TextFontCharSet =0
                    Left =3460
                    Top =1560
                    Width =1425
                    Height =300
                    Name ="Label57"
                    Caption ="برگشت به حساب:"
                End
                Begin Label
                    OverlapFlags =215
                    TextFontCharSet =0
                    Left =3625
                    Top =3015
                    Width =1350
                    Height =300
                    Name ="Label58"
                    Caption ="*وصول از حساب:"
                    OnDblClick ="[Event Procedure]"
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =0
                    Left =8799
                    Top =170
                    Width =1185
                    Height =300
                    Name ="Label59"
                    Caption ="مشخصات چك:"
                End
                Begin ComboBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =247
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5190
                    Top =3106
                    Width =3300
                    Height =315
                    ColumnOrder =22
                    TabIndex =6
                    Name ="NAME_TAH"
                    ControlSource ="NAME_TAH"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT PAY_GETP.NAME_TAH FROM PAY_GETP GROUP BY PAY_GETP.NAME_TAH ORDER BY PAY_G"
                        "ETP.NAME_TAH;"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextFontCharSet =0
                            Left =8642
                            Top =3105
                            Width =1335
                            Height =300
                            Name ="Label13"
                            Caption ="نام دريافت كننده:"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1410
                    Top =5265
                    Width =1161
                    Height =300
                    ColumnOrder =29
                    TabIndex =29
                    BackColor =8421376
                    ForeColor =16777215
                    Name ="TAG"
                    ControlSource ="TAG"
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =2580
                            Top =5265
                            Width =990
                            Height =300
                            BackColor =8421376
                            Name ="Label63"
                            Caption ="نوع فاكتور :"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =247
                    TextFontCharSet =0
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =7650
                    Top =3933
                    Width =831
                    Height =315
                    ColumnOrder =20
                    TabIndex =8
                    Name ="RADIF"
                    ControlSource ="RADIF"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextFontCharSet =0
                            Left =8669
                            Top =3930
                            Width =1035
                            Height =300
                            Name ="Label64"
                            Caption ="رديف دفتر:"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =215
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2512
                    Top =1215
                    Width =1206
                    Height =315
                    ColumnOrder =15
                    TabIndex =13
                    Name ="N_TAF"
                    ControlSource ="N_TAF"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TDETA_HES.TNUMBER FROM TDETA_HES WHERE (((TDETA_HES.NUMBER)=[Forms]![PAY_"
                        "GETP]![N_MOIN]) AND ((TDETA_HES.N_KOL)=[Forms]![PAY_GETP]![N_KOL])) GROUP BY TDE"
                        "TA_HES.TNUMBER ORDER BY TDETA_HES.TNUMBER;"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextFontCharSet =0
                            Left =3859
                            Top =1210
                            Width =975
                            Height =300
                            Name ="Label67"
                            Caption ="تفصيلي:"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =215
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =197
                    Top =1205
                    Width =2226
                    Height =315
                    ColumnOrder =14
                    TabIndex =14
                    Name ="AN_TAF"
                    ControlSource ="N_TAF"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TDETA_HES.TNUMBER, TDETA_HES.NAME FROM TDETA_HES WHERE (((TDETA_HES.NUMBE"
                        "R)=[Forms]![PAY_GETP]![N_MOIN]) AND ((TDETA_HES.N_KOL)=[Forms]![PAY_GETP]![N_KOL"
                        "])) GROUP BY TDETA_HES.TNUMBER, TDETA_HES.NAME ORDER BY TDETA_HES.NAME;"
                    ColumnWidths ="0"
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =215
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2505
                    Top =2639
                    Width =1206
                    Height =315
                    ColumnOrder =9
                    TabIndex =19
                    Name ="N_TAF2"
                    ControlSource ="N_TAF2"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TDETA_HES.TNUMBER FROM TDETA_HES WHERE (((TDETA_HES.NUMBER)=[Forms]![PAY_"
                        "GETP]![N_MOIN2]) AND ((TDETA_HES.N_KOL)=[Forms]![PAY_GETP]![N_KOL2])) GROUP BY T"
                        "DETA_HES.TNUMBER ORDER BY TDETA_HES.TNUMBER;"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextFontCharSet =0
                            Left =3858
                            Top =2639
                            Width =975
                            Height =300
                            Name ="Label70"
                            Caption ="تفصيلي:"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =215
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =195
                    Top =2629
                    Width =2226
                    Height =315
                    ColumnOrder =8
                    TabIndex =20
                    Name ="AN_TAF2"
                    ControlSource ="N_TAF2"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TDETA_HES.TNUMBER, TDETA_HES.NAME FROM TDETA_HES WHERE (((TDETA_HES.NUMBE"
                        "R)=[Forms]![PAY_GETP]![N_MOIN2]) AND ((TDETA_HES.N_KOL)=[Forms]![PAY_GETP]![N_KO"
                        "L2])) GROUP BY TDETA_HES.TNUMBER, TDETA_HES.NAME ORDER BY TDETA_HES.NAME;"
                    ColumnWidths ="0"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =215
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2505
                    Top =4065
                    Width =1206
                    Height =315
                    ColumnOrder =3
                    TabIndex =25
                    Name ="N_TAF3"
                    ControlSource ="N_TAF3"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TDETA_HES.TNUMBER FROM TDETA_HES WHERE (((TDETA_HES.NUMBER)=[Forms]![PAY_"
                        "GETP]![N_MOIN3]) AND ((TDETA_HES.N_KOL)=[Forms]![PAY_GETP]![N_KOL3])) GROUP BY T"
                        "DETA_HES.TNUMBER ORDER BY TDETA_HES.TNUMBER;"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextFontCharSet =0
                            Left =3858
                            Top =4065
                            Width =1020
                            Height =300
                            Name ="Label73"
                            Caption ="تفصيلي:"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    Locked = NotDefault
                    OverlapFlags =215
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =195
                    Top =4055
                    Width =2226
                    Height =315
                    ColumnOrder =2
                    TabIndex =26
                    Name ="AN_TAF3"
                    ControlSource ="N_TAF3"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TDETA_HES.TNUMBER, TDETA_HES.NAME FROM TDETA_HES WHERE (((TDETA_HES.NUMBE"
                        "R)=[Forms]![PAY_GETP]![N_MOIN3]) AND ((TDETA_HES.N_KOL)=[Forms]![PAY_GETP]![N_KO"
                        "L3])) GROUP BY TDETA_HES.TNUMBER, TDETA_HES.NAME ORDER BY TDETA_HES.NAME;"
                    ColumnWidths ="0"
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =247
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5385
                    Top =4308
                    Width =3111
                    Height =315
                    TabIndex =30
                    Name ="SAYADI"
                    ControlSource ="SAYADI"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextFontCharSet =0
                            Left =8605
                            Top =4310
                            Width =1095
                            Height =240
                            Name ="Label78"
                            Caption ="شماره صيادي:"
                        End
                    End
                End
            End
        End
        Begin FormFooter
            Height =0
            BackColor =-2147483633
            Name ="FormFooter"
        End
    End
End
CodeBehindForm
' See "PAY_GETP.cls"
