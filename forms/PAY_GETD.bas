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
    Width =10941
    DatasheetFontHeight =10
    ItemSuffix =79
    Left =2865
    Top =870
    Right =13545
    Bottom =6660
    HelpContextId =8001
    DatasheetGridlinesColor =12632256
    MaxRecords =1000000
    Filter ="((PAY_GETD.N_SERI=836672))"
    RecSrcDt = Begin
        0x8e8a12ef4aa9e240
    End
    RecordSource ="PAY_GETD"
    Caption ="چكهاي دريافتي"
    OnCurrent ="[Event Procedure]"
    BeforeInsert ="[Event Procedure]"
    OnDelete ="[Event Procedure]"
    BeforeUpdate ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    HelpFile ="dena.hlp"
    DatasheetFontName ="Arial (Arabic)"
    OnLoad ="[Event Procedure]"
    OnError ="[Event Procedure]"
    AllowDatasheetView =0
    RecordSourceQualifier ="dbo"
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
            Height =793
            BackColor =8421376
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =93
                    TextFontFamily =2
                    Left =4125
                    Top =71
                    Width =2685
                    Height =675
                    FontSize =18
                    FontWeight =700
                    ForeColor =16777215
                    Name ="Label60"
                    Caption ="دفتر چكهاي دريافتي"
                    FontName ="Titr"
                End
                Begin Label
                    OverlapFlags =215
                    TextFontFamily =2
                    Left =4130
                    Top =60
                    Width =2685
                    Height =675
                    FontSize =18
                    FontWeight =700
                    Name ="Label61"
                    Caption ="دفتر چكهاي دريافتي"
                    FontName ="Titr"
                End
            End
        End
        Begin Section
            SpecialEffect =1
            Height =5612
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
                    Left =7409
                    Top =585
                    Width =1971
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
                            Left =9532
                            Top =585
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
                    Left =7846
                    Top =1775
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
                            Left =9504
                            Top =1770
                            Width =1095
                            Height =300
                            Name ="Label5"
                            Caption ="تاريخ دريافت:"
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
                    Left =7846
                    Top =2145
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
                            Left =9504
                            Top =2145
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
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =7298
                    Top =2526
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
                            Left =9511
                            Top =2520
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
                    Left =7064
                    Top =3285
                    Width =2301
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
                            Left =9517
                            Top =3289
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
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =7063
                    Top =5222
                    Width =1431
                    Height =300
                    ColumnOrder =1
                    TabIndex =28
                    BackColor =8421376
                    ForeColor =16777215
                    Name ="N_S"
                    ControlSource ="N_S"
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =8505
                            Top =5220
                            Width =1140
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
                    Left =4575
                    Top =5220
                    Width =1356
                    Height =300
                    ColumnOrder =0
                    TabIndex =29
                    BackColor =8421376
                    ForeColor =16777215
                    Name ="NUMBER"
                    ControlSource ="NUMBER"
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =5940
                            Top =5220
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
                    Left =6848
                    Top =985
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
                            Left =9511
                            Top =990
                            Width =450
                            Height =300
                            Name ="Label3"
                            Caption ="بانك:"
                        End
                    End
                End
                Begin Rectangle
                    OverlapFlags =255
                    Left =5944
                    Top =56
                    Width =4876
                    Height =4859
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
                    Left =6406
                    Top =1380
                    Width =2955
                    Height =315
                    ColumnOrder =26
                    TabIndex =2
                    Name ="SHOBEH"
                    ControlSource ="SHOBEH"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT PAY_GETD.SHOBEH FROM PAY_GETD GROUP BY PAY_GETD.SHOBEH ORDER BY PAY_GETD."
                        "SHOBEH;"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextFontCharSet =0
                            Left =9513
                            Top =1384
                            Width =960
                            Height =300
                            Name ="Label9"
                            Caption ="نام شعبه:"
                        End
                    End
                End
                Begin Rectangle
                    BorderWidth =3
                    OverlapFlags =93
                    Left =942
                    Top =60
                    Width =4876
                    Height =4859
                    Name ="Box43"
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
                    Left =3392
                    Top =520
                    Width =1206
                    Height =315
                    ColumnOrder =19
                    TabIndex =10
                    Name ="N_KOL"
                    ControlSource ="N_KOL"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TOTA_HES.NUMBER FROM TOTA_HES ORDER BY TOTA_HES.NUMBER;"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextFontCharSet =0
                            Left =4739
                            Top =525
                            Width =975
                            Height =300
                            Name ="Label17"
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
                    Left =3394
                    Top =880
                    Width =1206
                    Height =315
                    ColumnOrder =17
                    TabIndex =12
                    Name ="N_MOIN"
                    ControlSource ="N_MOIN"
                    RowSourceType ="Table/View/StoredProc"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextFontCharSet =0
                            Left =4741
                            Top =875
                            Width =1005
                            Height =300
                            Name ="Label19"
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
                    Left =3394
                    Top =1945
                    Width =1206
                    Height =315
                    ColumnOrder =13
                    TabIndex =16
                    Name ="N_KOL2"
                    ControlSource ="N_KOL2"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TOTA_HES.NUMBER FROM TOTA_HES ORDER BY TOTA_HES.NUMBER;"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextFontCharSet =0
                            Left =4747
                            Top =1945
                            Width =945
                            Height =300
                            Name ="Label21"
                            Caption ="حساب كل:"
                        End
                    End
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
                    Left =3396
                    Top =2305
                    Width =1206
                    Height =315
                    ColumnOrder =11
                    TabIndex =18
                    Name ="N_MOIN2"
                    ControlSource ="N_MOIN2"
                    RowSourceType ="Table/View/StoredProc"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextFontCharSet =0
                            Left =4749
                            Top =2305
                            Width =975
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
                    Left =3394
                    Top =3355
                    Width =1206
                    Height =315
                    ColumnOrder =7
                    TabIndex =22
                    Name ="N_KOL3"
                    ControlSource ="N_KOL3"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TOTA_HES.NUMBER FROM TOTA_HES ORDER BY TOTA_HES.NUMBER;"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextFontCharSet =0
                            Left =4747
                            Top =3355
                            Width =945
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
                    Left =3396
                    Top =3715
                    Width =1206
                    Height =315
                    ColumnOrder =5
                    TabIndex =24
                    Name ="N_MOIN3"
                    ControlSource ="N_MOIN3"
                    RowSourceType ="Table/View/StoredProc"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextFontCharSet =0
                            Left =4749
                            Top =3715
                            Width =1050
                            Height =300
                            Name ="Label27"
                            Caption ="حساب معين :"
                        End
                    End
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
                    Left =1077
                    Top =510
                    Width =2226
                    Height =315
                    ColumnOrder =18
                    TabIndex =11
                    Name ="AN_KOL"
                    ControlSource ="N_KOL"
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
                    Left =1079
                    Top =870
                    Width =2226
                    Height =315
                    ColumnOrder =16
                    TabIndex =13
                    Name ="AN_MOIN"
                    ControlSource ="N_MOIN"
                    RowSourceType ="Table/View/StoredProc"
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
                    Left =1019
                    Top =1920
                    Width =2226
                    Height =315
                    ColumnOrder =12
                    TabIndex =17
                    Name ="AN_KOL2"
                    ControlSource ="N_KOL2"
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
                    Left =1021
                    Top =2280
                    Width =2226
                    Height =315
                    ColumnOrder =10
                    TabIndex =19
                    Name ="AN_MOIN2"
                    ControlSource ="N_MOIN2"
                    RowSourceType ="Table/View/StoredProc"
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
                    Left =1019
                    Top =3330
                    Width =2226
                    Height =315
                    ColumnOrder =6
                    TabIndex =23
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
                    Left =1021
                    Top =3690
                    Width =2226
                    Height =315
                    ColumnOrder =4
                    TabIndex =25
                    Name ="AN_MOIN3"
                    ControlSource ="N_MOIN3"
                    RowSourceType ="Table/View/StoredProc"
                    ColumnWidths ="0"
                    AfterUpdate ="[Event Procedure]"
                End
                Begin Label
                    OverlapFlags =215
                    TextFontCharSet =0
                    Left =4469
                    Top =120
                    Width =1290
                    Height =300
                    Name ="Label56"
                    Caption ="واگذار به حساب :"
                End
                Begin Label
                    OverlapFlags =223
                    TextFontCharSet =0
                    Left =4231
                    Top =1575
                    Width =1515
                    Height =300
                    Name ="Label57"
                    Caption ="برگشت به حساب:"
                End
                Begin Label
                    OverlapFlags =215
                    TextFontCharSet =0
                    Left =4454
                    Top =3030
                    Width =1335
                    Height =300
                    Name ="Label58"
                    Caption ="*وصول به حساب:"
                    OnDblClick ="[Event Procedure]"
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =0
                    TextAlign =3
                    Left =8853
                    Top =170
                    Width =1575
                    Height =285
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
                    Left =6059
                    Top =2911
                    Width =3300
                    Height =315
                    ColumnOrder =22
                    TabIndex =6
                    Name ="NAME_TAH"
                    ControlSource ="NAME_TAH"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT PAY_GETD.NAME_TAH FROM PAY_GETD GROUP BY PAY_GETD.NAME_TAH ORDER BY PAY_G"
                        "ETD.NAME_TAH;"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextFontCharSet =0
                            Left =9511
                            Top =2910
                            Width =1350
                            Height =300
                            Name ="Label13"
                            Caption ="نام پرداخت كننده:"
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
                    Left =2400
                    Top =5220
                    Width =1161
                    Height =300
                    ColumnOrder =29
                    TabIndex =30
                    BackColor =8421376
                    ForeColor =16777215
                    Name ="TAG"
                    ControlSource ="TAG"
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =3570
                            Top =5220
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
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =8525
                    Top =3667
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
                            Left =9469
                            Top =3664
                            Width =870
                            Height =300
                            Name ="Label64"
                            Caption ="رديف دفتر:"
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
                    Left =3394
                    Top =1255
                    Width =1206
                    Height =315
                    ColumnOrder =15
                    TabIndex =14
                    Name ="N_TAF"
                    ControlSource ="N_TAF"
                    RowSourceType ="Table/View/StoredProc"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextFontCharSet =0
                            Left =4741
                            Top =1250
                            Width =975
                            Height =300
                            Name ="Label67"
                            Caption ="تفصيلي:"
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
                    ColumnCount =2
                    Left =1019
                    Top =1230
                    Width =2226
                    Height =315
                    ColumnOrder =14
                    TabIndex =15
                    Name ="AN_TAF"
                    ControlSource ="N_TAF"
                    RowSourceType ="Table/View/StoredProc"
                    ColumnWidths ="0"
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
                    Left =3394
                    Top =2680
                    Width =1206
                    Height =315
                    ColumnOrder =9
                    TabIndex =20
                    Name ="N_TAF2"
                    ControlSource ="N_TAF2"
                    RowSourceType ="Table/View/StoredProc"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextFontCharSet =0
                            Left =4747
                            Top =2680
                            Width =975
                            Height =300
                            Name ="Label70"
                            Caption ="تفصيلي:"
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
                    ColumnCount =2
                    Left =1019
                    Top =2655
                    Width =2226
                    Height =315
                    ColumnOrder =8
                    TabIndex =21
                    Name ="AN_TAF2"
                    ControlSource ="N_TAF2"
                    RowSourceType ="Table/View/StoredProc"
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
                    Left =3394
                    Top =4075
                    Width =1206
                    Height =315
                    ColumnOrder =3
                    TabIndex =26
                    Name ="N_TAF3"
                    ControlSource ="N_TAF3"
                    RowSourceType ="Table/View/StoredProc"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextFontCharSet =0
                            Left =4747
                            Top =4075
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
                    OldBorderStyle =0
                    OverlapFlags =215
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =1019
                    Top =4050
                    Width =2226
                    Height =315
                    ColumnOrder =2
                    TabIndex =27
                    Name ="AN_TAF3"
                    ControlSource ="N_TAF3"
                    RowSourceType ="Table/View/StoredProc"
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
                    Left =8525
                    Top =4105
                    Width =831
                    Height =315
                    TabIndex =9
                    Name ="VAZ"
                    ControlSource ="VAZ"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextFontCharSet =0
                            Left =9469
                            Top =4102
                            Width =870
                            Height =300
                            Name ="Label76"
                            Caption ="صندوق:"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =247
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =6245
                    Top =4543
                    Width =3111
                    Height =315
                    TabIndex =31
                    Name ="SAYADI"
                    ControlSource ="SAYADI"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextFontCharSet =0
                            Left =9465
                            Top =4545
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
' See "PAY_GETD.cls"
