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
    Width =4920
    RowHeight =240
    ItemSuffix =44
    Left =465
    Top =3840
    Right =11460
    Bottom =6570
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xf6364be61b62e240
    End
    RecordSource ="PAY_GETP"
    Caption ="PAY_GETD"
    OnCurrent ="[Event Procedure]"
    BeforeInsert ="[Event Procedure]"
    OnDelete ="[Event Procedure]"
    AfterUpdate ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Tahoma"
    OnKeyPress ="[Event Procedure]"
    OnError ="[Event Procedure]"
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
            Height =7276
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1023
                    Top =113
                    ColumnWidth =1095
                    ColumnOrder =1
                    Name ="N_SERI"
                    ControlSource ="N_SERI"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3590
                            Top =113
                            Width =1005
                            Height =240
                            Name ="Label1"
                            Caption ="شماره سريال"
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
                    Left =1023
                    Top =793
                    ColumnWidth =1110
                    ColumnOrder =0
                    TabIndex =2
                    Name ="DATE"
                    ControlSource ="DATE"
                    DefaultValue ="=FARSIDATE(Date())"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3710
                            Top =793
                            Width =885
                            Height =240
                            Name ="Label5"
                            Caption ="تاريخ پرداخت"
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
                    Left =1023
                    Top =1133
                    ColumnWidth =1125
                    ColumnOrder =4
                    TabIndex =3
                    Name ="DATE_S"
                    ControlSource ="DATE_S"
                    BeforeUpdate ="[Event Procedure]"
                    OnExit ="[Event Procedure]"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3545
                            Top =1133
                            Width =1050
                            Height =240
                            Name ="Label7"
                            Caption ="تاريخ سررسيد"
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
                    Left =1023
                    Top =1814
                    ColumnWidth =1800
                    ColumnOrder =5
                    TabIndex =5
                    Name ="MABL"
                    ControlSource ="MABL"
                    Format ="Fixed"
                    OnKeyPress ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4220
                            Top =1814
                            Width =375
                            Height =240
                            Name ="Label11"
                            Caption ="مبلغ"
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
                    Top =2154
                    Width =2580
                    ColumnWidth =2820
                    ColumnOrder =6
                    TabIndex =6
                    Name ="NAME_TAH"
                    ControlSource ="NAME_TAH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3420
                            Top =2160
                            Width =1215
                            Height =240
                            Name ="Label13"
                            Caption ="نام دريافت كننده"
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
                    Left =1023
                    Top =2494
                    ColumnWidth =1050
                    ColumnOrder =7
                    TabIndex =7
                    Name ="N_HESAB"
                    ControlSource ="N_HESAB"
                    Format ="Fixed"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3560
                            Top =2494
                            Width =1035
                            Height =240
                            Name ="Label15"
                            Caption ="جاري چك"
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
                    Left =1023
                    Top =2834
                    ColumnWidth =1417
                    ColumnOrder =10
                    TabIndex =8
                    Name ="N_KOL"
                    ControlSource ="N_KOL"
                    DefaultValue ="=[Forms]![Baseknow]![BANKHA]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3815
                            Top =2834
                            Width =780
                            Height =240
                            Name ="Label17"
                            Caption ="حساب كل"
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
                    Left =1023
                    Top =3514
                    ColumnOrder =11
                    TabIndex =10
                    Name ="N_KOL2"
                    ControlSource ="N_KOL2"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =3725
                            Top =3514
                            Width =870
                            Height =240
                            Name ="Label21"
                            Caption ="حساب كل2"
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
                    Left =1023
                    Top =3855
                    ColumnOrder =12
                    TabIndex =11
                    Name ="N_MOIN2"
                    ControlSource ="N_MOIN2"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3515
                            Top =3855
                            Width =1080
                            Height =240
                            Name ="Label23"
                            Caption ="حساب معين 2"
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
                    Left =1023
                    Top =4195
                    ColumnOrder =13
                    TabIndex =12
                    Name ="N_KOL3"
                    ControlSource ="N_KOL3"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3680
                            Top =4195
                            Width =915
                            Height =240
                            Name ="Label25"
                            Caption ="حساب كل 3"
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
                    Left =1023
                    Top =4535
                    ColumnWidth =1140
                    ColumnOrder =14
                    TabIndex =13
                    Name ="N_MOIN3"
                    ControlSource ="N_MOIN3"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3515
                            Top =4535
                            Width =1080
                            Height =240
                            Name ="Label27"
                            Caption ="حساب معين 3"
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
                    Left =1023
                    Top =5215
                    ColumnWidth =885
                    ColumnOrder =15
                    TabIndex =14
                    Name ="N_S"
                    ControlSource ="N_S"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3710
                            Top =5215
                            Width =885
                            Height =240
                            Name ="Label31"
                            Caption ="شماره سند"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1023
                    Top =5555
                    ColumnWidth =885
                    ColumnOrder =16
                    TabIndex =15
                    Name ="RADIF"
                    ControlSource ="RADIF"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =4185
                            Top =5550
                            Width =675
                            Height =405
                            Name ="Label33"
                            Caption ="رديف دفتر"
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
                    Left =1023
                    Top =5896
                    ColumnWidth =1417
                    ColumnOrder =17
                    TabIndex =16
                    Name ="NUMBER"
                    ControlSource ="NUMBER"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =3635
                            Top =5896
                            Width =960
                            Height =240
                            Name ="Label35"
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
                    Left =1023
                    Top =6236
                    ColumnWidth =1417
                    ColumnOrder =18
                    TabIndex =17
                    Name ="TAG"
                    ControlSource ="TAG"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3965
                            Top =6236
                            Width =630
                            Height =240
                            Name ="Label37"
                            Caption ="برچسب"
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
                    Left =1023
                    Top =6576
                    ColumnWidth =1417
                    ColumnOrder =19
                    TabIndex =18
                    Name ="ANBAR"
                    ControlSource ="ANBAR"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4265
                            Top =6576
                            Width =330
                            Height =240
                            Name ="Label39"
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
                    Left =1023
                    Top =6916
                    ColumnWidth =1155
                    ColumnOrder =20
                    TabIndex =19
                    Name ="CUST_NO"
                    ControlSource ="CUST_NO"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3470
                            Top =6916
                            Width =1125
                            Height =240
                            Name ="Label41"
                            Caption ="شماره مشتري"
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
                    Left =1023
                    Top =453
                    Width =786
                    ColumnWidth =930
                    ColumnOrder =2
                    TabIndex =1
                    Name ="BANK"
                    ControlSource ="BANK"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TCOD_BANKS.CODE, TCOD_BANKS.NAMES FROM TCOD_BANKS ORDER BY TCOD_BANKS.NAM"
                        "ES;"
                    ColumnWidths ="0"
                    ValidationRule ="Not Is Null"
                    AfterUpdate ="[Event Procedure]"
                    OnNotInList ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4220
                            Top =453
                            Width =375
                            Height =240
                            Name ="Label3"
                            Caption ="بانك"
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
                    Left =798
                    Top =3174
                    Width =1926
                    ColumnWidth =1755
                    ColumnOrder =8
                    TabIndex =9
                    Name ="N_MOIN"
                    ControlSource ="N_MOIN"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT DETA_HES.NUMBER, DETA_HES.NAME FROM DETA_HES WHERE (((DETA_HES.N_KOL)=[Fo"
                        "rms]![Baseknow]![BANKHA])) GROUP BY DETA_HES.NUMBER, DETA_HES.NAME ORDER BY DETA"
                        "_HES.NAME;"
                    ColumnWidths ="0"
                    OnNotInList ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =223
                            Left =3645
                            Top =3180
                            Width =1275
                            Height =405
                            Name ="Label19"
                            Caption ="پرداخت از حساب معين"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =247
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =798
                    Top =3540
                    Width =1926
                    ColumnWidth =1860
                    ColumnOrder =9
                    TabIndex =20
                    Name ="N_TAF"
                    ControlSource ="N_TAF"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TDETA_HES.TNUMBER, TDETA_HES.NAME FROM TDETA_HES WHERE (((TDETA_HES.NUMBE"
                        "R)=[Forms]![HEAD_LST_FROOSH_BACK]![PAY_GETD_BACK_SUB].[Form]![N_MOIN]) AND ((TDE"
                        "TA_HES.N_KOL)=[Forms]![Baseknow]![BANKHA])) GROUP BY TDETA_HES.TNUMBER, TDETA_HE"
                        "S.NAME ORDER BY TDETA_HES.NAME;"
                    ColumnWidths ="0"
                    OnNotInList ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =3645
                            Top =3546
                            Width =1230
                            Height =240
                            Name ="Label45"
                            Caption ="پرداخت ازحساب تفصيلي"
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
                    Left =1014
                    Top =1474
                    Width =1710
                    ColumnWidth =1800
                    ColumnOrder =3
                    TabIndex =4
                    Name ="SHOBEH"
                    ControlSource ="SHOBEH"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT SHOBEH FROM PAY_GETD GROUP BY SHOBEH ORDER BY SHOBEH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3875
                            Top =1474
                            Width =720
                            Height =240
                            Name ="Label9"
                            Caption ="نام شعبه"
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
                    Name ="SCMAB"
                    ControlSource ="=Sum([MABL])"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3197
                            Width =375
                            Height =240
                            Name ="Label43"
                            Caption ="مبلغ"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "PAY_GETD_BACK_SUB.cls"
