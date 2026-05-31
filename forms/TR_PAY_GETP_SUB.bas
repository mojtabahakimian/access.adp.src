Version =20
VersionRequired =20
Begin Form
    OrderByOn = NotDefault
    DefaultView =2
    AllowUpdating =4
    ViewsAllowed =2
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =5669
    RowHeight =315
    ItemSuffix =3
    Left =510
    Top =3840
    Right =13665
    Bottom =8850
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x43d1dac0f0aee340
    End
    RecordSource ="TR_PAY_GETP"
    DatasheetFontName ="Tahoma"
    RecordSourceQualifier ="dbo"
    Begin
        Begin Label
            BackStyle =0
            TextFontCharSet =178
            FontName ="Tahoma"
        End
        Begin TextBox
            SpecialEffect =2
            OldBorderStyle =0
            TextFontCharSet =178
            Width =1701
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
            Height =7156
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1023
                    Top =113
                    ColumnOrder =1
                    Name ="N_SERI"
                    ControlSource ="N_SERI"
                    Format ="Fixed"
                    ValidationRule ="Not Is Null"
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
                    OverlapFlags =93
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1023
                    Top =793
                    ColumnOrder =0
                    TabIndex =1
                    Name ="DATE"
                    ControlSource ="DATE"
                    DefaultValue ="=FARSIDATE(Date())"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =93
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
                    ColumnOrder =4
                    TabIndex =2
                    Name ="DATE_S"
                    ControlSource ="DATE_S"
                    DefaultValue ="Null"
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
                    ColumnOrder =5
                    TabIndex =4
                    Name ="MABL"
                    ControlSource ="MABL"
                    Format ="Fixed"
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
                    ColumnWidth =2235
                    ColumnOrder =6
                    TabIndex =5
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
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1023
                    Top =2494
                    ColumnWidth =1050
                    ColumnOrder =7
                    TabIndex =6
                    Name ="N_HESAB"
                    ControlSource ="N_HESAB"
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
                    ColumnOrder =8
                    TabIndex =7
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
                    ColumnOrder =9
                    TabIndex =8
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
                    ColumnOrder =10
                    TabIndex =9
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
                    ColumnOrder =11
                    TabIndex =10
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
                    ColumnOrder =12
                    TabIndex =11
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
                    Top =4875
                    ColumnOrder =13
                    TabIndex =12
                    Name ="VAZ"
                    ControlSource ="VAZ"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4010
                            Top =4875
                            Width =585
                            Height =240
                            Name ="Label29"
                            Caption ="وضعيت"
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
                    ColumnOrder =14
                    TabIndex =13
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
                    ColumnWidth =810
                    ColumnOrder =17
                    TabIndex =14
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
                    ColumnOrder =18
                    TabIndex =15
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
                    ColumnOrder =19
                    TabIndex =16
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
                    ColumnWidth =480
                    ColumnOrder =20
                    TabIndex =17
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
                    ColumnOrder =21
                    TabIndex =18
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
                    LimitToList = NotDefault
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
                    ColumnOrder =2
                    TabIndex =19
                    Name ="BANK"
                    ControlSource ="BANK"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TCOD_BANKS.CODE, TCOD_BANKS.NAMES FROM TCOD_BANKS ORDER BY TCOD_BANKS.NAM"
                        "ES;"
                    ColumnWidths ="0"
                    ValidationRule ="Not Is Null"
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
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =1023
                    Top =3174
                    ColumnWidth =1755
                    ColumnOrder =15
                    TabIndex =20
                    Name ="N_MOIN"
                    ControlSource ="N_MOIN"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT     DETA_HES.NUMBER, DETA_HES.NAME FROM DETA_HES WHERE     (((DETA_HES.N_"
                        "KOL) =122)) GROUP BY DETA_HES.NUMBER, DETA_HES.NAME ORDER BY DETA_HES.NAME"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3645
                            Top =3180
                            Width =1230
                            Height =240
                            Name ="Label19"
                            Caption ="پرداخت ازحساب  معين"
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
                    Left =1023
                    Top =3540
                    ColumnWidth =3225
                    ColumnOrder =16
                    TabIndex =21
                    Name ="N_TAF"
                    ControlSource ="N_TAF"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TDETA_HES.TNUMBER, TDETA_HES.NAME FROM TDETA_HES WHERE (((TDETA_HES.NUMBE"
                        "R) =) AND ((TDETA_HES.N_KOL) =122))GROUP BY TDETA_HES.TNUMBER, TDETA_HES.NAME OR"
                        "DER BY TDETA_HES.NAME"
                    ColumnWidths ="0"
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
                    OverlapFlags =93
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1014
                    Top =1474
                    Width =1710
                    ColumnOrder =3
                    TabIndex =3
                    Name ="SHOBEH"
                    ControlSource ="SHOBEH"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT PAY_GETP.SHOBEH FROM PAY_GETP GROUP BY PAY_GETP.SHOBEH ORDER BY PAY_GETP."
                        "SHOBEH;"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =3875
                            Top =1474
                            Width =720
                            Height =240
                            Name ="Label9"
                            Caption ="نام شعبه"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =247
                    IMESentenceMode =3
                    Left =1417
                    Top =793
                    TabIndex =22
                    Name ="UP_DATE"
                    ControlSource ="UP_DATE"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =4009
                            Top =793
                            Width =810
                            Height =240
                            Name ="Label0"
                            Caption ="UP_DATE:"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =247
                    IMESentenceMode =3
                    Left =1813
                    Top =1530
                    ColumnWidth =1980
                    TabIndex =23
                    Name ="UP_TIME"
                    ControlSource ="UP_TIME"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =4435
                            Top =1530
                            Width =780
                            Height =240
                            Name ="Label2"
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
' See "TR_PAY_GETP_SUB.cls"
