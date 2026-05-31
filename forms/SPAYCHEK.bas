Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    DefaultView =0
    AllowUpdating =4
    ScrollBars =0
    ViewsAllowed =1
    TabularCharSet =178
    TabularFamily =42
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    Cycle =1
    GridY =10
    Width =8617
    DatasheetFontHeight =10
    ItemSuffix =49
    Left =3855
    Top =1710
    Right =11655
    Bottom =7485
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xef503b9ed263e240
    End
    RecordSource ="PAY_GETP"
    Caption ="چك  پرداختي"
    BeforeUpdate ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial (Arabic)"
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
            Height =645
            BackColor =8421376
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =93
                    TextFontFamily =2
                    Left =3715
                    Top =15
                    Width =1785
                    Height =630
                    FontSize =18
                    FontWeight =700
                    ForeColor =16777215
                    Name ="Label42"
                    Caption ="پرداخت چك"
                    FontName ="Titr"
                End
                Begin Label
                    OverlapFlags =215
                    TextFontFamily =2
                    Left =3715
                    Width =1785
                    Height =630
                    FontSize =18
                    FontWeight =700
                    Name ="Label43"
                    Caption ="پرداخت چك"
                    FontName ="Titr"
                End
            End
        End
        Begin Section
            Height =4894
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5618
                    Top =108
                    Width =1071
                    Height =315
                    Name ="RADIF"
                    ControlSource ="RADIF"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =6862
                            Top =105
                            Width =870
                            Height =300
                            Name ="Label64"
                            Caption ="رديف دفتر:"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =202
                    TextFontCharSet =0
                    Left =2449
                    Top =4425
                    Width =1569
                    Height =394
                    FontWeight =700
                    TabIndex =14
                    Name ="Command44"
                    Caption ="&تصويب و خروج"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    Cancel = NotDefault
                    OverlapFlags =85
                    AccessKey =206
                    TextFontCharSet =0
                    Left =4039
                    Top =4425
                    Width =1569
                    Height =394
                    FontWeight =700
                    TabIndex =15
                    Name ="Command45"
                    Caption ="&خـــــروج"
                    OnClick ="[Event Procedure]"
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4421
                    Top =465
                    Width =2271
                    Height =315
                    TabIndex =1
                    Name ="N_SERI"
                    ControlSource ="N_SERI"
                    Format ="Fixed"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =6844
                            Top =465
                            Width =1080
                            Height =300
                            Name ="Label1"
                            Caption ="شماره سريال:"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5170
                    Top =1535
                    Width =1521
                    Height =315
                    TabIndex =4
                    Name ="DATE"
                    ControlSource ="DATE"
                    BeforeUpdate ="[Event Procedure]"
                    DefaultValue ="=[Forms]![DEED_HEAD]![DATE_S]"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =6829
                            Top =1530
                            Width =1110
                            Height =300
                            Name ="Label5"
                            Caption ="تاريخ پرداخت:"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5170
                    Top =1890
                    Width =1521
                    Height =315
                    TabIndex =5
                    Name ="DATE_S"
                    ControlSource ="DATE_S"
                    BeforeUpdate ="[Event Procedure]"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =6828
                            Top =1890
                            Width =1110
                            Height =300
                            Name ="Label7"
                            Caption ="تاريخ سررسيد:"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4630
                    Top =2256
                    Width =2061
                    Height =315
                    TabIndex =6
                    Name ="MABL"
                    ControlSource ="MABL"
                    Format ="#,000\" ريال\";-#,000\" ريال\""
                    DefaultValue ="=[Forms]![DEED_HEAD]![Child14].[Form]![BES]"
                    OnKeyPress ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =6843
                            Top =2250
                            Width =480
                            Height =300
                            Name ="Label11"
                            Caption ="مبلغ:"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4311
                    Top =2940
                    Width =2376
                    Height =315
                    TabIndex =8
                    Name ="N_HESAB"
                    ControlSource ="N_HESAB"
                    Format ="Fixed"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =6839
                            Top =2944
                            Width =1035
                            Height =300
                            Name ="Label15"
                            Caption ="جاري چك:"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =4181
                    Top =805
                    Width =2511
                    Height =315
                    TabIndex =2
                    Name ="BANK"
                    ControlSource ="BANK"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TCOD_BANKS.CODE, TCOD_BANKS.NAMES FROM TCOD_BANKS ORDER BY TCOD_BANKS.NAM"
                        "ES;"
                    ColumnWidths ="0"
                    AfterUpdate ="[Event Procedure]"
                    OnNotInList ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =6844
                            Top =810
                            Width =450
                            Height =300
                            Name ="Label3"
                            Caption ="بانك:"
                        End
                    End
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3739
                    Top =1170
                    Width =2955
                    Height =315
                    TabIndex =3
                    Name ="SHOBEH"
                    ControlSource ="SHOBEH"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT PAY_GETP.SHOBEH FROM PAY_GETP GROUP BY PAY_GETP.SHOBEH ORDER BY PAY_GETP."
                        "SHOBEH;"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =6846
                            Top =1174
                            Width =975
                            Height =300
                            Name ="Label9"
                            Caption ="نام شعبه:"
                        End
                    End
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3394
                    Top =2596
                    Width =3300
                    Height =315
                    TabIndex =7
                    Name ="NAME_TAH"
                    ControlSource ="NAME_TAH"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT PAY_GETP.NAME_TAH FROM PAY_GETP GROUP BY PAY_GETP.NAME_TAH ORDER BY PAY_G"
                        "ETP.NAME_TAH;"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =6844
                            Top =2595
                            Width =1380
                            Height =300
                            Name ="Label13"
                            Caption ="نام  دريافت كننده:"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    Enabled = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1199
                    Top =453
                    Width =1176
                    Height =315
                    TabIndex =11
                    Name ="N_KOL"
                    ControlSource ="N_KOL"
                    DefaultValue ="=[Forms]![Baseknow]![BANKHA]"
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =2546
                            Top =458
                            Width =1005
                            Height =300
                            Name ="Label17"
                            Caption ="از حساب كل:"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1199
                    Top =813
                    Width =1176
                    Height =315
                    TabIndex =12
                    Name ="N_MOIN"
                    ControlSource ="N_MOIN"
                    DefaultValue ="=FIRSTM([Forms]![Baseknow]![BANKHA])"
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =2546
                            Top =808
                            Width =1140
                            Height =300
                            Name ="Label19"
                            Caption ="حساب معين:"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1189
                    Top =1183
                    Width =1176
                    Height =315
                    TabIndex =13
                    Name ="N_TAF"
                    ControlSource ="N_TAF"
                    DefaultValue ="=FIRSTT([Forms]![Baseknow]![BANKHA],[N_MOIN])"
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =2506
                            Top =1178
                            Width =1170
                            Height =300
                            Name ="Label67"
                            Caption ="تفصيلي:"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =124
                    Top =3315
                    Width =6561
                    Height =315
                    TabIndex =9
                    Name ="HES1"
                    ControlSource ="HES1"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT hes, hes + N' - ' + NAME AS Expr1 FROM CUST_HESAB WHERE (dbo.GETKOL(hes) "
                        "= 112)"
                    ColumnWidths ="0"
                    BeforeUpdate ="[Event Procedure]"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =6844
                            Top =3315
                            Width =1320
                            Height =285
                            Name ="Label4"
                            Caption ="پرداخت از حساب"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    Enabled = NotDefault
                    RowSourceTypeInt =1
                    OldBorderStyle =0
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =5045
                    Top =4081
                    Width =1581
                    Height =300
                    TabIndex =16
                    Name ="KIND"
                    ControlSource ="KIND"
                    RowSourceType ="Value List"
                    RowSource ="1;\"تجاري\";0;\"غير تجاري\""
                    ColumnWidths ="0"
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3585
                    Top =3690
                    Width =3111
                    Height =315
                    TabIndex =10
                    Name ="SAYADI"
                    ControlSource ="SAYADI"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =6835
                            Top =3692
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
' See "SPAYCHEK.cls"
