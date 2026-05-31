Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    OrderByOn = NotDefault
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
    Width =8503
    DatasheetFontHeight =10
    ItemSuffix =49
    Left =3375
    Top =555
    Right =11880
    Bottom =6015
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
    AllowDatasheetView =0
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
            Height =540
            BackColor =8421376
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =93
                    TextFontFamily =10
                    Left =3545
                    Top =15
                    Width =1635
                    Height =525
                    FontSize =16
                    ForeColor =16777215
                    Name ="Label42"
                    Caption ="پرداخت چك"
                    FontName ="Sina"
                End
                Begin Label
                    OverlapFlags =215
                    TextFontFamily =10
                    Left =3545
                    Width =1635
                    Height =525
                    FontSize =16
                    Name ="Label43"
                    Caption ="پرداخت چك"
                    FontName ="Sina"
                End
            End
        End
        Begin Section
            Height =5215
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5448
                    Top =108
                    Width =1071
                    Height =315
                    Name ="RADIF"
                    ControlSource ="RADIF"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =6692
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
                    Left =2280
                    Top =4695
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
                    Left =3870
                    Top =4695
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
                    Left =4251
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
                            Left =6674
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
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5000
                    Top =1535
                    Width =1521
                    Height =315
                    TabIndex =4
                    Name ="DATE"
                    ControlSource ="DATE"
                    BeforeUpdate ="[Event Procedure]"
                    DefaultValue ="=[Forms]![PGET_HED]![DATE]"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =6659
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
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5000
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
                            Left =6658
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
                    Left =4460
                    Top =2256
                    Width =2061
                    Height =315
                    TabIndex =6
                    Name ="MABL"
                    ControlSource ="MABL"
                    Format ="#,##0.00\" ريال \";-#,##0.00\" ريال \""
                    DefaultValue ="=[Forms]![PGET_HED]![PGET_LST_SUB].[Form]![MABL]"
                    OnKeyPress ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =6673
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
                    Left =4141
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
                            Left =6669
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
                    Left =4011
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
                            Left =6674
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
                    Left =3569
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
                            Left =6676
                            Top =1174
                            Width =870
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
                    Left =3224
                    Top =2596
                    Width =3300
                    Height =315
                    TabIndex =7
                    Name ="NAME_TAH"
                    ControlSource ="NAME_TAH"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT PAY_GETP.NAME_TAH FROM PAY_GETP GROUP BY PAY_GETP.NAME_TAH ORDER BY PAY_G"
                        "ETP.NAME_TAH;"
                    DefaultValue ="=GETTAFNAME([Forms]![PGET_HED]![PGET_LST_SUB].[Form]![THES])"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =6674
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
                    Left =1029
                    Top =1417
                    Width =1176
                    Height =345
                    TabIndex =11
                    Name ="N_KOL"
                    ControlSource ="N_KOL"
                    DefaultValue ="=[Forms]![Baseknow]![BANKHA]"
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =2376
                            Top =1422
                            Width =1005
                            Height =330
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
                    Left =1029
                    Top =1777
                    Width =1176
                    Height =345
                    TabIndex =12
                    Name ="N_MOIN"
                    ControlSource ="N_MOIN"
                    DefaultValue ="=FIRSTM([Forms]![Baseknow]![BANKHA])"
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =2376
                            Top =1772
                            Width =1140
                            Height =330
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
                    Left =1019
                    Top =2147
                    Width =1176
                    Height =345
                    TabIndex =13
                    Name ="N_TAF"
                    ControlSource ="N_TAF"
                    DefaultValue ="=FIRSTT([Forms]![Baseknow]![BANKHA],[N_MOIN])"
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =2336
                            Top =2142
                            Width =975
                            Height =330
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
                    Left =179
                    Top =3345
                    Width =6351
                    Height =315
                    TabIndex =9
                    Name ="HES1"
                    ControlSource ="HES1"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT hes, hes + N' - ' + ISNULL(NAME, N'') AS Expr1 FROM CUST_HESAB WHERE (dbo"
                        ".GETKOL(hes) = 121)"
                    ColumnWidths ="0"
                    BeforeUpdate ="[Event Procedure]"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =6689
                            Top =3345
                            Width =1365
                            Height =300
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
                    Left =4950
                    Top =4200
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
                    Left =3435
                    Top =3750
                    Width =3111
                    Height =315
                    TabIndex =10
                    Name ="SAYADI"
                    ControlSource ="SAYADI"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =6655
                            Top =3752
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
' See "PAYCHEK.cls"
