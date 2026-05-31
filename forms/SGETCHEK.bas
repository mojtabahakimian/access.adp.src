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
    Width =8163
    DatasheetFontHeight =10
    ItemSuffix =48
    Left =6105
    Top =1650
    Right =14505
    Bottom =7605
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xeeea06682da3e240
    End
    RecordSource ="PAY_GETD"
    Caption ="چك دريافتي"
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
            Height =737
            BackColor =8421376
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =93
                    TextFontFamily =2
                    Left =3477
                    Top =15
                    Width =1635
                    Height =630
                    FontSize =18
                    FontWeight =700
                    ForeColor =16777215
                    Name ="Label42"
                    Caption ="دريافت چك"
                    FontName ="Titr"
                End
                Begin Label
                    OverlapFlags =215
                    TextFontFamily =2
                    Left =3477
                    Width =1635
                    Height =630
                    FontSize =18
                    FontWeight =700
                    Name ="Label43"
                    Caption ="دريافت چك"
                    FontName ="Titr"
                End
            End
        End
        Begin Section
            Height =5944
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4203
                    Top =465
                    Width =2241
                    Height =315
                    TabIndex =1
                    Name ="N_SERI"
                    ControlSource ="N_SERI"
                    Format ="Fixed"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =6565
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
                    Left =4931
                    Top =1970
                    Width =1506
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
                            Left =6565
                            Top =1965
                            Width =1095
                            Height =300
                            Name ="Label5"
                            Caption ="تاريخ دريافت:"
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
                    Left =4931
                    Top =2325
                    Width =1506
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
                            Left =6565
                            Top =2325
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
                    Left =4370
                    Top =2691
                    Width =2061
                    Height =315
                    TabIndex =6
                    Name ="MABL"
                    ControlSource ="MABL"
                    Format ="#,000\" ريال\";-#,000\" ريال\""
                    DefaultValue ="=[Forms]![DEED_HEAD]![Child14].[Form]![BED]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =6565
                            Top =2685
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
                    Left =4010
                    Top =3405
                    Width =2421
                    Height =315
                    TabIndex =8
                    Name ="N_HESAB"
                    ControlSource ="N_HESAB"
                    Format ="Fixed"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =6565
                            Top =3409
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
                    Left =3935
                    Top =825
                    Width =2511
                    Height =315
                    TabIndex =2
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
                            TextFontCharSet =0
                            Left =6565
                            Top =830
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
                    Left =3493
                    Top =1185
                    Width =2955
                    Height =315
                    TabIndex =3
                    Name ="SHOBEH"
                    ControlSource ="SHOBEH"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT PAY_GETD.SHOBEH FROM PAY_GETD GROUP BY PAY_GETD.SHOBEH ORDER BY PAY_GETD."
                        "SHOBEH;"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =6565
                            Top =1189
                            Width =945
                            Height =300
                            Name ="Label9"
                            Caption ="نام شعبه:"
                        End
                    End
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3131
                    Top =3046
                    Width =3300
                    Height =315
                    TabIndex =7
                    Name ="NAME_TAH"
                    ControlSource ="NAME_TAH"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT PAY_GETD.NAME_TAH FROM PAY_GETD GROUP BY PAY_GETD.NAME_TAH ORDER BY PAY_G"
                        "ETD.NAME_TAH;"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =6565
                            Top =3045
                            Width =1350
                            Height =300
                            Name ="Label13"
                            Caption ="نام پرداخت كننده:"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5380
                    Top =108
                    Width =1071
                    Height =315
                    Name ="RADIF"
                    ControlSource ="RADIF"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =6565
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
                    Left =2220
                    Top =5550
                    Width =1569
                    Height =394
                    FontWeight =700
                    TabIndex =16
                    Name ="Command44"
                    Caption ="&تصويب و خروج"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    Cancel = NotDefault
                    OverlapFlags =85
                    AccessKey =206
                    TextFontCharSet =0
                    Left =3810
                    Top =5550
                    Width =1569
                    Height =394
                    FontWeight =700
                    TabIndex =17
                    Name ="Command45"
                    Caption ="&خـــــروج"
                    OnClick ="[Event Procedure]"
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =722
                    Top =1755
                    Width =1176
                    Height =315
                    TabIndex =10
                    Name ="N_KOL"
                    ControlSource ="N_KOL"
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =2069
                            Top =1760
                            Width =1050
                            Height =300
                            Name ="Label17"
                            Caption =" به حساب كل:"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =87
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =722
                    Top =2076
                    Width =1176
                    Height =315
                    TabIndex =11
                    Name ="N_MOIN"
                    ControlSource ="N_MOIN"
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =2069
                            Top =2071
                            Width =975
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
                    Left =722
                    Top =2449
                    Width =1176
                    Height =315
                    TabIndex =12
                    Name ="N_TAF"
                    ControlSource ="N_TAF"
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =2039
                            Top =2444
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
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =1481
                    Top =3780
                    Width =4956
                    Height =330
                    TabIndex =9
                    Name ="HES"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT RTRIM(CAST(TOTA_HES.NUMBER AS nvarchar)) + '-' + RTRIM(CAST(DETA_HES.NUMB"
                        "ER AS nvarchar)) + '-' + RTRIM(CAST(TDETA_HES.TNUMBER AS nvarchar)) AS hes, RTRI"
                        "M(CAST(TOTA_HES.NUMBER AS nvarchar)) + '-' + RTRIM(CAST(DETA_HES.NUMBER AS nvarc"
                        "har)) + '-' + RTRIM(CAST(TDETA_HES.TNUMBER AS nvarchar)) + ' - ' + TDETA_HES.NAM"
                        "E AS Expr1, TDETA_HES.N_KOL, SAZMAN.BANKHA FROM TOTA_HES INNER JOIN DETA_HES INN"
                        "ER JOIN TDETA_HES ON DETA_HES.NUMBER = TDETA_HES.NUMBER AND DETA_HES.N_KOL = TDE"
                        "TA_HES.N_KOL ON TOTA_HES.NUMBER = DETA_HES.N_KOL INNER JOIN SAZMAN ON TDETA_HES."
                        "N_KOL = SAZMAN.BANKHA AND TDETA_HES.N_KOL = SAZMAN.BANKHA"
                    ColumnWidths ="0"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =6565
                            Top =3780
                            Width =915
                            Height =285
                            Name ="Label4"
                            Caption ="به حساب"
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
                    ColumnCount =3
                    Left =2204
                    Top =4155
                    Width =4221
                    Height =315
                    TabIndex =13
                    BackColor =13434828
                    ForeColor =8388608
                    Name ="CUST_NO"
                    ControlSource ="CUST_NO"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT RTRIM(CAST(N_KOL AS nvarchar)) + '-' + RTRIM(CAST(NUMBER AS nvarchar)) + "
                        "'-' + RTRIM(CAST(TNUMBER AS nvarchar)) AS hes, NAME AS nam, RTRIM(CAST(N_KOL AS "
                        "nvarchar)) + '-' + RTRIM(CAST(NUMBER AS nvarchar)) + '-' + RTRIM(CAST(TNUMBER AS"
                        " nvarchar)) AS Expr1 FROM TDETA_HES"
                    ColumnWidths ="0;2835;567"
                    Tag ="115-1-1241"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =6565
                            Top =4155
                            Width =1245
                            Height =240
                            ForeColor =8388608
                            Name ="Label46"
                            Caption =" حساب مشتري:"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    AutoExpand = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =464
                    Top =4155
                    Width =1641
                    Height =315
                    TabIndex =14
                    BackColor =13434828
                    ForeColor =8388608
                    Name ="CUST_NO2"
                    ControlSource ="CUST_NO"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT RTRIM(CAST(N_KOL AS nvarchar)) + '-' + RTRIM(CAST(NUMBER AS nvarchar)) + "
                        "'-' + RTRIM(CAST(TNUMBER AS nvarchar)) AS Expr1 FROM TDETA_HES"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    Enabled = NotDefault
                    RowSourceTypeInt =1
                    OldBorderStyle =0
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =465
                    Top =4875
                    Width =1641
                    Height =300
                    TabIndex =18
                    Name ="KIND"
                    ControlSource ="KIND"
                    RowSourceType ="Value List"
                    RowSource ="1;\"تجاري\";0;\"غير تجاري\""
                    ColumnWidths ="0"
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
                    Left =2204
                    Top =4890
                    Width =4221
                    Height =315
                    TabIndex =19
                    Name ="SANDUGH"
                    ControlSource ="SANDUGH"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TNUMBER, NAME FROM TDETA_HES WHERE (N_KOL = 113) AND (NUMBER = 1)"
                    ColumnWidths ="0"
                    DefaultValue ="1"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =6565
                            Top =4890
                            Width =1080
                            Height =345
                            Name ="Label50"
                            Caption ="موقعيت چك:"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    Visible = NotDefault
                    RowSourceTypeInt =1
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Width =2211
                    Height =315
                    TabIndex =20
                    Name ="VAZ"
                    ControlSource ="VAZ"
                    RowSourceType ="Value List"
                    RowSource ="1;\"نزد صندوق\";2;\"نزد بانك\";3;\"وصول شده\";4;\"واگذار شده\";5;\"برگشت شده\";6"
                        ";\"مسترد شده\""
                    ColumnWidths ="0"
                    DefaultValue ="1"
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =2342
                            Width =1005
                            Height =315
                            Name ="Label52"
                            Caption ="وضعيت چك"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =210
                    Top =3084
                    TabIndex =21
                    Name ="ANBAR"
                    ControlSource ="ANBAR"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =3222
                            Top =3084
                            Width =390
                            Height =240
                            Name ="Label47"
                            Caption ="انبار:"
                        End
                    End
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4890
                    Top =1575
                    Width =1560
                    Height =315
                    TabIndex =22
                    Name ="LIST_NO"
                    ControlSource ="LIST_NO"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT LIST_NO FROM PAY_GETD GROUP BY LIST_NO"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =6565
                            Top =1579
                            Width =960
                            Height =300
                            Name ="Label76"
                            Caption ="كد شعبه:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3315
                    Top =4530
                    Width =3111
                    Height =315
                    TabIndex =15
                    Name ="SAYADI"
                    ControlSource ="SAYADI"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =6565
                            Top =4532
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
' See "SGETCHEK.cls"
