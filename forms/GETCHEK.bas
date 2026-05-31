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
    Width =7880
    DatasheetFontHeight =10
    ItemSuffix =55
    Left =3690
    Top =1710
    Right =11805
    Bottom =7470
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xef204f601963e240
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
            Height =540
            BackColor =8421376
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =93
                    TextFontFamily =10
                    Left =3135
                    Top =15
                    Width =1635
                    Height =525
                    FontSize =16
                    FontWeight =700
                    ForeColor =16777215
                    Name ="Label42"
                    Caption ="دريافت چك"
                    FontName ="Sina"
                End
                Begin Label
                    OverlapFlags =215
                    TextFontFamily =10
                    Left =3135
                    Width =1635
                    Height =525
                    FontSize =16
                    FontWeight =700
                    Name ="Label43"
                    Caption ="دريافت چك"
                    FontName ="Sina"
                End
            End
        End
        Begin Section
            Height =5782
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
                    Left =3861
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
                            Left =6254
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
                    Left =4587
                    Top =1930
                    Width =1506
                    Height =315
                    TabIndex =5
                    Name ="DATE"
                    ControlSource ="DATE"
                    BeforeUpdate ="[Event Procedure]"
                    DefaultValue ="=[Forms]![PGET_HED]![DATE]"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =6245
                            Top =1925
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
                    Left =4587
                    Top =2285
                    Width =1506
                    Height =315
                    TabIndex =6
                    Name ="DATE_S"
                    ControlSource ="DATE_S"
                    BeforeUpdate ="[Event Procedure]"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =6245
                            Top =2285
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
                    Left =4026
                    Top =2651
                    Width =2061
                    Height =315
                    TabIndex =7
                    Name ="MABL"
                    ControlSource ="MABL"
                    Format ="#,##0.00\" ريال \";-#,##0.00\" ريال \""
                    BeforeUpdate ="[Event Procedure]"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="=[Forms]![PGET_HED]![PGET_LST_SUB].[Form]![MABL]"
                    OnKeyPress ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =6239
                            Top =2645
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
                    Left =3666
                    Top =3365
                    Width =2421
                    Height =315
                    TabIndex =9
                    Name ="N_HESAB"
                    ControlSource ="N_HESAB"
                    Format ="Fixed"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =6239
                            Top =3369
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
                    Left =3593
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
                            Left =6256
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
                    Left =3151
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
                            Left =6258
                            Top =1189
                            Width =900
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
                    Left =2787
                    Top =3006
                    Width =3300
                    Height =315
                    TabIndex =8
                    Name ="NAME_TAH"
                    ControlSource ="NAME_TAH"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT PAY_GETD.NAME_TAH FROM PAY_GETD GROUP BY PAY_GETD.NAME_TAH ORDER BY PAY_G"
                        "ETD.NAME_TAH;"
                    DefaultValue ="=GETTAFNAME([Forms]![PGET_HED]![PGET_LST_SUB].[Form]![fHES])"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =6239
                            Top =3005
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
                    Left =5038
                    Top =108
                    Width =1071
                    Height =315
                    Name ="RADIF"
                    ControlSource ="RADIF"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =6222
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
                    Left =1815
                    Top =5220
                    Width =1569
                    Height =394
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
                    Left =3405
                    Top =5220
                    Width =1569
                    Height =394
                    TabIndex =17
                    Name ="Command45"
                    Caption ="&خـــــروج"
                    OnClick ="[Event Procedure]"
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
                    Left =1152
                    Top =3755
                    Width =4956
                    Height =315
                    TabIndex =10
                    Name ="HES"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT RTRIM(CAST(TOTA_HES.NUMBER AS nvarchar)) + '-' + RTRIM(CAST(DETA_HES.NUMB"
                        "ER AS nvarchar)) + '-' + RTRIM(CAST(TDETA_HES.TNUMBER AS nvarchar)) AS hes, TDET"
                        "A_HES.NAME FROM TOTA_HES INNER JOIN DETA_HES INNER JOIN TDETA_HES ON DETA_HES.NU"
                        "MBER = TDETA_HES.NUMBER AND DETA_HES.N_KOL = TDETA_HES.N_KOL ON TOTA_HES.NUMBER "
                        "= DETA_HES.N_KOL"
                    ColumnWidths ="1419;2835"
                    BeforeUpdate ="[Event Procedure]"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =6239
                            Top =3755
                            Width =1005
                            Height =315
                            Name ="Label4"
                            Caption ="به حساب:"
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
                    Left =1172
                    Top =1920
                    Width =1176
                    Height =315
                    TabIndex =13
                    Name ="N_KOL"
                    ControlSource ="N_KOL"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =2519
                            Top =1925
                            Width =1050
                            Height =300
                            Name ="Label17"
                            Caption =" به حساب كل:"
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
                    Left =1172
                    Top =2260
                    Width =1176
                    Height =315
                    TabIndex =14
                    Name ="N_MOIN"
                    ControlSource ="N_MOIN"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =2519
                            Top =2255
                            Width =1095
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
                    Left =1172
                    Top =2618
                    Width =1176
                    Height =315
                    TabIndex =15
                    Name ="N_TAF"
                    ControlSource ="N_TAF"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =2489
                            Top =2613
                            Width =975
                            Height =300
                            Name ="Label67"
                            Caption ="تفصيلي:"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    Visible = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =56
                    Top =56
                    Width =3231
                    Height =315
                    TabIndex =12
                    Name ="VAZ"
                    ControlSource ="VAZ"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TNUMBER, NAME FROM TDETA_HES WHERE (N_KOL = 113) AND (NUMBER = 1)"
                    ColumnWidths ="0"
                    DefaultValue ="1"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =3418
                            Top =56
                            Width =1005
                            Height =315
                            Name ="Label50"
                            Caption ="نزد صندوق:"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =566
                    Top =453
                    TabIndex =18
                    Name ="CUST_NO"
                    ControlSource ="CUST_NO"
                End
                Begin ComboBox
                    Enabled = NotDefault
                    RowSourceTypeInt =1
                    OldBorderStyle =0
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =1152
                    Top =4175
                    Width =1581
                    Height =300
                    TabIndex =19
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
                    Left =2862
                    Top =4145
                    Width =3231
                    Height =315
                    TabIndex =20
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
                            Left =6224
                            Top =4145
                            Width =1005
                            Height =315
                            Name ="Label53"
                            Caption ="موقعيت چك:"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =435
                    Top =3384
                    TabIndex =21
                    Name ="ANBAR"
                    ControlSource ="ANBAR"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2772
                            Top =3384
                            Width =345
                            Height =240
                            Name ="Label54"
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
                    Left =3975
                    Top =1560
                    Width =2115
                    Height =315
                    TabIndex =4
                    Name ="LIST_NO"
                    ControlSource ="LIST_NO"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT LIST_NO FROM PAY_GETD GROUP BY LIST_NO"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =6242
                            Top =1564
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
                    Left =3000
                    Top =4590
                    Width =3111
                    Height =315
                    TabIndex =11
                    Name ="SAYADI"
                    ControlSource ="SAYADI"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =6220
                            Top =4592
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
' See "GETCHEK.cls"
