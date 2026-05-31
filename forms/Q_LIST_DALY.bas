Version =20
VersionRequired =20
Begin Form
    DefaultView =0
    AllowUpdating =4
    ViewsAllowed =2
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =7317
    RowHeight =300
    ItemSuffix =95
    Left =1995
    Top =720
    Right =10290
    Bottom =4845
    DatasheetGridlinesColor =12632256
    MaxRecords =10000000
    Tag ="@DT1=FORMS![F_MENU_DATE]![DT1],@DT2=FORMS![F_MENU_DATE]![DT2]"
    RecSrcDt = Begin
        0xe2654c2cdfbce440
    End
    InputParameters ="@DT1=FORMS![F_MENU_DATE]![DT1],@DT2=FORMS![F_MENU_DATE]![DT2]"
    RecordSource ="Q_LIST_DAYLY_1"
    Caption ="خلاصه ليست فروش روزانه"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Tahoma"
    OnLoad ="[Event Procedure]"
    UnknownProp = {258 ,66 ,12 ,4 ,122 } Begin
        0x40004400540031003d0046004f0052004d00530021005b0046005f004d004500 ,
        0x4e0055005f0044004100540045005d0021005b004400540031005d002c004000 ,
        0x4400540032003d0046004f0052004d00530021005b0046005f004d0045004e00 ,
        0x55005f0044004100540045005d0021005b004400540032005d00
    End
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
        Begin Section
            Height =28476
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3465
                    Top =113
                    ColumnWidth =1080
                    ColumnOrder =0
                    Name ="DATE_N"
                    ControlSource ="DATE_N"
                    OnDblClick ="[Event Procedure]"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6207
                            Top =120
                            Width =915
                            Height =240
                            Name ="Label1"
                            Caption ="تاريخ فاكتور*"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3465
                    Top =453
                    ColumnWidth =1125
                    ColumnOrder =3
                    TabIndex =1
                    Name ="NUMBER"
                    ControlSource ="NUMBER"
                    OnDblClick ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6072
                            Top =450
                            Width =1050
                            Height =240
                            Name ="Label3"
                            Caption ="شماره حواله*"
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
                    Left =3465
                    Top =793
                    ColumnWidth =1335
                    ColumnOrder =8
                    TabIndex =2
                    Name ="SumOfM_NAGHD"
                    ControlSource ="SumOfM_NAGHD"
                    Format ="#,###"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5747
                            Top =793
                            Width =1290
                            Height =240
                            Name ="Label5"
                            Caption ="مبلغ نقد"
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
                    Left =3465
                    Top =1133
                    ColumnWidth =975
                    ColumnOrder =9
                    TabIndex =3
                    Name ="SumOfMABL_VAR"
                    ControlSource ="SumOfMABL_VAR"
                    Format ="#,###"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5702
                            Top =1133
                            Width =1335
                            Height =240
                            Name ="Label7"
                            Caption ="مبلغ واريزي"
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
                    Left =3465
                    Top =1474
                    ColumnWidth =885
                    ColumnOrder =10
                    TabIndex =4
                    Name ="SumOfMABL_HAV"
                    ControlSource ="SumOfMABL_HAV"
                    Format ="#,###"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5732
                            Top =1474
                            Width =1305
                            Height =240
                            Name ="Label9"
                            Caption ="مبلغ حواله"
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
                    Left =3465
                    Top =1814
                    ColumnWidth =1065
                    ColumnOrder =11
                    TabIndex =5
                    Name ="SumOfMABL_HAZ"
                    ControlSource ="SumOfMABL_HAZ"
                    Format ="#,###"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5702
                            Top =1814
                            Width =1335
                            Height =240
                            Name ="Label11"
                            Caption ="مبلغ سرويس"
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
                    Left =3465
                    Top =2154
                    ColumnWidth =1110
                    ColumnOrder =12
                    TabIndex =6
                    Name ="SumOfTAKHFIF"
                    ControlSource ="SumOfTAKHFIF"
                    Format ="#,###"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5852
                            Top =2154
                            Width =1185
                            Height =240
                            Name ="Label13"
                            Caption ="مبلغ تخفيف"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2676
                    Top =2494
                    Width =2490
                    ColumnWidth =1020
                    ColumnOrder =15
                    TabIndex =7
                    Name ="hes"
                    ControlSource ="hes"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6702
                            Top =2490
                            Width =540
                            Height =240
                            Name ="Label15"
                            Caption ="حساب"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =201
                    Top =2834
                    Width =4965
                    ColumnWidth =4530
                    ColumnOrder =4
                    TabIndex =8
                    Name ="NAME"
                    ControlSource ="NAME"
                    OnDblClick ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5852
                            Top =2834
                            Width =1185
                            Height =240
                            Name ="Label17"
                            Caption ="*نام حساب "
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
                    Left =3465
                    Top =3174
                    ColumnOrder =14
                    TabIndex =9
                    Name ="JAMCH"
                    ControlSource ="JAMCH"
                    Format ="#,###"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6447
                            Top =3180
                            Width =690
                            Height =240
                            Name ="Label19"
                            Caption ="جمع چك"
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
                    Left =3061
                    Top =3462
                    ColumnOrder =6
                    TabIndex =10
                    Name ="JAMF"
                    ControlSource ="JAMF"
                    Format ="#,###"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6163
                            Top =3458
                            Width =735
                            Height =405
                            Name ="Label21"
                            Caption ="جمع فاكتور"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1716
                    Top =4535
                    Width =3450
                    ColumnOrder =18
                    TabIndex =11
                    Name ="USER_NAME"
                    ControlSource ="USER_NAME"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6062
                            Top =4535
                            Width =975
                            Height =240
                            Name ="Label27"
                            Caption ="نام كاربر"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3456
                    Top =5215
                    Width =1710
                    ColumnWidth =930
                    ColumnOrder =19
                    TabIndex =12
                    Name ="SHNAME"
                    ControlSource ="SHNAME"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6497
                            Top =5215
                            Width =540
                            Height =240
                            Name ="Label31"
                            Caption ="شيفت"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2676
                    Top =5555
                    Width =2490
                    ColumnWidth =1770
                    ColumnOrder =20
                    TabIndex =13
                    Name ="DEPNAME"
                    ControlSource ="DEPNAME"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6762
                            Top =5550
                            Width =555
                            Height =405
                            Name ="Label33"
                            Caption ="نام واحد"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =93
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =3288
                    Top =6009
                    ColumnOrder =13
                    TabIndex =14
                    Name ="SMBAA"
                    ControlSource ="SMBAA"
                    Format ="#,###"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6042
                            Top =6015
                            Width =840
                            Height =240
                            Name ="Label36"
                            Caption ="ارش افزوده"
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
                    IMESentenceMode =3
                    Left =3571
                    Top =6519
                    ColumnOrder =7
                    TabIndex =15
                    Name ="mand"
                    ControlSource ="mand"
                    Format ="#,###"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6447
                            Top =6525
                            Width =855
                            Height =240
                            Name ="Label37"
                            Caption ="مانده فاكتور"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =3628
                    Top =6859
                    ColumnWidth =3000
                    ColumnOrder =5
                    TabIndex =16
                    Name ="MOLAH"
                    ControlSource ="MOLAH"
                    OnDblClick ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =6280
                            Top =6859
                            Width =750
                            Height =240
                            Name ="Label38"
                            Caption ="*ملاحظات"
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
                    Left =2551
                    Top =4081
                    ColumnWidth =1417
                    ColumnOrder =16
                    TabIndex =17
                    Name ="SHIFT"
                    ControlSource ="SHIFT"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5353
                            Top =4081
                            Width =600
                            Height =240
                            Name ="Label39"
                            Caption ="شيفت"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =247
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2948
                    Top =7029
                    ColumnWidth =1417
                    ColumnOrder =17
                    TabIndex =18
                    Name ="DEPATMAN"
                    ControlSource ="DEPATMAN"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =5405
                            Top =7029
                            Width =945
                            Height =240
                            Name ="Label40"
                            Caption ="واحد"
                        End
                    End
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    ColumnCount =3
                    Left =1871
                    Top =7596
                    ColumnWidth =3495
                    ColumnOrder =22
                    TabIndex =19
                    Name ="MOIN_HAV"
                    ControlSource ="MOIN_HAV"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT     RTRIM(CAST(N_KOL AS nvarchar)) + '-' + RTRIM(CAST(NUMBER AS nvarchar)"
                        ") + '-' + RTRIM(CAST(TNUMBER AS nvarchar)) AS Expr1, NAME, \015\012             "
                        "         RTRIM(CAST(N_KOL AS nvarchar)) + '-' + RTRIM(CAST(NUMBER AS nvarchar)) "
                        "+ '-' + RTRIM(CAST(TNUMBER AS nvarchar)) AS Expr2\015\012FROM         TDETA_HES\015"
                        "\012ORDER BY NAME"
                    ColumnWidths ="0;3402"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4373
                            Top =7596
                            Width =900
                            Height =240
                            Name ="Label41"
                            Caption ="معين حواله"
                        End
                    End
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    ColumnCount =3
                    Left =1871
                    Top =7936
                    ColumnWidth =3750
                    ColumnOrder =23
                    TabIndex =20
                    Name ="MOIN_HAZ"
                    ControlSource ="MOIN_HAZ"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT     RTRIM(CAST(N_KOL AS nvarchar)) + '-' + RTRIM(CAST(NUMBER AS nvarchar)"
                        ") + '-' + RTRIM(CAST(TNUMBER AS nvarchar)) AS Expr1, NAME, \015\012             "
                        "         RTRIM(CAST(N_KOL AS nvarchar)) + '-' + RTRIM(CAST(NUMBER AS nvarchar)) "
                        "+ '-' + RTRIM(CAST(TNUMBER AS nvarchar)) AS Expr2\015\012FROM         TDETA_HES\015"
                        "\012ORDER BY NAME"
                    ColumnWidths ="0;3402"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4343
                            Top =7936
                            Width =930
                            Height =240
                            Name ="Label42"
                            Caption ="معين هزينه"
                        End
                    End
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    ColumnCount =3
                    Left =1871
                    Top =8276
                    ColumnWidth =3375
                    ColumnOrder =21
                    TabIndex =21
                    Name ="MOIN_VAR"
                    ControlSource ="MOIN_VAR"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT     RTRIM(CAST(N_KOL AS nvarchar)) + '-' + RTRIM(CAST(NUMBER AS nvarchar)"
                        ") + '-' + RTRIM(CAST(TNUMBER AS nvarchar)) AS Expr1, NAME, \015\012             "
                        "         RTRIM(CAST(N_KOL AS nvarchar)) + '-' + RTRIM(CAST(NUMBER AS nvarchar)) "
                        "+ '-' + RTRIM(CAST(TNUMBER AS nvarchar)) AS Expr2\015\012FROM         TDETA_HES\015"
                        "\012ORDER BY NAME"
                    ColumnWidths ="0;3402"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4283
                            Top =8276
                            Width =990
                            Height =240
                            Name ="Label43"
                            Caption ="معين واريزي"
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
                    IMESentenceMode =3
                    Left =87
                    Top =10155
                    ColumnOrder =24
                    TabIndex =22
                    Name ="SIMBAA"
                    ControlSource ="SIMBAA"
                    Format ="#,###"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2847
                            Top =10155
                            Width =1260
                            Height =240
                            Name ="Label45"
                            Caption ="کنترل ارزش افزوده"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =2721
                    Top =9070
                    ColumnOrder =2
                    TabIndex =23
                    Name ="date_h"
                    ControlSource ="date_h"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5478
                            Top =9070
                            Width =795
                            Height =240
                            Name ="Label46"
                            Caption ="تاريخ حواله"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =2721
                    Top =9430
                    ColumnOrder =25
                    TabIndex =24
                    Name ="Text47"
                    ControlSource ="ddf"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5472
                            Top =9435
                            Width =825
                            Height =240
                            Name ="Label48"
                            Caption ="تفاوت تاريخ"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2712
                    Top =10653
                    ColumnWidth =1065
                    ColumnOrder =1
                    TabIndex =25
                    Name ="NUMBER1"
                    ControlSource ="NUMBER1"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5319
                            Top =10650
                            Width =1050
                            Height =240
                            Name ="Label50"
                            Caption ="شماره فاكتور"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    ReadingOrder =1
                    Left =2008
                    Top =12699
                    ColumnWidth =915
                    TabIndex =26
                    Name ="TICMBAA"
                    ControlSource ="TICMBAA"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =1347
                            Top =12675
                            Width =810
                            Height =240
                            Name ="Label53"
                            Caption ="تيک ماليات"
                        End
                    End
                End
                Begin TextBox
                    FELineBreak = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1757
                    Top =13209
                    TabIndex =27
                    Name ="N_S"
                    ControlSource ="N_S"
                    AsianLineBreak =255
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =4214
                            Top =13209
                            Width =945
                            Height =240
                            Name ="Label62"
                            Caption ="شماره سند"
                        End
                    End
                End
                Begin TextBox
                    FELineBreak = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1757
                    Top =13549
                    TabIndex =28
                    Name ="VAS"
                    ControlSource ="VAS"
                    AsianLineBreak =255
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =4737
                            Top =13545
                            Width =990
                            Height =240
                            Name ="Label63"
                            Caption ="کرايه به عهده"
                        End
                    End
                End
                Begin TextBox
                    FELineBreak = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1757
                    Top =13889
                    ColumnWidth =1095
                    TabIndex =29
                    Name ="MAS"
                    ControlSource ="MAS"
                    AsianLineBreak =255
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =4711
                            Top =13891
                            Width =1185
                            Height =240
                            Name ="Label64"
                            Caption ="مدت"
                        End
                    End
                End
                Begin TextBox
                    FELineBreak = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1757
                    Top =14229
                    TabIndex =30
                    Name ="TAH"
                    ControlSource ="TAH"
                    AsianLineBreak =255
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =4094
                            Top =14229
                            Width =1065
                            Height =240
                            Name ="Label65"
                            Caption ="تحويل گيردنده"
                        End
                    End
                End
                Begin TextBox
                    FELineBreak = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1757
                    Top =14570
                    TabIndex =31
                    Name ="Text54"
                    ControlSource ="NUMBER1"
                    AsianLineBreak =255
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =3569
                            Top =14570
                            Width =1590
                            Height =240
                            Name ="Label66"
                            Caption ="شماره فاكتور برگشت"
                        End
                    End
                End
                Begin TextBox
                    FELineBreak = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1757
                    Top =14910
                    TabIndex =32
                    Name ="ANBAR"
                    ControlSource ="ANBAR"
                    AsianLineBreak =255
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =4769
                            Top =14910
                            Width =390
                            Height =240
                            Name ="Label67"
                            Caption ="انبار"
                        End
                    End
                End
                Begin TextBox
                    FELineBreak = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1757
                    Top =15250
                    TabIndex =33
                    Name ="Text55"
                    ControlSource ="MOIN_VAR"
                    AsianLineBreak =255
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =4169
                            Top =15250
                            Width =990
                            Height =240
                            Name ="Label68"
                            Caption ="معين واريزي"
                        End
                    End
                End
                Begin TextBox
                    FELineBreak = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1757
                    Top =15590
                    TabIndex =34
                    Name ="Text56"
                    ControlSource ="MOIN_HAV"
                    AsianLineBreak =255
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =4259
                            Top =15590
                            Width =900
                            Height =240
                            Name ="Label69"
                            Caption ="معين حواله"
                        End
                    End
                End
                Begin TextBox
                    FELineBreak = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1757
                    Top =15930
                    TabIndex =35
                    Name ="Text57"
                    ControlSource ="MOIN_HAZ"
                    AsianLineBreak =255
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =4229
                            Top =15930
                            Width =930
                            Height =240
                            Name ="Label70"
                            Caption ="معين هزينه"
                        End
                    End
                End
                Begin TextBox
                    FELineBreak = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1757
                    Top =16270
                    TabIndex =36
                    Name ="MOIN_KHF"
                    ControlSource ="MOIN_KHF"
                    AsianLineBreak =255
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =4169
                            Top =16270
                            Width =990
                            Height =240
                            Name ="Label71"
                            Caption ="معين تخفيف"
                        End
                    End
                End
                Begin TextBox
                    FELineBreak = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1757
                    Top =16611
                    TabIndex =37
                    Name ="ANBARF"
                    ControlSource ="ANBARF"
                    AsianLineBreak =255
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =4304
                            Top =16611
                            Width =855
                            Height =240
                            Name ="Label72"
                            Caption ="انبار فرعي"
                        End
                    End
                End
                Begin TextBox
                    FELineBreak = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1757
                    Top =16951
                    TabIndex =38
                    Name ="Text58"
                    ControlSource ="DEPATMAN"
                    AsianLineBreak =255
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =4559
                            Top =16951
                            Width =600
                            Height =240
                            Name ="Label73"
                            Caption ="دپاتمان"
                        End
                    End
                End
                Begin TextBox
                    FELineBreak = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1757
                    Top =17291
                    TabIndex =39
                    Name ="Text59"
                    ControlSource ="SHIFT"
                    AsianLineBreak =255
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =4559
                            Top =17291
                            Width =600
                            Height =240
                            Name ="Label74"
                            Caption ="شيفت"
                        End
                    End
                End
                Begin TextBox
                    FELineBreak = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1758
                    Top =17688
                    TabIndex =40
                    Name ="CUST_KIND"
                    ControlSource ="CUST_KIND"
                    AsianLineBreak =255
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =4200
                            Top =17688
                            Width =960
                            Height =240
                            Name ="Label75"
                            Caption ="نوع مشتري"
                        End
                    End
                End
                Begin TextBox
                    FELineBreak = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1757
                    Top =17971
                    TabIndex =41
                    Name ="SHARAYET"
                    ControlSource ="SHARAYET"
                    AsianLineBreak =255
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =4244
                            Top =17971
                            Width =915
                            Height =240
                            Name ="Label76"
                            Caption ="شرايط"
                        End
                    End
                End
                Begin TextBox
                    FELineBreak = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1757
                    Top =18311
                    TabIndex =42
                    Name ="SGN1"
                    ControlSource ="SGN1"
                    AsianLineBreak =255
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =4634
                            Top =18311
                            Width =525
                            Height =240
                            Name ="Label77"
                            Caption ="امضا1"
                        End
                    End
                End
                Begin TextBox
                    FELineBreak = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1757
                    Top =18652
                    TabIndex =43
                    Name ="SGN2"
                    ControlSource ="SGN2"
                    AsianLineBreak =255
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =4634
                            Top =18652
                            Width =525
                            Height =240
                            Name ="Label78"
                            Caption ="امضا2"
                        End
                    End
                End
                Begin TextBox
                    FELineBreak = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1757
                    Top =18992
                    TabIndex =44
                    Name ="SGN3"
                    ControlSource ="SGN3"
                    AsianLineBreak =255
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =4634
                            Top =18992
                            Width =525
                            Height =240
                            Name ="Label79"
                            Caption ="امضا3"
                        End
                    End
                End
                Begin TextBox
                    FELineBreak = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1757
                    Top =19332
                    TabIndex =45
                    Name ="SGN4"
                    ControlSource ="SGN4"
                    AsianLineBreak =255
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =4634
                            Top =19332
                            Width =525
                            Height =240
                            Name ="Label80"
                            Caption ="امضا4"
                        End
                    End
                End
                Begin TextBox
                    FELineBreak = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1757
                    Top =19672
                    TabIndex =46
                    Name ="HMBAA"
                    ControlSource ="HMBAA"
                    AsianLineBreak =255
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =4501
                            Top =19666
                            Width =1410
                            Height =240
                            Name ="Label81"
                            Caption ="حساب ارزش افزوده"
                        End
                    End
                End
                Begin TextBox
                    FELineBreak = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1757
                    Top =20012
                    TabIndex =47
                    Name ="TAMIR"
                    ControlSource ="TAMIR"
                    AsianLineBreak =255
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =4546
                            Top =20011
                            Width =675
                            Height =240
                            Name ="Label82"
                            Caption ="تعمير کار"
                        End
                    End
                End
                Begin TextBox
                    FELineBreak = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1757
                    Top =20352
                    TabIndex =48
                    Name ="Text60"
                    ControlSource ="TICMBAA"
                    AsianLineBreak =255
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =4351
                            Top =20356
                            Width =1215
                            Height =240
                            Name ="Label83"
                            Caption ="تيک ارزش افزوده"
                        End
                    End
                End
                Begin TextBox
                    FELineBreak = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1757
                    Top =20692
                    TabIndex =49
                    Name ="TKHF"
                    ControlSource ="TKHF"
                    AsianLineBreak =255
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =4651
                            Top =20686
                            Width =945
                            Height =240
                            Name ="Label84"
                            Caption ="%ت.ن"
                        End
                    End
                End
                Begin TextBox
                    FELineBreak = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1757
                    Top =21033
                    TabIndex =50
                    Name ="SADER"
                    ControlSource ="SADER"
                    AsianLineBreak =255
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =4529
                            Top =21033
                            Width =630
                            Height =240
                            Name ="Label85"
                            Caption ="صادرات"
                        End
                    End
                End
                Begin TextBox
                    FELineBreak = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1757
                    Top =21373
                    TabIndex =51
                    Name ="ARZD"
                    ControlSource ="ARZD"
                    AsianLineBreak =255
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =4621
                            Top =21376
                            Width =675
                            Height =240
                            Name ="Label86"
                            Caption ="قيمت ارز"
                        End
                    End
                End
                Begin TextBox
                    FELineBreak = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1757
                    Top =21713
                    TabIndex =52
                    Name ="ARZKIND"
                    ControlSource ="ARZKIND"
                    AsianLineBreak =255
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =4364
                            Top =21713
                            Width =795
                            Height =240
                            Name ="Label87"
                            Caption ="نوع ارز"
                        End
                    End
                End
                Begin TextBox
                    FELineBreak = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1757
                    Top =22053
                    TabIndex =53
                    Name ="JAY"
                    ControlSource ="JAY"
                    AsianLineBreak =255
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =4756
                            Top =22051
                            Width =735
                            Height =240
                            Name ="Label88"
                            Caption ="تيک جايزه"
                        End
                    End
                End
                Begin TextBox
                    FELineBreak = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1757
                    Top =22393
                    TabIndex =54
                    Name ="MODAT_PPID"
                    ControlSource ="MODAT_PPID"
                    AsianLineBreak =255
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =4049
                            Top =22393
                            Width =1110
                            Height =240
                            Name ="Label89"
                            Caption ="نحوه پرداخت"
                        End
                    End
                End
                Begin TextBox
                    FELineBreak = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1757
                    Top =22733
                    TabIndex =55
                    Name ="PEPID"
                    ControlSource ="PEPID"
                    AsianLineBreak =255
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =4591
                            Top =22726
                            Width =915
                            Height =240
                            Name ="Label90"
                            Caption ="الگوي قيمت"
                        End
                    End
                End
                Begin TextBox
                    FELineBreak = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1757
                    Top =23074
                    TabIndex =56
                    Name ="PEID"
                    ControlSource ="PEID"
                    AsianLineBreak =255
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =4681
                            Top =23071
                            Width =975
                            Height =240
                            Name ="Label91"
                            Caption ="الگوي تخفيف"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Top =5725
                    TabIndex =57
                    Name ="MCODEM"
                    ControlSource ="MCODEM"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =2607
                            Top =5725
                            Width =795
                            Height =240
                            Name ="Label61"
                            Caption ="کدملي"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Top =6066
                    TabIndex =58
                    Name ="MOBILE"
                    ControlSource ="MOBILE"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =2712
                            Top =6066
                            Width =690
                            Height =240
                            Name ="Label92"
                            Caption ="موبايل"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Top =6406
                    TabIndex =59
                    Name ="OSTANID"
                    ControlSource ="OSTANID"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2592
                            Top =6406
                            Width =810
                            Height =240
                            Name ="Label93"
                            Caption ="استان"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Top =6746
                    TabIndex =60
                    Name ="SHAHRID"
                    ControlSource ="SHAHRID"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2592
                            Top =6746
                            Width =810
                            Height =240
                            Name ="Label94"
                            Caption ="شهر"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "Q_LIST_DALY.cls"
