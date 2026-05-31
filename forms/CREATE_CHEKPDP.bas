Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    DefaultView =0
    AllowUpdating =4
    ScrollBars =0
    ViewsAllowed =1
    TabularCharSet =178
    TabularFamily =0
    BorderStyle =3
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =12705
    DatasheetFontHeight =10
    ItemSuffix =9
    Left =4395
    Top =3075
    Right =17100
    Bottom =6180
    DatasheetGridlinesColor =12632256
    MaxRecords =1000000
    RecSrcDt = Begin
        0x0544c29cd1d1e340
    End
    Caption ="  چك گروهي پرداختي"
    DatasheetFontName ="Arial"
    OnLoad ="[Event Procedure]"
    OnError ="[Event Procedure]"
    AllowDatasheetView =0
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
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            OldBorderStyle =0
            TextFontCharSet =178
            Width =1701
            LabelX =-1701
            FontName ="Tahoma"
            AsianLineBreak =255
        End
        Begin ComboBox
            SpecialEffect =2
            TextFontCharSet =178
            Width =1701
            LabelX =-1701
            FontName ="Tahoma"
        End
        Begin Section
            Height =3118
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    FELineBreak = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =10247
                    Top =795
                    Width =1266
                    Height =426
                    TabIndex =2
                    Name ="N_SERI"
                    AsianLineBreak =0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =11629
                            Top =795
                            Width =1005
                            Height =426
                            Name ="Label1"
                            Caption ="شماره سريال"
                        End
                    End
                End
                Begin TextBox
                    FELineBreak = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =6798
                    Top =800
                    Width =1581
                    Height =426
                    TabIndex =3
                    Name ="DATE_S"
                    InputMask ="####/##/##"
                    AsianLineBreak =0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =8458
                            Top =805
                            Width =1725
                            Height =426
                            Name ="Label7"
                            Caption ="تاريخ سررسيد اولين چك"
                        End
                    End
                End
                Begin TextBox
                    FELineBreak = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =9810
                    Top =1425
                    Width =1881
                    Height =426
                    TabIndex =6
                    Name ="MABL"
                    Format ="Standard"
                    OnKeyPress ="[Event Procedure]"
                    AsianLineBreak =0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =11805
                            Top =1425
                            Width =900
                            Height =426
                            Name ="Label11"
                            Caption ="مبلغ هر چك"
                        End
                    End
                End
                Begin TextBox
                    FELineBreak = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2565
                    Top =1425
                    Width =5430
                    Height =426
                    TabIndex =7
                    Name ="NAME_TAH"
                    AsianLineBreak =0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =8036
                            Top =1427
                            Width =1725
                            Height =426
                            Name ="Label13"
                            Caption ="نام دريافت كننده و شرح"
                        End
                    End
                End
                Begin TextBox
                    FELineBreak = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =120
                    Top =1425
                    Width =1476
                    Height =426
                    TabIndex =8
                    Name ="N_HESAB"
                    AsianLineBreak =0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1652
                            Top =1425
                            Width =720
                            Height =426
                            Name ="Label15"
                            Caption ="جاري چك"
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
                    Left =105
                    Top =795
                    Width =3180
                    Height =426
                    TabIndex =5
                    Name ="SHOBEH"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT PAY_GETD.SHOBEH FROM PAY_GETD GROUP BY PAY_GETD.SHOBEH ORDER BY PAY_GETD."
                        "SHOBEH;"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3416
                            Top =795
                            Width =720
                            Height =426
                            Name ="Label9"
                            Caption ="نام شعبه"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =4587
                    Top =2664
                    Width =1191
                    Height =397
                    TabIndex =13
                    Name ="Command3"
                    Caption ="تاييد"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    Cancel = NotDefault
                    OverlapFlags =85
                    Left =3401
                    Top =2664
                    Width =1191
                    Height =397
                    TabIndex =14
                    Name ="Command4"
                    Caption ="لغو"
                    OnClick ="[Event Procedure]"
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =4260
                    Top =795
                    Width =1815
                    Height =426
                    TabIndex =4
                    Name ="bank"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TCOD_BANKS.CODE, TCOD_BANKS.NAMES FROM TCOD_BANKS ORDER BY TCOD_BANKS.NAM"
                        "ES;"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6206
                            Top =795
                            Width =345
                            Height =426
                            Name ="Label6"
                            Caption ="بانك"
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
                    Left =4425
                    Top =120
                    Width =7401
                    Height =315
                    BackColor =16777164
                    ForeColor =8388608
                    Name ="CUST_NO"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT hes, NAME + N' ' + hes AS nam, hes AS Expr1 FROM CUST_HESAB"
                    ColumnWidths ="0;4536;567"
                    AfterUpdate ="[Event Procedure]"
                    OnExit ="[Event Procedure]"
                    OnNotInList ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =11910
                            Top =120
                            Width =735
                            Height =300
                            ForeColor =8388608
                            Name ="Label19"
                            Caption =" مشتري:"
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
                    Left =1870
                    Top =113
                    Width =2451
                    Height =315
                    TabIndex =1
                    BackColor =16777164
                    ForeColor =8388608
                    Name ="CUST_NO2"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT hes FROM CUST_HESAB"
                    AfterUpdate ="[Event Procedure]"
                    OnNotInList ="[Event Procedure]"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =4695
                    Top =2040
                    Width =1581
                    Height =427
                    TabIndex =10
                    Name ="KIND"
                    RowSourceType ="Value List"
                    RowSource ="1;\"تجاري\";0;\"غير تجاري\""
                    ColumnWidths ="0"
                    DefaultValue ="1"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6325
                            Top =2097
                            Width =585
                            Height =240
                            Name ="Label8"
                            Caption ="نوع چك"
                        End
                    End
                End
                Begin TextBox
                    FELineBreak = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2714
                    Top =1984
                    Width =1071
                    Height =555
                    FontSize =14
                    FontWeight =700
                    TabIndex =11
                    BackColor =16777164
                    ForeColor =255
                    Name ="NUM"
                    DefaultValue ="1"
                    FontName ="Titr"
                    AsianLineBreak =0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3841
                            Top =2104
                            Width =840
                            Height =240
                            Name ="Label2"
                            Caption ="تعداد فقره:"
                        End
                    End
                End
                Begin TextBox
                    FELineBreak = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =113
                    Top =2040
                    Width =1566
                    Height =555
                    FontSize =14
                    FontWeight =700
                    TabIndex =12
                    BackColor =16777164
                    ForeColor =255
                    Name ="GAP"
                    DefaultValue ="1"
                    FontName ="Titr"
                    InputMask ="#\"      ماه \""
                    AsianLineBreak =0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1730
                            Top =2160
                            Width =915
                            Height =240
                            Name ="Label10"
                            Caption ="فاصله چكها:"
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
                    Left =6975
                    Top =2041
                    Width =4356
                    Height =427
                    TabIndex =9
                    Name ="HES"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT RTRIM(CAST(TOTA_HES.NUMBER AS nvarchar)) + '-' + RTRIM(CAST(DETA_HES.NUMB"
                        "ER AS nvarchar)) + '-' + RTRIM(CAST(TDETA_HES.TNUMBER AS nvarchar)) AS hes, TDET"
                        "A_HES.NAME, TDETA_HES.N_KOL, SAZMAN.BANKHA FROM TOTA_HES INNER JOIN DETA_HES INN"
                        "ER JOIN TDETA_HES ON DETA_HES.NUMBER = TDETA_HES.NUMBER AND DETA_HES.N_KOL = TDE"
                        "TA_HES.N_KOL ON TOTA_HES.NUMBER = DETA_HES.N_KOL INNER JOIN SAZMAN ON TDETA_HES."
                        "N_KOL = SAZMAN.BANKHA AND TDETA_HES.N_KOL = SAZMAN.BANKHA"
                    ColumnWidths ="1421;2835"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            TextAlign =3
                            Left =11415
                            Top =2040
                            Width =1230
                            Height =240
                            Name ="Label4"
                            Caption ="پرداخت از حساب"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "CREATE_CHEKPDP.cls"
