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
    Width =11848
    DatasheetFontHeight =10
    ItemSuffix =11
    Left =1770
    Top =6045
    Right =13620
    Bottom =8700
    DatasheetGridlinesColor =12632256
    MaxRecords =1000000
    RecSrcDt = Begin
        0x0544c29cd1d1e340
    End
    Caption ="چك گروهي دريافتي"
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
            Height =3219
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
                    Left =9287
                    Top =795
                    Width =1266
                    Height =426
                    TabIndex =2
                    Name ="N_SERI"
                    AsianLineBreak =0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =10669
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
                    Left =5838
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
                            Left =7498
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
                    Left =6410
                    Top =1425
                    Width =1881
                    Height =426
                    TabIndex =7
                    Name ="MABL"
                    Format ="Standard"
                    OnKeyPress ="[Event Procedure]"
                    AsianLineBreak =0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =8405
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
                    Left =2459
                    Top =1425
                    Width =2160
                    Height =426
                    TabIndex =8
                    Name ="NAME_TAH"
                    AsianLineBreak =0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4660
                            Top =1427
                            Width =1725
                            Height =426
                            Name ="Label13"
                            Caption ="نام پرداخت كننده و شرح"
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
                    TabIndex =9
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
                    Width =2325
                    Height =426
                    TabIndex =5
                    Name ="SHOBEH"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT PAY_GETD.SHOBEH FROM PAY_GETD GROUP BY PAY_GETD.SHOBEH ORDER BY PAY_GETD."
                        "SHOBEH;"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2561
                            Top =795
                            Width =720
                            Height =426
                            Name ="Label9"
                            Caption ="نام شعبه"
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
                    Left =2721
                    Top =1984
                    Width =1071
                    Height =555
                    FontSize =14
                    FontWeight =700
                    TabIndex =12
                    BackColor =13434828
                    ForeColor =255
                    Name ="NUM"
                    DefaultValue ="1"
                    FontName ="Titr"
                    AsianLineBreak =0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3848
                            Top =2104
                            Width =840
                            Height =240
                            Name ="Label2"
                            Caption ="تعداد فقره:"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =4587
                    Top =2664
                    Width =1191
                    Height =397
                    TabIndex =14
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
                    TabIndex =15
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
                    Left =3337
                    Top =800
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
                            Left =5283
                            Top =800
                            Width =345
                            Height =426
                            Name ="Label6"
                            Caption ="بانك"
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
                    Left =9330
                    Top =1425
                    Width =1560
                    Height =426
                    TabIndex =6
                    Name ="LIST_NO"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT LIST_NO FROM PAY_GETD GROUP BY LIST_NO"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =11042
                            Top =1429
                            Width =690
                            Height =426
                            Name ="Label76"
                            Caption ="كد شعبه"
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
                    Width =6441
                    Height =315
                    BackColor =13434828
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
                            Left =10950
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
                    BackColor =13434828
                    ForeColor =8388608
                    Name ="CUST_NO2"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT hes FROM CUST_HESAB"
                    AfterUpdate ="[Event Procedure]"
                    OnNotInList ="[Event Procedure]"
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
                    Left =7337
                    Top =2055
                    Width =3231
                    Height =427
                    TabIndex =10
                    Name ="VAZ"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TNUMBER, NAME FROM TDETA_HES WHERE (N_KOL = 113) AND (NUMBER = 1)"
                    ColumnWidths ="0"
                    DefaultValue ="1"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =10699
                            Top =2055
                            Width =1005
                            Height =427
                            Name ="Label50"
                            Caption ="نزد صندوق:"
                        End
                    End
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
                    Left =4762
                    Top =2040
                    Width =1581
                    Height =427
                    TabIndex =11
                    Name ="KIND"
                    RowSourceType ="Value List"
                    RowSource ="1;\"تجاري\";0;\"غير تجاري\""
                    ColumnWidths ="0"
                    DefaultValue ="1"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6392
                            Top =2097
                            Width =840
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
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =120
                    Top =2040
                    Width =1566
                    Height =555
                    FontSize =14
                    FontWeight =700
                    TabIndex =13
                    BackColor =13434828
                    ForeColor =255
                    Name ="GAP"
                    DefaultValue ="1"
                    FontName ="Titr"
                    InputMask ="#\"      ماه \""
                    AsianLineBreak =0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1737
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
                    Left =7350
                    Top =2550
                    Width =3231
                    Height =405
                    TabIndex =16
                    Name ="SANDUGH"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TNUMBER, NAME FROM TDETA_HES WHERE (N_KOL = 113) AND (NUMBER = 1)"
                    ColumnWidths ="0"
                    DefaultValue ="1"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =10712
                            Top =2550
                            Width =1005
                            Height =390
                            Name ="Label53"
                            Caption ="موقعيت چك:"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "CREATE_CHEKDP.cls"
