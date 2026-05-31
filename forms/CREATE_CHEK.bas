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
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =11225
    DatasheetFontHeight =10
    ItemSuffix =7
    Left =2025
    Top =1740
    Right =13245
    Bottom =3645
    DatasheetGridlinesColor =12632256
    MaxRecords =1000000
    RecSrcDt = Begin
        0x0544c29cd1d1e340
    End
    Caption ="چك گروهي"
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
            Height =1927
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
                    Left =8725
                    Top =340
                    Width =1266
                    Height =345
                    Name ="N_SERI"
                    AsianLineBreak =0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =10107
                            Top =340
                            Width =1005
                            Height =345
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
                    Left =5681
                    Top =345
                    Width =1176
                    Height =360
                    TabIndex =1
                    Name ="DATE_S"
                    InputMask ="####/##/##"
                    AsianLineBreak =0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6936
                            Top =350
                            Width =1725
                            Height =345
                            Name ="Label7"
                            Caption ="تاريخ سررسيد اولين چك"
                        End
                    End
                End
                Begin TextBox
                    FELineBreak = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =93
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =6395
                    Top =795
                    Width =1326
                    Height =345
                    TabIndex =5
                    Name ="MABL"
                    Format ="Standard"
                    OnKeyPress ="[Event Procedure]"
                    AsianLineBreak =0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =7835
                            Top =795
                            Width =900
                            Height =330
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
                    Top =795
                    Width =2160
                    Height =345
                    TabIndex =6
                    Name ="NAME_TAH"
                    AsianLineBreak =0
                    Begin
                        Begin Label
                            OverlapFlags =87
                            Left =4660
                            Top =797
                            Width =1725
                            Height =345
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
                    Top =795
                    Width =1476
                    Height =345
                    TabIndex =7
                    Name ="N_HESAB"
                    AsianLineBreak =0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1652
                            Top =795
                            Width =720
                            Height =345
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
                    Left =113
                    Top =340
                    Width =2325
                    Height =360
                    TabIndex =3
                    Name ="SHOBEH"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT PAY_GETD.SHOBEH FROM PAY_GETD GROUP BY PAY_GETD.SHOBEH ORDER BY PAY_GETD."
                        "SHOBEH;"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2569
                            Top =340
                            Width =720
                            Height =345
                            Name ="Label9"
                            Caption ="نام شعبه"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =4081
                    Top =1360
                    Width =1191
                    Height =397
                    TabIndex =10
                    Name ="Command3"
                    Caption ="تاييد"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    Cancel = NotDefault
                    OverlapFlags =85
                    Left =2895
                    Top =1365
                    Width =1191
                    Height =397
                    TabIndex =11
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
                    Left =3345
                    Top =345
                    Width =1815
                    Height =345
                    TabIndex =2
                    Name ="bank"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TCOD_BANKS.CODE, TCOD_BANKS.NAMES FROM TCOD_BANKS ORDER BY TCOD_BANKS.NAM"
                        "ES;"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5291
                            Top =345
                            Width =345
                            Height =345
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
                    Left =8760
                    Top =795
                    Width =1560
                    Height =315
                    TabIndex =4
                    Name ="LIST_NO"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT LIST_NO FROM PAY_GETD GROUP BY LIST_NO"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =10472
                            Top =799
                            Width =690
                            Height =240
                            Name ="Label76"
                            Caption ="كد شعبه"
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
                    Left =9066
                    Top =1215
                    Width =1071
                    Height =555
                    FontSize =14
                    FontWeight =700
                    TabIndex =8
                    BackColor =13434828
                    ForeColor =255
                    Name ="NUM"
                    DefaultValue ="1"
                    FontName ="Titr"
                    AsianLineBreak =0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =10193
                            Top =1335
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
                    Left =6465
                    Top =1271
                    Width =1566
                    Height =555
                    FontSize =14
                    FontWeight =700
                    TabIndex =9
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
                            Left =8082
                            Top =1391
                            Width =915
                            Height =240
                            Name ="Label10"
                            Caption ="فاصله چكها:"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "CREATE_CHEK.cls"
