Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    ControlBox = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    DefaultView =0
    AllowUpdating =4
    ScrollBars =0
    ViewsAllowed =1
    TabularCharSet =178
    TabularFamily =42
    BorderStyle =3
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =9862
    DatasheetFontHeight =10
    ItemSuffix =54
    Left =9390
    Top =3180
    Right =19245
    Bottom =5490
    HelpContextId =500203
    DatasheetGridlinesColor =12632256
    MaxRecords =100000
    RecSrcDt = Begin
        0xbace0f5a1a4ce240
    End
    Caption ="دفتر معين تفصيلي"
    OnOpen ="[Event Procedure]"
    HelpFile ="dena.hlp"
    DatasheetFontName ="Arial (Arabic)"
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
        Begin OptionButton
            SpecialEffect =2
            LabelX =230
            LabelY =-30
        End
        Begin OptionGroup
            SpecialEffect =3
            Width =1701
            Height =1701
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
        Begin UnboundObjectFrame
            SpecialEffect =2
            OldBorderStyle =1
            Width =4536
            Height =2835
        End
        Begin Section
            Height =2324
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    TextFontCharSet =0
                    Left =5567
                    Top =1860
                    Width =870
                    Height =345
                    FontWeight =700
                    TabIndex =10
                    Name ="Command5"
                    Caption ="تاييد"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    Cancel = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    Left =4699
                    Top =1860
                    Width =870
                    Height =347
                    FontWeight =700
                    TabIndex =11
                    Name ="Command6"
                    Caption ="انصراف"
                    OnClick ="[Event Procedure]"
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =93
                    TextFontCharSet =0
                    Left =9558
                    Width =234
                    Height =255
                    TabIndex =5
                    Name ="DTT"
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    AllowedText =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =7064
                    Top =615
                    Width =1431
                    Height =300
                    TabIndex =6
                    Name ="DT1"
                    Format ="0000\\/00\\/00"
                    BeforeUpdate ="[Event Procedure]"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =8589
                            Top =617
                            Width =645
                            Height =285
                            Name ="Label21"
                            Caption ="از تاريخ:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    AllowedText =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3526
                    Top =630
                    Width =1416
                    Height =300
                    TabIndex =7
                    Name ="DT2"
                    Format ="0000\\/00\\/00"
                    BeforeUpdate ="[Event Procedure]"
                    DefaultValue ="=[Forms]![baseknow]![DTT]"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =5036
                            Top =632
                            Width =630
                            Height =285
                            Name ="Label23"
                            Caption ="تا تاريخ:"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    Visible = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    AllowedText =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =9365
                    Top =735
                    Width =456
                    Height =300
                    TabIndex =2
                    Name ="HKOL"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TOTA_HES.NUMBER FROM TOTA_HES ORDER BY TOTA_HES.NUMBER;"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    Visible = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    AllowedText =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =9365
                    Top =1093
                    Width =441
                    Height =300
                    TabIndex =3
                    Name ="HMOIN"
                    RowSourceType ="Table/View/StoredProc"
                End
                Begin TextBox
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    AllowedText =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4022
                    Top =1470
                    Width =4431
                    Height =300
                    TabIndex =9
                    Name ="PY"
                    DefaultValue ="\" \""
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextFontCharSet =0
                            Left =8540
                            Top =1478
                            Width =1275
                            Height =285
                            Name ="Label29"
                            Caption ="پيغام پايان گزارش:"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    Visible = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =223
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    AllowedText =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =9327
                    Top =1474
                    Width =396
                    Height =300
                    TabIndex =4
                    Name ="HTAF"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TDETA_HES.TNUMBER, TDETA_HES.N_KOL, TDETA_HES.NUMBER FROM DETA_HES INNER "
                        "JOIN TDETA_HES ON (DETA_HES.NUMBER = TDETA_HES.NUMBER) AND (DETA_HES.N_KOL = TDE"
                        "TA_HES.N_KOL) AND (DETA_HES.NUMBER = TDETA_HES.NUMBER) AND (DETA_HES.N_KOL = TDE"
                        "TA_HES.N_KOL) WHERE (((TDETA_HES.N_KOL)=[Forms]![F_MENU_KOL_MOIN_TAFZIL]![HKOL])"
                        " AND ((TDETA_HES.NUMBER)=[Forms]![F_MENU_KOL_MOIN_TAFZIL]![Hmoin]) AND ((TDETA_H"
                        "ES.CODE_E) Like IIf([Forms]![Baseknow]![PERSON]=2,\"0\",\"*\") Or (TDETA_HES.COD"
                        "E_E) Is Null));"
                    OnNotInList ="[Event Procedure]"
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =93
                    TextFontCharSet =0
                    Left =9558
                    Top =375
                    Width =234
                    Height =255
                    TabIndex =12
                    Name ="HTTAF"
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
                    Left =7022
                    Top =225
                    Width =2121
                    Height =300
                    Name ="Combo34"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT hes FROM CUST_HESAB"
                    AfterUpdate ="[Event Procedure]"
                    OnExit ="[Event Procedure]"
                    OnNotInList ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextFontCharSet =0
                            Left =9232
                            Top =225
                            Width =630
                            Height =285
                            Name ="Label35"
                            Caption ="حساب :"
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
                    Left =1244
                    Top =225
                    Width =5721
                    Height =300
                    TabIndex =1
                    Name ="Combo36"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT hes, NAME AS nam, hes AS Expr1 FROM CUST_HESAB"
                    ColumnWidths ="0"
                    AfterUpdate ="[Event Procedure]"
                    OnExit ="[Event Procedure]"
                End
                Begin OptionGroup
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =247
                    Left =4019
                    Top =1020
                    Width =5327
                    Height =796
                    TabIndex =8
                    Name ="Frame34"
                    DefaultValue ="3"
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =247
                            TextFontCharSet =0
                            Left =8596
                            Top =1077
                            Width =750
                            Height =285
                            BackColor =-2147483633
                            Name ="Label38"
                            Caption ="بر اساس:"
                        End
                        Begin OptionButton
                            OverlapFlags =119
                            ReadingOrder =1
                            Left =8264
                            Top =1138
                            OptionValue =1
                            Name ="Option40"
                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    TextFontCharSet =0
                                    Left =7379
                                    Top =1080
                                    Width =885
                                    Height =270
                                    Name ="Label41"
                                    Caption ="شماره سند"
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =119
                            ReadingOrder =1
                            Left =5627
                            Top =1168
                            OptionValue =3
                            Name ="Option45"
                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    TextFontCharSet =0
                                    Left =4802
                                    Top =1080
                                    Width =825
                                    Height =300
                                    Name ="Label46"
                                    Caption ="صورتحساب"
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =119
                            ReadingOrder =1
                            Left =6869
                            Top =1168
                            OptionValue =2
                            Name ="Option42"
                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    TextFontCharSet =0
                                    Left =6044
                                    Top =1080
                                    Width =825
                                    Height =300
                                    Name ="Label43"
                                    Caption ="شماره مبنا"
                                End
                            End
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
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =7891
                    Top =1894
                    Width =396
                    Height =300
                    TabIndex =13
                    Name ="HTAF2"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TDETA_HES.TNUMBER, TDETA_HES.N_KOL, TDETA_HES.NUMBER FROM DETA_HES INNER "
                        "JOIN TDETA_HES ON (DETA_HES.NUMBER = TDETA_HES.NUMBER) AND (DETA_HES.N_KOL = TDE"
                        "TA_HES.N_KOL) AND (DETA_HES.NUMBER = TDETA_HES.NUMBER) AND (DETA_HES.N_KOL = TDE"
                        "TA_HES.N_KOL) WHERE (((TDETA_HES.N_KOL)=[Forms]![F_MENU_KOL_MOIN_TAFZIL]![HKOL])"
                        " AND ((TDETA_HES.NUMBER)=[Forms]![F_MENU_KOL_MOIN_TAFZIL]![Hmoin]) AND ((TDETA_H"
                        "ES.CODE_E) Like IIf([Forms]![Baseknow]![PERSON]=2,\"0\",\"*\") Or (TDETA_HES.COD"
                        "E_E) Is Null));"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    Visible = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4062
                    Top =1927
                    Width =396
                    Height =300
                    TabIndex =14
                    Name ="HTAF3"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TDETA_HES.TNUMBER, TDETA_HES.N_KOL, TDETA_HES.NUMBER FROM DETA_HES INNER "
                        "JOIN TDETA_HES ON (DETA_HES.NUMBER = TDETA_HES.NUMBER) AND (DETA_HES.N_KOL = TDE"
                        "TA_HES.N_KOL) AND (DETA_HES.NUMBER = TDETA_HES.NUMBER) AND (DETA_HES.N_KOL = TDE"
                        "TA_HES.N_KOL) WHERE (((TDETA_HES.N_KOL)=[Forms]![F_MENU_KOL_MOIN_TAFZIL]![HKOL])"
                        " AND ((TDETA_HES.NUMBER)=[Forms]![F_MENU_KOL_MOIN_TAFZIL]![Hmoin]) AND ((TDETA_H"
                        "ES.CODE_E) Like IIf([Forms]![Baseknow]![PERSON]=2,\"0\",\"*\") Or (TDETA_HES.COD"
                        "E_E) Is Null));"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    Visible = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3211
                    Top =1927
                    Width =396
                    Height =300
                    TabIndex =15
                    Name ="HTAF4"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TDETA_HES.TNUMBER, TDETA_HES.N_KOL, TDETA_HES.NUMBER FROM DETA_HES INNER "
                        "JOIN TDETA_HES ON (DETA_HES.NUMBER = TDETA_HES.NUMBER) AND (DETA_HES.N_KOL = TDE"
                        "TA_HES.N_KOL) AND (DETA_HES.NUMBER = TDETA_HES.NUMBER) AND (DETA_HES.N_KOL = TDE"
                        "TA_HES.N_KOL) WHERE (((TDETA_HES.N_KOL)=[Forms]![F_MENU_KOL_MOIN_TAFZIL]![HKOL])"
                        " AND ((TDETA_HES.NUMBER)=[Forms]![F_MENU_KOL_MOIN_TAFZIL]![Hmoin]) AND ((TDETA_H"
                        "ES.CODE_E) Like IIf([Forms]![Baseknow]![PERSON]=2,\"0\",\"*\") Or (TDETA_HES.COD"
                        "E_E) Is Null));"
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =113
                    Top =1077
                    TabIndex =16
                    Name ="txtSort"
                    DefaultValue ="\"DATE_S, BED DESC\""
                End
            End
        End
    End
End
CodeBehindForm
' See "F_MENU_KOL_MOIN_TAFZIL.cls"
