Version =20
VersionRequired =20
Begin Form
    AllowDeletions = NotDefault
    AllowAdditions = NotDefault
    DefaultView =2
    AllowUpdating =4
    ViewsAllowed =2
    TabularCharSet =178
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =4940
    RowHeight =315
    ItemSuffix =34
    Left =270
    Top =210
    Right =6990
    Bottom =7890
    DatasheetForeColor =33554432
    DatasheetGridlinesColor =12632256
    MaxRecords =10000000
    RecSrcDt = Begin
        0x19e9e38d14c9e540
    End
    ServerFilter ="N_SERI = 172548 AND BANK = 11 AND date_s = 14001220"
    RecordSource ="SELECT PAY_GETD_LOG.*, TCOD_BANKS.NAMES, PAY_GETD.SHOBEH, PAY_GETD.MABL, PAY_GET"
        "D.NAME_TAH, PAY_GETD.RADIF, PAY_GETD.CUST_NO FROM PAY_GETD_LOG INNER JOIN PAY_GE"
        "TD ON PAY_GETD_LOG.N_SERI = PAY_GETD.N_SERI AND PAY_GETD_LOG.BANK = PAY_GETD.BAN"
        "K AND PAY_GETD_LOG.DATE_S = PAY_GETD.DATE_S INNER JOIN TCOD_BANKS ON PAY_GETD.BA"
        "NK = TCOD_BANKS.CODE"
    Caption ="سوابق تغيير وضعيت چك"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Tahoma"
    OnLoad ="[Event Procedure]"
    OnError ="[Event Procedure]"
    AllowFormView =0
    RecordSourceQualifier ="dbo"
    Begin
        Begin Label
            BackStyle =0
            TextFontCharSet =178
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
            Height =5682
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =93
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2097
                    Top =680
                    ColumnWidth =1005
                    ColumnOrder =4
                    Name ="N_SERI"
                    ControlSource ="N_SERI"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =396
                            Top =680
                            Width =675
                            Height =240
                            Name ="Label18"
                            Caption ="سريال"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2097
                    Top =1360
                    ColumnWidth =1020
                    ColumnOrder =7
                    TabIndex =1
                    Name ="DATE_S"
                    ControlSource ="DATE_S"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =396
                            Top =1360
                            Width =705
                            Height =240
                            Name ="Label20"
                            Caption ="سررسيد"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2097
                    Top =1700
                    ColumnWidth =1020
                    ColumnOrder =2
                    TabIndex =2
                    Name ="DATE_V"
                    ControlSource ="DATE_V"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =390
                            Top =1695
                            Width =765
                            Height =240
                            Name ="Label21"
                            Caption ="تاريخ تغيير"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2097
                    Top =2040
                    ColumnWidth =840
                    ColumnOrder =3
                    TabIndex =3
                    Name ="DATETIM"
                    ControlSource ="DATETIM"
                    Format ="hh:nn:ss"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =396
                            Top =2040
                            Width =795
                            Height =240
                            Name ="Label22"
                            Caption ="ساعت"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2097
                    Top =3061
                    ColumnOrder =12
                    TabIndex =4
                    Name ="USER_NAME"
                    ControlSource ="USER_NAME"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =396
                            Top =3061
                            Width =1035
                            Height =240
                            Name ="Label25"
                            Caption ="كاربر"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2097
                    Top =3741
                    ColumnOrder =5
                    TabIndex =5
                    Name ="NAMES"
                    ControlSource ="NAMES"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =396
                            Top =3741
                            Width =645
                            Height =240
                            Name ="Label27"
                            Caption ="بانك "
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2097
                    Top =4081
                    ColumnOrder =6
                    TabIndex =6
                    Name ="SHOBEH"
                    ControlSource ="SHOBEH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =396
                            Top =4081
                            Width =735
                            Height =240
                            Name ="Label28"
                            Caption ="شعبه"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2097
                    Top =4422
                    ColumnOrder =8
                    TabIndex =7
                    Name ="MABL"
                    ControlSource ="MABL"
                    Format ="Standard"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =396
                            Top =4422
                            Width =525
                            Height =240
                            Name ="Label29"
                            Caption ="مبلغ"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2097
                    Top =4762
                    ColumnWidth =3675
                    ColumnOrder =9
                    TabIndex =8
                    Name ="NAME_TAH"
                    ControlSource ="NAME_TAH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =390
                            Top =4755
                            Width =975
                            Height =240
                            Name ="Label30"
                            Caption ="پرداخت كننده"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2097
                    Top =5102
                    ColumnWidth =630
                    ColumnOrder =11
                    TabIndex =9
                    Name ="RADIF"
                    ControlSource ="RADIF"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =396
                            Top =5102
                            Width =600
                            Height =240
                            Name ="Label31"
                            Caption ="شماره"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2777
                    Top =5442
                    ColumnOrder =10
                    TabIndex =10
                    Name ="CUST_NO"
                    ControlSource ="CUST_NO"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1076
                            Top =5442
                            Width =1185
                            Height =240
                            Name ="Label32"
                            Caption ="كد مشتري"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    Locked = NotDefault
                    RowSourceTypeInt =1
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =1417
                    Top =56
                    Width =2211
                    Height =315
                    ColumnOrder =0
                    TabIndex =11
                    Name ="VAZ"
                    ControlSource ="VAZ"
                    RowSourceType ="Value List"
                    RowSource ="1;\"نزد صندوق\";2;\"نزد بانك\";3;\"وصول شده\";4;\"واگذار شده\";5;\"برگشت شده\";6"
                        ";\"مسترد شده\""
                    ColumnWidths ="0"
                    DefaultValue ="1"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =3759
                            Top =56
                            Width =1005
                            Height =315
                            Name ="Label50"
                            Caption ="وضعيت چك"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =247
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =1417
                    Top =491
                    Width =2211
                    Height =315
                    ColumnWidth =1725
                    ColumnOrder =1
                    TabIndex =12
                    Name ="SANDUGH"
                    ControlSource ="SANDUGH"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TNUMBER, NAME FROM TDETA_HES WHERE (N_KOL = 113) AND (NUMBER = 1)"
                    ColumnWidths ="0"
                    DefaultValue ="1"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextFontCharSet =0
                            Left =3759
                            Top =491
                            Width =1005
                            Height =315
                            Name ="Label40"
                            Caption ="موقعيت چك"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =906
                    Top =2494
                    TabIndex =13
                    Name ="BANK"
                    ControlSource ="BANK"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3780
                            Top =2490
                            Width =555
                            Height =240
                            Name ="Label33"
                            Caption ="كد بانك"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "PAY_GETD_LOG_FORM.cls"
