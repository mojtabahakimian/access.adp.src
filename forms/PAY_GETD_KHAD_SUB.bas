Version =20
VersionRequired =20
Begin Form
    KeyPreview = NotDefault
    DefaultView =2
    AllowUpdating =4
    ViewsAllowed =2
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =4880
    RowHeight =270
    ItemSuffix =47
    Left =405
    Top =4380
    Right =13800
    Bottom =9075
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x05ae53ecb85fe240
    End
    RecordSource ="PAY_GETD"
    Caption ="PAY_GETD"
    OnCurrent ="[Event Procedure]"
    BeforeInsert ="[Event Procedure]"
    OnDelete ="[Event Procedure]"
    BeforeUpdate ="[Event Procedure]"
    AfterUpdate ="[Event Procedure]"
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
            Height =0
            BackColor =-2147483633
            Name ="FormHeader"
        End
        Begin Section
            Height =8176
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1020
                    ColumnWidth =1110
                    ColumnOrder =1
                    Name ="N_SERI"
                    ControlSource ="N_SERI"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =3587
                            Width =1005
                            Height =240
                            Name ="Label1"
                            Caption ="شماره سريال"
                        End
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1020
                    Top =680
                    ColumnWidth =990
                    ColumnOrder =0
                    TabIndex =1
                    Name ="DATE"
                    ControlSource ="DATE"
                    DefaultValue ="=FARSIDATE(Date())"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3707
                            Top =680
                            Width =885
                            Height =240
                            Name ="Label5"
                            Caption ="تاريخ دريافت"
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
                    Left =1020
                    Top =1020
                    Height =255
                    ColumnWidth =1155
                    ColumnOrder =4
                    TabIndex =2
                    Name ="DATE_S"
                    ControlSource ="DATE_S"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3542
                            Top =1020
                            Width =1050
                            Height =240
                            Name ="Label7"
                            Caption ="تاريخ سررسيد"
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
                    Left =1020
                    Top =1701
                    ColumnWidth =1500
                    ColumnOrder =5
                    TabIndex =3
                    Name ="MABL"
                    ControlSource ="MABL"
                    Format ="Standard"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4217
                            Top =1701
                            Width =375
                            Height =240
                            Name ="Label11"
                            Caption ="مبلغ"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =110
                    Top =2041
                    Width =2580
                    ColumnWidth =2355
                    ColumnOrder =6
                    TabIndex =4
                    Name ="NAME_TAH"
                    ControlSource ="NAME_TAH"
                    DefaultValue ="=gettafname([Forms]![HEAD_LST_khadamat]![CUST_NO])"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3386
                            Top =2047
                            Width =1215
                            Height =240
                            Name ="Label13"
                            Caption ="نام پرداخت كننده"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1020
                    Top =2381
                    ColumnOrder =7
                    TabIndex =5
                    Name ="N_HESAB"
                    ControlSource ="N_HESAB"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =3557
                            Top =2381
                            Width =1035
                            Height =240
                            Name ="Label15"
                            Caption ="جاري چك"
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
                    Left =1020
                    Top =5102
                    ColumnOrder =13
                    TabIndex =6
                    Name ="N_S"
                    ControlSource ="N_S"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3707
                            Top =5102
                            Width =885
                            Height =240
                            Name ="Label31"
                            Caption ="شماره سند"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1020
                    Top =5442
                    ColumnOrder =12
                    TabIndex =7
                    Name ="RADIF"
                    ControlSource ="RADIF"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =4182
                            Top =5437
                            Width =675
                            Height =405
                            Name ="Label33"
                            Caption ="رديف دفتر"
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
                    Left =1020
                    Top =5783
                    ColumnOrder =14
                    TabIndex =8
                    Name ="NUMBER"
                    ControlSource ="NUMBER"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =3632
                            Top =5783
                            Width =960
                            Height =240
                            Name ="Label35"
                            Caption ="شماره فاكتور"
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
                    Left =1020
                    Top =6123
                    ColumnOrder =15
                    TabIndex =9
                    Name ="TAG"
                    ControlSource ="TAG"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3962
                            Top =6123
                            Width =630
                            Height =240
                            Name ="Label37"
                            Caption ="برچسب"
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
                    Left =1020
                    Top =6463
                    ColumnOrder =16
                    TabIndex =10
                    Name ="ANBAR"
                    ControlSource ="ANBAR"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4262
                            Top =6463
                            Width =330
                            Height =240
                            Name ="Label39"
                            Caption ="انبار"
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
                    Left =1020
                    Top =6803
                    ColumnOrder =17
                    TabIndex =11
                    Name ="CUST_NO"
                    ControlSource ="CUST_NO"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3467
                            Top =6803
                            Width =1125
                            Height =240
                            Name ="Label41"
                            Caption ="شماره مشتري"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =1020
                    Top =340
                    Width =786
                    ColumnOrder =2
                    TabIndex =12
                    Name ="BANK"
                    ControlSource ="BANK"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TCOD_BANKS.CODE, TCOD_BANKS.NAMES FROM TCOD_BANKS ORDER BY TCOD_BANKS.NAM"
                        "ES;"
                    ColumnWidths ="0"
                    ValidationRule ="Not Is Null"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4217
                            Top =340
                            Width =375
                            Height =240
                            Name ="Label3"
                            Caption ="بانك"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =113
                    Top =3968
                    Width =3621
                    ColumnWidth =1320
                    ColumnOrder =9
                    TabIndex =13
                    Name ="N_KOL"
                    ControlSource ="N_KOL"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT     NUMBER, NAME FROM TOTA_HES WHERE (NUMBER = 112)ORDER BY NAME"
                    ColumnWidths ="567;2835"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3920
                            Top =3969
                            Width =960
                            Height =240
                            Name ="Label17"
                            Caption ="به حساب كل"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =1020
                    Top =3081
                    ColumnOrder =10
                    TabIndex =14
                    Name ="N_MOIN"
                    ControlSource ="N_MOIN"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT     DETA_HES.NUMBER, DETA_HES.NAME FROM DETA_HES WHERE     (((DETA_HES.N_"
                        "KOL) =)) GROUP BY DETA_HES.NUMBER, DETA_HES.NAME ORDER BY DETA_HES.NAME"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3642
                            Top =3087
                            Width =1230
                            Height =240
                            Name ="Label19"
                            Caption ="به حساب  معين"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =247
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =226
                    Top =2551
                    ColumnWidth =3570
                    ColumnOrder =11
                    TabIndex =15
                    Name ="N_TAF"
                    ControlSource ="N_TAF"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TDETA_HES.TNUMBER, TDETA_HES.NAME FROM TDETA_HES WHERE (((TDETA_HES.NUMBE"
                        "R) =) AND ((TDETA_HES.N_KOL) =))GROUP BY TDETA_HES.TNUMBER, TDETA_HES.NAME ORDER"
                        " BY TDETA_HES.NAME"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =2848
                            Top =2557
                            Width =1230
                            Height =240
                            Name ="Label45"
                            Caption ="به حساب  تفصيلي"
                        End
                    End
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1011
                    Top =1361
                    Width =1710
                    Height =255
                    ColumnWidth =2475
                    ColumnOrder =3
                    TabIndex =16
                    Name ="SHOBEH"
                    ControlSource ="SHOBEH"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT PAY_GETD.SHOBEH FROM PAY_GETD GROUP BY PAY_GETD.SHOBEH ORDER BY PAY_GETD."
                        "SHOBEH;"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3872
                            Top =1361
                            Width =720
                            Height =240
                            Name ="Label9"
                            Caption ="نام شعبه"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =793
                    Top =7426
                    Width =2211
                    Height =315
                    ColumnOrder =18
                    TabIndex =17
                    Name ="VAZ"
                    ControlSource ="VAZ"
                    RowSourceType ="Value List"
                    RowSource ="1;\"نزد صندوق\";2;\"نزد بانك\";3;\"وصول شده\";4;\"واگذار شده\";5;\"برگشت شده\";6"
                        ";\"مسترد شده\""
                    ColumnWidths ="0"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="1"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =3135
                            Top =7426
                            Width =1005
                            Height =315
                            Name ="Label50"
                            Caption ="وضعيت چك*"
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
                    Left =793
                    Top =7861
                    Width =2211
                    Height =315
                    ColumnWidth =3885
                    ColumnOrder =19
                    TabIndex =18
                    Name ="SANDUGH"
                    ControlSource ="SANDUGH"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TNUMBER, NAME FROM TDETA_HES WHERE (N_KOL = 113) AND (NUMBER = 1)"
                    ColumnWidths ="0"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="1"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =3135
                            Top =7861
                            Width =1005
                            Height =315
                            Name ="Label40"
                            Caption ="موقعيت چك*"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =247
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Width =3111
                    Height =315
                    ColumnWidth =2880
                    ColumnOrder =8
                    TabIndex =19
                    Name ="SAYADI"
                    ControlSource ="SAYADI"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextFontCharSet =0
                            Left =3220
                            Top =2
                            Width =1095
                            Height =240
                            Name ="Label78"
                            Caption ="شماره صيادي"
                        End
                    End
                End
            End
        End
        Begin FormFooter
            Height =1134
            BackColor =-2147483633
            Name ="FormFooter"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Name ="SCMAB"
                    ControlSource ="=Sum([MABL])"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3197
                            Width =375
                            Height =240
                            Name ="Label43"
                            Caption ="مبلغ"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "PAY_GETD_KHAD_SUB.cls"
