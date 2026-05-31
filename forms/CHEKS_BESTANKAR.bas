Version =20
VersionRequired =20
Begin Form
    AllowDeletions = NotDefault
    AllowAdditions = NotDefault
    KeyPreview = NotDefault
    AllowEdits = NotDefault
    DefaultView =2
    AllowUpdating =4
    ViewsAllowed =2
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =5100
    RowHeight =300
    ItemSuffix =31
    Left =1020
    Top =600
    Right =12960
    Bottom =5790
    DatasheetGridlinesColor =12632256
    MaxRecords =100000
    RecSrcDt = Begin
        0xdb1800b01135e440
    End
    RecordSource ="SELECT PAY_GETD.*, TCOD_BANKS.NAMES FROM PAY_GETD INNER JOIN TCOD_BANKS ON PAY_G"
        "ETD.BANK = TCOD_BANKS.CODE WHERE (PAY_GETD.N_KOL <> 911) AND (PAY_GETD.N_KOL <> "
        "112) OR (NOT (PAY_GETD.N_KOL2 IS NULL)) OR (NOT (PAY_GETD.N_KOL3 IS NULL)) ORDER"
        " BY PAY_GETD.N_SERI"
    Caption ="چكهاي وصول شده واگذار شده يا برگشت شده"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    HelpFile ="dena.hlp"
    DatasheetFontName ="Tahoma"
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
        Begin PageHeader
            DisplayWhen =1
            Height =0
            BackColor =-2147483633
            Name ="PageHeader"
        End
        Begin Section
            Height =6730
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =93
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1233
                    Top =113
                    Height =255
                    ColumnWidth =1095
                    ColumnOrder =1
                    Name ="N_SERI"
                    ControlSource ="N_SERI"
                    Format ="Fixed"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3800
                            Top =113
                            Width =1005
                            Height =240
                            Name ="Label1"
                            Caption ="شماره سريال"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1233
                    Top =793
                    Height =255
                    ColumnWidth =1125
                    ColumnOrder =3
                    TabIndex =1
                    Name ="DATE_S"
                    ControlSource ="DATE_S"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3755
                            Top =793
                            Width =1050
                            Height =240
                            Name ="Label5"
                            Caption ="تاريخ سررسيد"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1233
                    Top =1133
                    Height =255
                    ColumnWidth =1110
                    ColumnOrder =4
                    TabIndex =2
                    Name ="DATE"
                    ControlSource ="DATE"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3920
                            Top =1133
                            Width =885
                            Height =240
                            Name ="Label7"
                            Caption ="تاريخ دريافت"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1224
                    Top =1474
                    Width =1710
                    Height =255
                    ColumnWidth =2025
                    ColumnOrder =6
                    TabIndex =3
                    Name ="SHOBEH"
                    ControlSource ="SHOBEH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4085
                            Top =1474
                            Width =720
                            Height =240
                            Name ="Label9"
                            Caption ="نام شعبه"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1233
                    Top =1814
                    Height =255
                    ColumnWidth =1305
                    ColumnOrder =2
                    TabIndex =4
                    Name ="MABL"
                    ControlSource ="MABL"
                    Format ="Standard"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4430
                            Top =1814
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
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =354
                    Top =2154
                    Width =2580
                    Height =255
                    ColumnWidth =2925
                    ColumnOrder =7
                    TabIndex =5
                    Name ="NAME_TAH"
                    ControlSource ="NAME_TAH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3630
                            Top =2160
                            Width =1260
                            Height =240
                            Name ="Label13"
                            Caption ="نام  پرداخت كننده"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1233
                    Top =2494
                    Height =255
                    ColumnOrder =8
                    TabIndex =6
                    Name ="N_HESAB"
                    ControlSource ="N_HESAB"
                    Format ="Fixed"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3770
                            Top =2494
                            Width =1035
                            Height =240
                            Name ="Label15"
                            Caption ="شماره حساب"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1233
                    Top =2834
                    Height =255
                    ColumnWidth =885
                    ColumnOrder =9
                    TabIndex =7
                    Name ="N_S"
                    ControlSource ="N_S"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3920
                            Top =2834
                            Width =885
                            Height =240
                            Name ="Label17"
                            Caption ="شماره سند"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =144
                    Top =3174
                    Width =2790
                    Height =255
                    ColumnWidth =780
                    ColumnOrder =5
                    TabIndex =8
                    Name ="NAMES"
                    ControlSource ="NAMES"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4190
                            Top =3174
                            Width =615
                            Height =240
                            Name ="Label19"
                            Caption ="نام بانك"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1245
                    Top =453
                    Height =255
                    ColumnWidth =645
                    ColumnOrder =0
                    TabIndex =9
                    Name ="RADIF"
                    ControlSource ="RADIF"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4177
                            Top =450
                            Width =630
                            Height =240
                            Name ="Label20"
                            Caption ="شماره "
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =400
                    Top =3549
                    ColumnOrder =11
                    TabIndex =10
                    Name ="N_KOL"
                    ControlSource ="N_KOL"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2965
                            Top =3549
                            Width =1005
                            Height =240
                            Name ="Label22"
                            Caption =" به حساب كل"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =400
                    Top =3889
                    ColumnOrder =12
                    TabIndex =11
                    Name ="N_MOIN"
                    ControlSource ="N_MOIN"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2800
                            Top =3894
                            Width =1125
                            Height =240
                            Name ="Label23"
                            Caption ="به حساب معين"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =400
                    Top =4229
                    ColumnWidth =1545
                    ColumnOrder =14
                    TabIndex =12
                    Name ="N_KOL2"
                    ControlSource ="N_KOL2"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =2530
                            Top =4224
                            Width =1665
                            Height =405
                            Name ="Label24"
                            Caption ="برگشت به حساب كل"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =400
                    Top =4569
                    ColumnWidth =1710
                    ColumnOrder =15
                    TabIndex =13
                    Name ="N_MOIN2"
                    ControlSource ="N_MOIN2"
                    Begin
                        Begin Label
                            OverlapFlags =223
                            Left =2665
                            Top =4569
                            Width =1545
                            Height =405
                            Name ="Label25"
                            Caption ="برگشت به حساب معين"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =400
                    Top =4909
                    ColumnWidth =1725
                    ColumnOrder =17
                    TabIndex =14
                    Name ="N_KOL3"
                    ControlSource ="N_KOL3"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =2830
                            Top =4914
                            Width =1395
                            Height =240
                            Name ="Label26"
                            Caption ="وصول به حساب كل"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =400
                    Top =5249
                    ColumnWidth =1680
                    ColumnOrder =18
                    TabIndex =15
                    Name ="N_MOIN3"
                    ControlSource ="N_MOIN3"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2665
                            Top =5244
                            Width =1560
                            Height =240
                            Name ="Label27"
                            Caption ="وصول به حساب معين"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =144
                    Top =5809
                    ColumnWidth =1860
                    ColumnOrder =16
                    TabIndex =16
                    Name ="N_TAF2"
                    ControlSource ="N_TAF2"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2353
                            Top =5814
                            Width =1905
                            Height =240
                            Name ="Label35"
                            Caption ="برگشت به حساب تفصيلي"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =144
                    Top =6150
                    ColumnWidth =1830
                    ColumnOrder =19
                    TabIndex =17
                    Name ="N_TAF3"
                    ControlSource ="N_TAF3"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2353
                            Top =6144
                            Width =1755
                            Height =240
                            Name ="Label36"
                            Caption ="وصول به حساب تفصيلي"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =144
                    Top =6490
                    ColumnWidth =1335
                    ColumnOrder =13
                    TabIndex =18
                    Name ="N_TAF"
                    ControlSource ="N_TAF"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2353
                            Top =6489
                            Width =1320
                            Height =240
                            Name ="Label37"
                            Caption ="به حساب تفصيلي"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =247
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =226
                    Top =1417
                    Height =255
                    ColumnWidth =690
                    TabIndex =19
                    Name ="VAZ"
                    ControlSource ="VAZ"
                    Format ="Fixed"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =2793
                            Top =1417
                            Width =1005
                            Height =240
                            Name ="Label29"
                            Caption ="صندوق"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OldBorderStyle =0
                    OverlapFlags =247
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Width =1581
                    Height =300
                    TabIndex =20
                    Name ="KIND"
                    ControlSource ="KIND"
                    RowSourceType ="Value List"
                    RowSource ="1;\"تجاري\";0;\"غير تجاري\""
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Width =1005
                            Height =240
                            Name ="Label30"
                            Caption ="نوع"
                        End
                    End
                End
            End
        End
        Begin PageFooter
            DisplayWhen =1
            Height =1134
            BackColor =-2147483633
            Name ="PageFooter"
        End
    End
End
CodeBehindForm
' See "CHEKS_BESTANKAR.cls"
