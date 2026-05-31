Version =20
VersionRequired =20
Begin Form
    DefaultView =2
    AllowUpdating =4
    ViewsAllowed =2
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =5669
    RowHeight =285
    Left =405
    Top =3840
    Right =11400
    Bottom =6570
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x6fa4bb1d6d88e240
    End
    RecordSource ="PAY_GETP"
    OnCurrent ="[Event Procedure]"
    BeforeInsert ="[Event Procedure]"
    OnDelete ="[Event Procedure]"
    AfterUpdate ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Tahoma"
    OnError ="[Event Procedure]"
    RecordSourceQualifier ="dbo"
    Begin
        Begin Label
            BackStyle =0
            TextFontCharSet =178
            FontName ="Tahoma"
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
        Begin FormHeader
            Height =0
            BackColor =-2147483633
            Name ="FormHeader"
        End
        Begin Section
            Height =7156
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1023
                    Top =113
                    ColumnWidth =1095
                    ColumnOrder =1
                    Name ="N_SERI"
                    ControlSource ="N_SERI"
                    Format ="Fixed"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3590
                            Top =113
                            Width =1005
                            Height =240
                            Name ="Label1"
                            Caption ="شماره سريال"
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
                    Left =1023
                    Top =793
                    ColumnWidth =1125
                    ColumnOrder =0
                    TabIndex =1
                    Name ="DATE"
                    ControlSource ="DATE"
                    BeforeUpdate ="[Event Procedure]"
                    DefaultValue ="farsidate(Date())"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3710
                            Top =793
                            Width =885
                            Height =240
                            Name ="Label5"
                            Caption ="تاريخ پرداخت"
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
                    Left =1023
                    Top =1133
                    ColumnWidth =1125
                    ColumnOrder =4
                    TabIndex =2
                    Name ="DATE_S"
                    ControlSource ="DATE_S"
                    BeforeUpdate ="[Event Procedure]"
                    OnExit ="[Event Procedure]"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3545
                            Top =1133
                            Width =1050
                            Height =240
                            Name ="Label7"
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
                    Left =1014
                    Top =1474
                    Width =1710
                    ColumnOrder =3
                    TabIndex =3
                    Name ="SHOBEH"
                    ControlSource ="SHOBEH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3875
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
                    Left =1023
                    Top =1814
                    ColumnOrder =5
                    TabIndex =4
                    Name ="MABL"
                    ControlSource ="MABL"
                    Format ="Fixed"
                    OnKeyPress ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4220
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
                    Left =144
                    Top =2154
                    Width =2580
                    ColumnWidth =2325
                    ColumnOrder =6
                    TabIndex =5
                    Name ="NAME_TAH"
                    ControlSource ="NAME_TAH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3420
                            Top =2160
                            Width =1215
                            Height =240
                            Name ="Label13"
                            Caption ="نام دريافت كننده"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1023
                    Top =2494
                    ColumnWidth =1050
                    ColumnOrder =7
                    TabIndex =6
                    Name ="N_HESAB"
                    ControlSource ="N_HESAB"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3560
                            Top =2494
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
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1023
                    Top =2834
                    ColumnOrder =8
                    TabIndex =7
                    Name ="N_KOL"
                    ControlSource ="N_KOL"
                    DefaultValue ="=[Forms]![Baseknow]![BANKHA]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3815
                            Top =2834
                            Width =780
                            Height =240
                            Name ="Label17"
                            Caption ="حساب كل"
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
                    Left =1023
                    Top =5215
                    ColumnOrder =9
                    TabIndex =8
                    Name ="N_S"
                    ControlSource ="N_S"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3710
                            Top =5215
                            Width =885
                            Height =240
                            Name ="Label31"
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
                    Left =1023
                    Top =5555
                    ColumnWidth =885
                    ColumnOrder =16
                    TabIndex =9
                    Name ="RADIF"
                    ControlSource ="RADIF"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =4185
                            Top =5550
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
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1023
                    Top =5896
                    ColumnOrder =10
                    TabIndex =10
                    Name ="NUMBER"
                    ControlSource ="NUMBER"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =3635
                            Top =5896
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
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1023
                    Top =6236
                    ColumnOrder =11
                    TabIndex =11
                    Name ="TAG"
                    ControlSource ="TAG"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3965
                            Top =6236
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
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1023
                    Top =6576
                    ColumnOrder =12
                    TabIndex =12
                    Name ="ANBAR"
                    ControlSource ="ANBAR"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4265
                            Top =6576
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
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1023
                    Top =6916
                    ColumnWidth =1155
                    ColumnOrder =13
                    TabIndex =13
                    Name ="CUST_NO"
                    ControlSource ="CUST_NO"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3470
                            Top =6916
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
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =1023
                    Top =453
                    Width =786
                    ColumnOrder =2
                    TabIndex =14
                    Name ="BANK"
                    ControlSource ="BANK"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TCOD_BANKS.CODE, TCOD_BANKS.NAMES FROM TCOD_BANKS ORDER BY TCOD_BANKS.NAM"
                        "ES;"
                    ColumnWidths ="0"
                    AfterUpdate ="[Event Procedure]"
                    OnNotInList ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4220
                            Top =453
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
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =1023
                    Top =3174
                    ColumnWidth =1755
                    ColumnOrder =14
                    TabIndex =15
                    Name ="N_MOIN"
                    ControlSource ="N_MOIN"
                    RowSourceType ="Table/View/StoredProc"
                    ColumnWidths ="0"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3645
                            Top =3180
                            Width =1230
                            Height =240
                            Name ="Label19"
                            Caption ="پرداخت ازحساب  معين"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =1077
                    Top =3968
                    ColumnWidth =3045
                    ColumnOrder =15
                    TabIndex =16
                    Name ="N_TAF"
                    ControlSource ="N_TAF"
                    RowSourceType ="Table/View/StoredProc"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3699
                            Top =3974
                            Width =1230
                            Height =240
                            Name ="Label45"
                            Caption ="پرداخت ازحساب تفصيلي"
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
' See "PAY_GETP_SND_SUB.cls"
