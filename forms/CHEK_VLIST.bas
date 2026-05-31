Version =20
VersionRequired =20
Begin Form
    AutoCenter = NotDefault
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
    Width =5669
    RowHeight =315
    ItemSuffix =4
    Left =435
    Top =750
    Right =15165
    Bottom =7890
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xfe28af235dade240
    End
    InputParameters ="@forms___baseknow___bankha=forms![baseknow]![bankha]"
    RecordSource ="dbo.CHKE_VLIST"
    Caption ="ليست چكهاي دريافتي"
    DatasheetFontName ="Tahoma"
    OnKeyPress ="[Event Procedure]"
    OnLoad ="[Event Procedure]"
    OnError ="[Event Procedure]"
    UnknownProp = {258 ,66 ,12 ,4 ,104 } Begin
        0x400066006f0072006d0073005f005f005f0062006100730065006b006e006f00 ,
        0x77005f005f005f00620061006e006b00680061003d0066006f0072006d007300 ,
        0x21005b0062006100730065006b006e006f0077005d0021005b00620061006e00 ,
        0x6b00680061005d00
    End
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
            Height =6589
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
                    ColumnOrder =0
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
                    ColumnOrder =1
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
                    ColumnOrder =2
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
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1224
                    Top =1474
                    Width =1710
                    Height =255
                    ColumnOrder =3
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
                    ColumnOrder =4
                    TabIndex =4
                    Name ="MABL"
                    ControlSource ="MABL"
                    Format ="#,###"
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
                    ColumnWidth =3690
                    ColumnOrder =5
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
                    ColumnOrder =6
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
                    ColumnOrder =7
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
                    ColumnOrder =8
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
                    OverlapFlags =93
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1245
                    Top =453
                    Height =255
                    ColumnOrder =9
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
                    Left =1275
                    Top =3465
                    ColumnOrder =10
                    TabIndex =10
                    Name ="N_KOL"
                    ControlSource ="N_KOL"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3840
                            Top =3465
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
                    Left =1275
                    Top =3805
                    ColumnOrder =11
                    TabIndex =11
                    Name ="N_MOIN"
                    ControlSource ="N_MOIN"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3675
                            Top =3810
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
                    Left =1303
                    Top =4483
                    ColumnWidth =1590
                    ColumnOrder =13
                    TabIndex =12
                    Name ="N_KOL2"
                    ControlSource ="N_KOL2"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =3433
                            Top =4478
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
                    Left =1303
                    Top =4823
                    ColumnWidth =1755
                    ColumnOrder =14
                    TabIndex =13
                    Name ="N_MOIN2"
                    ControlSource ="N_MOIN2"
                    Begin
                        Begin Label
                            OverlapFlags =223
                            Left =3568
                            Top =4823
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
                    Left =1303
                    Top =5163
                    ColumnWidth =1560
                    ColumnOrder =16
                    TabIndex =14
                    Name ="N_KOL3"
                    ControlSource ="N_KOL3"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =3733
                            Top =5168
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
                    Left =1303
                    Top =5503
                    ColumnWidth =1725
                    ColumnOrder =17
                    TabIndex =15
                    Name ="N_MOIN3"
                    ControlSource ="N_MOIN3"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3568
                            Top =5498
                            Width =1560
                            Height =240
                            Name ="Label27"
                            Caption ="وصول به حساب معين"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    ColumnHidden = NotDefault
                    OverlapFlags =93
                    NumeralShapes =2
                    Width =288
                    Height =288
                    ColumnWidth =540
                    ColumnOrder =18
                    TabIndex =16
                    Name ="NNS"
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1275
                    Top =4170
                    ColumnWidth =1365
                    ColumnOrder =12
                    TabIndex =17
                    Name ="N_TAF"
                    ControlSource ="N_TAF"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3675
                            Top =4170
                            Width =1320
                            Height =240
                            Name ="Label29"
                            Caption ="به حساب تفصيلي"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1305
                    Top =5865
                    ColumnWidth =1905
                    ColumnOrder =15
                    TabIndex =18
                    Name ="N_TAF2"
                    ControlSource ="N_TAF2"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3225
                            Top =5865
                            Width =1890
                            Height =240
                            Name ="Label31"
                            Caption ="برگشت به حساب تفصيلي"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1305
                    Top =6225
                    ColumnWidth =1875
                    ColumnOrder =20
                    TabIndex =19
                    Name ="N_TAF3"
                    ControlSource ="N_TAF3"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =3225
                            Top =6225
                            Width =1890
                            Height =240
                            Name ="Label33"
                            Caption ="وصول به حساب تفصيلي"
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
                    Top =453
                    Height =255
                    TabIndex =20
                    Name ="VAZ"
                    ControlSource ="VAZ"
                    Format ="Fixed"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =2567
                            Top =453
                            Width =1005
                            Height =240
                            Name ="Label2"
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
                    TabIndex =21
                    Name ="KIND"
                    ControlSource ="KIND"
                    RowSourceType ="Value List"
                    RowSource ="1;\"تجاري\";0;\"غير تجاري\""
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =255
                            Width =1005
                            Height =240
                            Name ="Label30"
                            Caption ="نوع"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =247
                    IMESentenceMode =3
                    Left =1417
                    Top =6349
                    TabIndex =22
                    Name ="BANK"
                    ControlSource ="BANK"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =4414
                            Top =6349
                            Width =405
                            Height =240
                            Name ="Label3"
                            Caption ="بانك:"
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
                    Name ="smabl"
                    ControlSource ="=Sum([MABL])"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3197
                            Width =375
                            Height =240
                            Name ="Label21"
                            Caption ="مبلغ"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "CHEK_VLIST.cls"
