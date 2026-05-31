Version =20
VersionRequired =20
Begin Form
    AllowDeletions = NotDefault
    AllowAdditions = NotDefault
    AllowEdits = NotDefault
    DefaultView =2
    AllowUpdating =4
    ViewsAllowed =2
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =5062
    RowHeight =345
    ItemSuffix =24
    Left =465
    Top =750
    Right =14940
    Bottom =9600
    DatasheetGridlinesColor =12632256
    MaxRecords =10000000
    RecSrcDt = Begin
        0xb7ed2cf8b08be540
    End
    RecordSource ="CHEK_PLIST"
    Caption ="ليست چكهاي پرداختي"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
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
            Height =9070
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
                    ColumnWidth =1095
                    ColumnOrder =0
                    Name ="N_SERI"
                    ControlSource ="N_SERI"
                    Format ="Fixed"
                    Begin
                        Begin Label
                            OverlapFlags =93
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
                    ColumnWidth =1125
                    ColumnOrder =2
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
                    ColumnWidth =1110
                    ColumnOrder =3
                    TabIndex =2
                    Name ="DATE"
                    ControlSource ="DATE"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3915
                            Top =1140
                            Width =930
                            Height =240
                            Name ="Label7"
                            Caption ="تاريخ پرداخت"
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
                    ColumnOrder =5
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
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1233
                    Top =1814
                    ColumnOrder =1
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
                    ColumnWidth =3960
                    ColumnOrder =6
                    TabIndex =5
                    Name ="NAME_TAH"
                    ControlSource ="NAME_TAH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3635
                            Top =2154
                            Width =1170
                            Height =240
                            Name ="Label13"
                            Caption ="نام تحويل گيرنده"
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
                    ColumnOrder =7
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
                    ColumnOrder =4
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
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1233
                    Top =473
                    ColumnOrder =10
                    TabIndex =9
                    Name ="RADIF"
                    ControlSource ="RADIF"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3800
                            Top =473
                            Width =1005
                            Height =240
                            Name ="Label23"
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
                    Left =736
                    Top =3514
                    ColumnOrder =11
                    TabIndex =10
                    Name ="N_KOL"
                    ControlSource ="N_KOL"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3301
                            Top =3514
                            Width =1005
                            Height =240
                            Name ="Label22"
                            Caption ="از حساب كل"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =566
                    Top =3809
                    ColumnOrder =12
                    TabIndex =11
                    Name ="N_MOIN"
                    ControlSource ="N_MOIN"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2966
                            Top =3814
                            Width =1125
                            Height =240
                            Name ="Label31"
                            Caption ="از حساب معين"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =594
                    Top =4487
                    ColumnWidth =1545
                    ColumnOrder =14
                    TabIndex =12
                    Name ="N_KOL2"
                    ControlSource ="N_KOL2"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =2724
                            Top =4482
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
                    Left =594
                    Top =4827
                    ColumnWidth =1710
                    ColumnOrder =15
                    TabIndex =13
                    Name ="N_MOIN2"
                    ControlSource ="N_MOIN2"
                    Begin
                        Begin Label
                            OverlapFlags =223
                            Left =2859
                            Top =4827
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
                    Left =594
                    Top =5167
                    ColumnWidth =1515
                    ColumnOrder =17
                    TabIndex =14
                    Name ="N_KOL3"
                    ControlSource ="N_KOL3"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =3024
                            Top =5172
                            Width =1395
                            Height =240
                            Name ="Label26"
                            Caption ="وصول از حساب كل"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =594
                    Top =5507
                    ColumnWidth =1680
                    ColumnOrder =18
                    TabIndex =15
                    Name ="N_MOIN3"
                    ControlSource ="N_MOIN3"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2859
                            Top =5502
                            Width =1560
                            Height =240
                            Name ="Label27"
                            Caption ="وصول از حساب معين"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =566
                    Top =4174
                    ColumnOrder =13
                    TabIndex =16
                    Name ="N_TAF"
                    ControlSource ="N_TAF"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2966
                            Top =4174
                            Width =1320
                            Height =240
                            Name ="Label29"
                            Caption ="از حساب تفصيلي"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =596
                    Top =5869
                    ColumnWidth =1860
                    ColumnOrder =16
                    TabIndex =17
                    Name ="N_TAF2"
                    ControlSource ="N_TAF2"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2516
                            Top =5869
                            Width =1890
                            Height =240
                            Name ="Label32"
                            Caption ="برگشت به حساب تفصيلي"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =596
                    Top =6229
                    ColumnWidth =1785
                    ColumnOrder =19
                    TabIndex =18
                    Name ="N_TAF3"
                    ControlSource ="N_TAF3"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2516
                            Top =6229
                            Width =1890
                            Height =240
                            Name ="Label33"
                            Caption ="وصول ازحساب تفصيلي"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    Locked = NotDefault
                    RowSourceTypeInt =1
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =623
                    Top =6576
                    ColumnOrder =20
                    TabIndex =19
                    Name ="VAZ"
                    ControlSource ="VAZ"
                    RowSourceType ="Value List"
                    RowSource ="1;\"نزد  شخص\";2;\"عودت شده\""
                    ColumnWidths ="0"
                    OnDblClick ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2543
                            Top =6576
                            Width =1890
                            Height =240
                            Name ="Label28"
                            Caption ="وضعيت چك"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =1
                    NumeralShapes =1
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =30
                    Top =7035
                    ColumnWidth =6285
                    ColumnOrder =21
                    TabIndex =20
                    Name ="bhes"
                    ControlSource ="HES1"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT hes, hes + N' : ' + NAME AS Expr1 FROM CUST_HESAB"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2595
                            Top =7035
                            Width =1320
                            Height =240
                            Name ="Label39"
                            Caption =" پرداخت از حساب "
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =1
                    NumeralShapes =1
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =30
                    Top =7715
                    ColumnWidth =4065
                    ColumnOrder =23
                    TabIndex =21
                    Name ="brhes"
                    ControlSource ="HES2"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT hes, hes + N' : ' + NAME AS Expr1 FROM CUST_HESAB"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2160
                            Top =7710
                            Width =1665
                            Height =405
                            Name ="Label45"
                            Caption ="برگشت به حساب "
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =1
                    NumeralShapes =1
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =30
                    Top =8395
                    ColumnWidth =4455
                    ColumnOrder =22
                    TabIndex =22
                    Name ="vshes"
                    ControlSource ="HES3"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT hes, hes + N' : ' + NAME AS Expr1 FROM CUST_HESAB"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2460
                            Top =8400
                            Width =1395
                            Height =240
                            Name ="Label43"
                            Caption ="وصول به حساب "
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
                    ColumnWidth =2160
                    ColumnOrder =8
                    TabIndex =23
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
                    Name ="smabl"
                    ControlSource ="=Sum([MABL])"
                    Format ="#,###"
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
' See "CHEK_PLIST.cls"
