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
    Width =5062
    RowHeight =285
    ItemSuffix =37
    Left =450
    Top =780
    Right =12195
    Bottom =5340
    HelpContextId =8006
    DatasheetGridlinesColor =12632256
    Tag =" @Forms___Baseknow___BANKHA=[FORMS]![BASEKNOW]![BANKHA]"
    RecSrcDt = Begin
        0x2f9b130e8e86e440
    End
    InputParameters =" @Forms___Baseknow___BANKHA=[FORMS]![BASEKNOW]![BANKHA]"
    RecordSource ="CHEK_VPLS"
    Caption ="ليست چكهاي پرداختي  اعلام وصول نشده"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    HelpFile ="dena.hlp"
    DatasheetFontName ="Tahoma"
    OnKeyPress ="[Event Procedure]"
    OnLoad ="[Event Procedure]"
    OnError ="[Event Procedure]"
    UnknownProp = {258 ,66 ,12 ,4 ,110 } Begin
        0x2000400046006f0072006d0073005f005f005f0042006100730065006b006e00 ,
        0x6f0077005f005f005f00420041004e004b00480041003d005b0046004f005200 ,
        0x4d0053005d0021005b0042004100530045004b004e004f0057005d0021005b00 ,
        0x420041004e004b00480041005d00
    End
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
            Height =8640
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
                    ColumnWidth =1500
                    ColumnOrder =2
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
                    ColumnOrder =4
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
                    ColumnOrder =5
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
                    ColumnWidth =4035
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
                    ColumnWidth =3255
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
                    OverlapFlags =93
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =226
                    Top =3458
                    Width =2790
                    ColumnOrder =3
                    TabIndex =8
                    Name ="NAMES"
                    ControlSource ="NAMES"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =4272
                            Top =3458
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
                    Visible = NotDefault
                    ColumnHidden = NotDefault
                    OverlapFlags =93
                    NumeralShapes =2
                    Width =288
                    Height =288
                    ColumnOrder =11
                    TabIndex =10
                    Name ="NNS"
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =170
                    Top =3855
                    ColumnOrder =12
                    TabIndex =11
                    Name ="N_KOL"
                    ControlSource ="N_KOL"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2735
                            Top =3855
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
                    Top =4150
                    ColumnOrder =13
                    TabIndex =12
                    Name ="N_MOIN"
                    ControlSource ="N_MOIN"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2400
                            Top =4155
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
                    Left =28
                    Top =4828
                    ColumnWidth =1545
                    ColumnOrder =15
                    TabIndex =13
                    Name ="N_KOL2"
                    ControlSource ="N_KOL2"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =2158
                            Top =4823
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
                    Left =28
                    Top =5168
                    ColumnWidth =1710
                    ColumnOrder =16
                    TabIndex =14
                    Name ="N_MOIN2"
                    ControlSource ="N_MOIN2"
                    Begin
                        Begin Label
                            OverlapFlags =223
                            Left =2293
                            Top =5168
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
                    Left =28
                    Top =5508
                    ColumnWidth =1515
                    ColumnOrder =18
                    TabIndex =15
                    Name ="N_KOL3"
                    ControlSource ="N_KOL3"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =2458
                            Top =5513
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
                    Left =28
                    Top =5848
                    ColumnWidth =1680
                    ColumnOrder =19
                    TabIndex =16
                    Name ="N_MOIN3"
                    ControlSource ="N_MOIN3"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2293
                            Top =5843
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
                    Top =4515
                    ColumnWidth =1335
                    ColumnOrder =14
                    TabIndex =17
                    Name ="N_TAF"
                    ControlSource ="N_TAF"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2400
                            Top =4515
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
                    Left =30
                    Top =6210
                    ColumnWidth =1860
                    ColumnOrder =17
                    TabIndex =18
                    Name ="N_TAF2"
                    ControlSource ="N_TAF2"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1950
                            Top =6210
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
                    Left =30
                    Top =6570
                    ColumnWidth =1785
                    ColumnOrder =20
                    TabIndex =19
                    Name ="N_TAF3"
                    ControlSource ="N_TAF3"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1950
                            Top =6570
                            Width =1890
                            Height =240
                            Name ="Label33"
                            Caption ="وصول ازحساب تفصيلي"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OldBorderStyle =0
                    OverlapFlags =255
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Width =1581
                    Height =300
                    ColumnOrder =21
                    TabIndex =20
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
                    ColumnWidth =6000
                    ColumnOrder =22
                    TabIndex =21
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
                    ColumnWidth =4380
                    ColumnOrder =23
                    TabIndex =22
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
                    ColumnWidth =5760
                    ColumnOrder =24
                    TabIndex =23
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
                    OverlapFlags =87
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1133
                    Top =3231
                    ColumnOrder =25
                    TabIndex =24
                    Name ="BANK"
                    ControlSource ="BANK"
                    Begin
                        Begin Label
                            OverlapFlags =87
                            Left =3890
                            Top =3231
                            Width =645
                            Height =240
                            Name ="Label35"
                            Caption ="كد بانك"
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
                    ColumnWidth =1935
                    ColumnOrder =7
                    TabIndex =25
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
' See "CHECK_PVLIST.cls"
