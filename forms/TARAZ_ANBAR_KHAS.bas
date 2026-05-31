Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    DividingLines = NotDefault
    FilterOn = NotDefault
    DefaultView =2
    AllowUpdating =4
    ViewsAllowed =2
    TabularCharSet =178
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =8957
    RowHeight =300
    ItemSuffix =37
    Left =2175
    Top =1095
    Right =13155
    Bottom =6045
    DatasheetForeColor =33554432
    DatasheetGridlinesColor =12632256
    MaxRecords =100000
    Tag ="@FORMS___F_MENU_ANBAR_TARAZ___DT2=[FORMS]![F_MENU_ANBAR_TARAZ]![DT2]"
    RecSrcDt = Begin
        0xa4c60eee8a4be340
    End
    InputParameters ="@FORMS___F_MENU_ANBAR_TARAZ___DT2=[FORMS]![F_MENU_ANBAR_TARAZ]![DT2],@ANB=[FORMS"
        "]![F_MENU_ANBAR_TARAZ]![ANBAR]"
    RecordSource ="dbo.TARAZ_ANBAR_KHAS"
    Caption ="تراز موجودي انبار"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Tahoma"
    OnError ="[Event Procedure]"
    UnknownProp = {258 ,66 ,12 ,4 ,220 } Begin
        0x400046004f0052004d0053005f005f005f0046005f004d0045004e0055005f00 ,
        0x41004e004200410052005f0054004100520041005a005f005f005f0044005400 ,
        0x32003d005b0046004f0052004d0053005d0021005b0046005f004d0045004e00 ,
        0x55005f0041004e004200410052005f0054004100520041005a005d0021005b00 ,
        0x4400540032005d002c00400041004e0042003d005b0046004f0052004d005300 ,
        0x5d0021005b0046005f004d0045004e0055005f0041004e004200410052005f00 ,
        0x54004100520041005a005d0021005b0041004e004200410052005d00
    End
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
        Begin Section
            Height =4662
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5194
                    Top =113
                    Width =1440
                    ColumnWidth =900
                    ColumnOrder =0
                    Name ="CODE"
                    ControlSource ="CODE"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =8010
                            Top =113
                            Width =495
                            Height =240
                            Name ="Label1"
                            Caption ="كد كالا"
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
                    Left =4933
                    Top =453
                    ColumnWidth =915
                    ColumnOrder =2
                    TabIndex =1
                    Name ="MEG"
                    ControlSource ="MEG"
                    Format ="Fixed"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =8115
                            Top =453
                            Width =390
                            Height =240
                            Name ="Label3"
                            Caption ="مقدار اوليه"
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
                    Left =4933
                    Top =793
                    ColumnWidth =1590
                    ColumnOrder =4
                    TabIndex =2
                    Name ="SumOfMABL_A"
                    ControlSource ="SumOfMABL_A"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =7365
                            Top =793
                            Width =1140
                            Height =240
                            Name ="Label5"
                            Caption ="مبلغ اوليه"
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
                    Left =4933
                    Top =1133
                    ColumnWidth =1095
                    ColumnOrder =5
                    TabIndex =3
                    Name ="MEGHKH"
                    ControlSource ="MEGHKH"
                    Format ="Fixed"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =7815
                            Top =1133
                            Width =990
                            Height =240
                            Name ="Label7"
                            Caption ="مقدار افزايش"
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
                    Left =4933
                    Top =1474
                    ColumnWidth =1455
                    ColumnOrder =7
                    TabIndex =4
                    Name ="MABKH"
                    ControlSource ="MABKH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =7920
                            Top =1474
                            Width =915
                            Height =240
                            Name ="Label9"
                            Caption ="مبلغ افزايش"
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
                    Left =4933
                    Top =1814
                    ColumnWidth =1050
                    ColumnOrder =8
                    TabIndex =5
                    Name ="MEGFR"
                    ControlSource ="MEGFR"
                    Format ="Fixed"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =7920
                            Top =1814
                            Width =945
                            Height =240
                            Name ="Label11"
                            Caption ="مقدار كاهش"
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
                    Left =4933
                    Top =2154
                    ColumnWidth =1380
                    ColumnOrder =10
                    TabIndex =6
                    Name ="MABFR"
                    ControlSource ="MABFR"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =7920
                            Top =2154
                            Width =870
                            Height =240
                            Name ="Label13"
                            Caption ="مبلغ كاهش"
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
                    Left =4933
                    Top =2494
                    ColumnWidth =1170
                    ColumnOrder =11
                    TabIndex =7
                    Name ="MEGMA"
                    ControlSource ="=[MEG]+[MEGHKH]-[MEGFR]"
                    Format ="Fixed"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =7890
                            Top =2494
                            Width =1065
                            Height =240
                            Name ="Label15"
                            Caption ="مقدار موجودي"
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
                    Left =4933
                    Top =2834
                    ColumnWidth =1410
                    ColumnOrder =13
                    TabIndex =8
                    Name ="MABMA"
                    ControlSource ="=Fix([MABKH]+[SumOfMABL_A]-[MABFR])"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =7890
                            Top =2834
                            Width =990
                            Height =240
                            Name ="Label17"
                            Caption ="مبلغ موجودي"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3184
                    Top =3174
                    Width =3450
                    ColumnWidth =5325
                    ColumnOrder =1
                    TabIndex =9
                    Name ="NAME"
                    ControlSource ="NAME"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =8010
                            Top =3174
                            Width =540
                            Height =240
                            Name ="Label19"
                            Caption ="نام كالا"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3115
                    Top =3540
                    Width =3450
                    ColumnWidth =2175
                    ColumnOrder =17
                    TabIndex =10
                    Name ="NAMES"
                    ControlSource ="NAMES"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =7936
                            Top =3546
                            Width =570
                            Height =240
                            Name ="Label21"
                            Caption ="نام انبار"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =247
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3833
                    Width =1440
                    ColumnWidth =615
                    ColumnOrder =14
                    TabIndex =11
                    Name ="Text22"
                    ControlSource ="VCOD"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6649
                            Width =495
                            Height =240
                            Name ="Label23"
                            Caption ="كد كالا"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =5364
                    Top =3855
                    ColumnWidth =645
                    ColumnOrder =15
                    TabIndex =12
                    Name ="ANBAR"
                    ControlSource ="ANBAR"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =8166
                            Top =3855
                            Width =600
                            Height =240
                            Name ="Label25"
                            Caption ="كد انبار"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =3890
                    Top =4422
                    ColumnOrder =16
                    TabIndex =13
                    Name ="VAHED"
                    ControlSource ="VAHED"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6662
                            Top =4422
                            Width =630
                            Height =240
                            Name ="Label26"
                            Caption ="واحد"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =113
                    Top =737
                    ColumnWidth =1170
                    ColumnOrder =6
                    TabIndex =14
                    Name ="Text27"
                    ControlSource ="=[MABKH]/IIf([MEGHKH]=0,1,[MEGHKH])"
                    Format ="#,##0.00;-#,##0.00"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3100
                            Top =737
                            Width =915
                            Height =240
                            Name ="Label28"
                            Caption ="في افزايش"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =113
                    Top =1097
                    ColumnWidth =1170
                    ColumnOrder =9
                    TabIndex =15
                    Name ="Text29"
                    ControlSource ="=[MABFR]/IIf([MEGFR]=0,1,[MEGFR])"
                    Format ="#,##0.00;-#,##0.00"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3100
                            Top =1097
                            Width =915
                            Height =240
                            Name ="Label30"
                            Caption ="في كاهش"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =113
                    Top =1457
                    ColumnWidth =1170
                    ColumnOrder =12
                    TabIndex =16
                    Name ="Text31"
                    ControlSource ="=Fix([MABKH]+[SumOfMABL_A]-[MABFR])/IIf(([MEG]+[MEGHKH]-[MEGFR])=0,1,([MEG]+[MEG"
                        "HKH]-[MEGFR]))"
                    Format ="#,##0.00;-#,##0.00"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3105
                            Top =1455
                            Width =930
                            Height =240
                            Name ="Label32"
                            Caption ="في موجودي"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =113
                    Top =1817
                    ColumnWidth =1170
                    ColumnOrder =3
                    TabIndex =17
                    Name ="Text33"
                    ControlSource ="=[SumOfMABL_A]/IIf([MEG]=0,1,[MEG])"
                    Format ="#,##0.00;-#,##0.00"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3100
                            Top =1817
                            Width =915
                            Height =240
                            Name ="Label34"
                            Caption ="في اوليه"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =283
                    Top =3174
                    TabIndex =18
                    Name ="grname"
                    ControlSource ="grname"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =2740
                            Top =3174
                            Width =945
                            Height =240
                            Name ="Label35"
                            Caption ="نام گروه كالا"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =283
                    Top =4195
                    ColumnWidth =690
                    TabIndex =19
                    Name ="RADAH"
                    ControlSource ="RADAH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3025
                            Top =4195
                            Width =660
                            Height =240
                            Name ="Label36"
                            Caption ="كد گروه"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "TARAZ_ANBAR_KHAS.cls"
