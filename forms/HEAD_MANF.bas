Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    DividingLines = NotDefault
    FilterOn = NotDefault
    OrderByOn = NotDefault
    DefaultView =0
    AllowUpdating =4
    ScrollBars =0
    TabularCharSet =178
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =15307
    RowHeight =390
    ItemSuffix =39
    Left =4170
    Top =420
    Right =19215
    Bottom =9315
    DatasheetForeColor =33554432
    DatasheetGridlinesColor =12632256
    Filter ="((GHEYMAT=12))"
    OrderBy ="TOZIH"
    RecSrcDt = Begin
        0x874c6fa3d2fee440
    End
    RecordSource ="SELECT HEAD_MANF.* FROM HEAD_MANF WHERE (NOT (CODE IS NULL)) ORDER BY FNUMB"
    Caption ="فرمول توليد محصول"
    OnCurrent ="[Event Procedure]"
    OnDelete ="[Event Procedure]"
    BeforeUpdate ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Tahoma"
    OnLoad ="[Event Procedure]"
    OnError ="[Event Procedure]"
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
        Begin Subform
            SpecialEffect =2
            Width =1701
            Height =1701
        End
        Begin Section
            CanGrow = NotDefault
            Height =8900
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =12923
                    Top =60
                    Width =1056
                    Height =315
                    ColumnWidth =1320
                    ColumnOrder =0
                    BackColor =16777164
                    ForeColor =8388608
                    Name ="FNUMB"
                    ControlSource ="FNUMB"
                    OnDblClick ="[Event Procedure]"
                    ControlTipText ="با دابل كليك روي شماره فرمول سرفصهايي كه تعريف نشده اند براي حسابداري صنعتي تعري"
                        "ف ميشوند"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =14109
                            Top =60
                            Width =1140
                            Height =300
                            Name ="Label1"
                            Caption ="شماره فرمول:"
                            ControlTipText ="با دابل كليك روي شماره فرمول سرفصهايي كه تعريف نشده اند براي حسابداري صنعتي تعري"
                                "ف ميشوند"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =13043
                    Top =810
                    Width =936
                    Height =315
                    ColumnWidth =1005
                    ColumnOrder =2
                    TabIndex =2
                    BackColor =13434828
                    ForeColor =8388608
                    Name ="DATE_ACTIV"
                    ControlSource ="DATE_ACTIV"
                    BeforeUpdate ="[Event Procedure]"
                    DefaultValue ="=FARSIDATE(Date())"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =14138
                            Top =810
                            Width =1050
                            Height =300
                            Name ="Label5"
                            Caption ="تاريخ اعمال:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =6887
                    Top =810
                    Width =1086
                    Height =315
                    ColumnWidth =1875
                    ColumnOrder =3
                    TabIndex =5
                    BackColor =13434828
                    ForeColor =8388608
                    Name ="IMBIBE_MANF"
                    ControlSource ="IMBIBE_MANF"
                    OnKeyPress ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =8069
                            Top =810
                            Width =1815
                            Height =300
                            Name ="Label7"
                            Caption ="جذب هزينه دستمزد:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3793
                    Top =793
                    Width =1596
                    Height =315
                    ColumnWidth =1695
                    ColumnOrder =4
                    TabIndex =6
                    BackColor =13434828
                    ForeColor =8388608
                    Name ="IMBIBE_SAR"
                    ControlSource ="IMBIBE_SAR"
                    OnKeyPress ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =5478
                            Top =793
                            Width =1335
                            Height =240
                            Name ="Label9"
                            Caption ="جذب هزينه سربار:"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =3
                    Left =9923
                    Top =420
                    Width =4050
                    Height =315
                    ColumnWidth =5610
                    ColumnOrder =1
                    TabIndex =1
                    BackColor =13434828
                    ForeColor =8388608
                    Name ="CODE"
                    ControlSource ="CODE"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT CODE, NAME + N' ' + CODE AS name, CODE AS Expr1, RADAH FROM STUF_DEF WHER"
                        "E (RADAH > 1) ORDER BY NAME + N' ' + CODE"
                    ColumnWidths ="0;2268;567"
                    AfterUpdate ="[Event Procedure]"
                    OnNotInList ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =14123
                            Top =420
                            Width =435
                            Height =300
                            Name ="Label11"
                            Caption =" كالا:"
                        End
                    End
                End
                Begin Subform
                    OverlapFlags =85
                    Left =54
                    Top =1230
                    Width =15120
                    Height =6870
                    TabIndex =7
                    Name ="DTL_MANF_SUB1"
                    SourceObject ="Form.DTL_MANF_SUB1"
                    LinkChildFields ="FNUMB"
                    LinkMasterFields ="FNUMB"
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =7857
                    Top =8169
                    Width =1596
                    Height =315
                    ColumnOrder =5
                    TabIndex =8
                    BackColor =16777164
                    Name ="Text14"
                    ControlSource ="=[DTL_MANF_SUB1].[Form]![SKOL]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =9540
                            Top =8175
                            Width =1845
                            Height =300
                            Name ="Label15"
                            Caption ="جمع مواد اوليه مصرفي:"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2792
                    Top =8160
                    Width =2961
                    Height =630
                    ColumnOrder =6
                    FontSize =14
                    FontWeight =700
                    TabIndex =9
                    BackColor =13434828
                    Name ="Text16"
                    ControlSource ="=[DTL_MANF_SUB1].[Form]![SMABK]+[IMBIBE_MANF]+[IMBIBE_SAR]"
                    Format ="#,###"
                    FontName ="B Titr"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =5837
                            Top =8160
                            Width =1905
                            Height =300
                            Name ="Label17"
                            Caption ="قيمت تمام شده استاندارد:"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =11655
                    Top =8225
                    Width =861
                    ColumnWidth =1020
                    ColumnOrder =7
                    TabIndex =10
                    Name ="GHEYMAT"
                    ControlSource ="GHEYMAT"
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontCharSet =0
                    Left =3827
                    Top =60
                    Width =1140
                    Height =330
                    FontWeight =700
                    TabIndex =11
                    ForeColor =255
                    Name ="Command19"
                    Caption ="چاپ فرمول"
                    OnClick ="[Event Procedure]"
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =6887
                    Top =60
                    Width =1071
                    Height =330
                    ColumnWidth =1560
                    ColumnOrder =9
                    TabIndex =3
                    BackColor =16777164
                    Name ="SA_HOUR"
                    ControlSource ="SA_HOUR"
                    AfterUpdate ="[Event Procedure]"
                    OnKeyPress ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =8069
                            Top =60
                            Width =1710
                            Height =240
                            Name ="Label26"
                            Caption ="جذب ساعت كار:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =6887
                    Top =430
                    Width =1071
                    Height =330
                    ColumnWidth =2220
                    ColumnOrder =10
                    TabIndex =4
                    BackColor =16777164
                    Name ="SA_NHOU"
                    ControlSource ="SA_NHOU"
                    AfterUpdate ="[Event Procedure]"
                    OnKeyPress ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =8084
                            Top =430
                            Width =1950
                            Height =240
                            Name ="Label27"
                            Caption ="نرخ ساعت كار استاندارد:"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontCharSet =0
                    Left =4989
                    Top =64
                    Width =1785
                    Height =330
                    FontWeight =700
                    TabIndex =12
                    ForeColor =255
                    Name ="Command29"
                    Caption ="قيمت تمام شده روز"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontCharSet =0
                    Left =3834
                    Top =409
                    Width =1140
                    Height =330
                    FontWeight =700
                    TabIndex =13
                    ForeColor =255
                    Name ="Command30"
                    Caption ="چاپ فرمول2"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontCharSet =0
                    Left =4989
                    Top =409
                    Width =1785
                    Height =330
                    FontWeight =700
                    TabIndex =14
                    ForeColor =255
                    Name ="Command31"
                    Caption ="ليست فرمولها و مواد"
                    OnClick ="[Event Procedure]"
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =135
                    Top =60
                    Width =2226
                    Height =1020
                    ColumnWidth =5400
                    ColumnOrder =8
                    TabIndex =15
                    BackColor =16764057
                    ForeColor =8388608
                    Name ="TOZIH"
                    ControlSource ="TOZIH"
                    ControlTipText ="هرگونه توضيحي راجع به فرمول"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =2450
                            Top =60
                            Width =735
                            Height =240
                            Name ="Label33"
                            Caption ="توضيحات:"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =93
                    TextFontCharSet =0
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =14321
                    Top =8214
                    Width =861
                    TabIndex =16
                    Name ="MAVADR"
                    DefaultValue ="0"
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =247
                    TextFontCharSet =0
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =13754
                    Top =8214
                    Width =861
                    TabIndex =17
                    Name ="DASTR"
                    DefaultValue ="0"
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =12734
                    Top =8214
                    Width =861
                    TabIndex =18
                    Name ="SARR"
                    DefaultValue ="0"
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    Left =2664
                    Top =623
                    Width =735
                    Height =465
                    FontWeight =700
                    TabIndex =19
                    Name ="ESLAH"
                    Caption ="اصلاح "
                    OnClick ="[Event Procedure]"
                End
                Begin ComboBox
                    RowSourceTypeInt =1
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =9977
                    Top =850
                    Width =2601
                    Height =255
                    TabIndex =20
                    Name ="Text37"
                    ControlSource ="GHEYMAT"
                    RowSourceType ="Value List"
                    RowSource ="1;\"فروردين\";2;\"ارديبهشت\";3;\"خرداد\";4;\"تير\";5;\"مرداد\";6;\"شهريور\";7;\""
                        "مهر\";8;\"آبان\";9;\"آذر\";10;\"دي\";11;\"بهمن\";12;\"اسفند\""
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =12629
                            Top =850
                            Width =345
                            Height =240
                            Name ="Label38"
                            Caption ="ماه"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "HEAD_MANF.cls"
