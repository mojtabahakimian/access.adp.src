Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    DividingLines = NotDefault
    DefaultView =0
    AllowUpdating =4
    ScrollBars =0
    TabularCharSet =178
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =14683
    RowHeight =390
    ItemSuffix =41
    Left =150
    Top =15
    Right =14580
    Bottom =7785
    DatasheetForeColor =33554432
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xbfb13f513b17e540
    End
    RecordSource ="SELECT PRGHEAD.PRGID, PRGHEAD.PRG_DATE, PRGHEAD.PCODE, STUF_DEF.NAME FROM PRGHEA"
        "D INNER JOIN STUF_DEF ON PRGHEAD.PCODE = STUF_DEF.CODE"
    Caption ="فرمول توليد محصول"
    DatasheetFontName ="Tahoma"
    OnError ="[Event Procedure]"
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
            Height =9240
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =11961
                    Top =60
                    Width =1401
                    Height =315
                    ColumnWidth =1320
                    ColumnOrder =0
                    ForeColor =8388608
                    Name ="PRGID"
                    ControlSource ="PRGID"
                    ControlTipText ="با دابل كليك روي شماره فرمول سرفصهايي كه تعريف نشده اند براي حسابداري صنعتي تعري"
                        "ف ميشوند"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =13492
                            Top =60
                            Width =1140
                            Height =300
                            Name ="Label1"
                            Caption ="شماره :"
                            ControlTipText ="با دابل كليك روي شماره فرمول سرفصهايي كه تعريف نشده اند براي حسابداري صنعتي تعري"
                                "ف ميشوند"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =11946
                    Top =895
                    Width =1416
                    Height =315
                    ColumnWidth =1005
                    ColumnOrder =2
                    TabIndex =2
                    ForeColor =8388608
                    Name ="PRG_DATE"
                    ControlSource ="PRG_DATE"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =13521
                            Top =895
                            Width =1050
                            Height =300
                            Name ="Label5"
                            Caption ="تاريخ اعمال:"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5524
                    Top =895
                    Width =1072
                    Height =315
                    ColumnWidth =1875
                    ColumnOrder =3
                    TabIndex =5
                    BackColor =13434828
                    ForeColor =8388608
                    Name ="IMBIBE_MANF"
                    ControlSource ="=[Forms]![HEAD_MANF]![DASTR]"
                    DefaultValue ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =6706
                            Top =895
                            Width =1815
                            Height =300
                            Name ="Label7"
                            Caption ="جذب هزينه دستمزد:"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2430
                    Top =878
                    Width =1596
                    Height =315
                    ColumnWidth =1695
                    ColumnOrder =4
                    TabIndex =6
                    BackColor =13434828
                    ForeColor =8388608
                    Name ="IMBIBE_SAR"
                    ControlSource ="=[Forms]![HEAD_MANF]![SARR]"
                    DefaultValue ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =4115
                            Top =878
                            Width =1335
                            Height =240
                            Name ="Label9"
                            Caption ="جذب هزينه سربار:"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =3
                    Left =8781
                    Top =420
                    Width =4575
                    Height =315
                    ColumnWidth =3375
                    ColumnOrder =1
                    TabIndex =1
                    ForeColor =8388608
                    Name ="CODE"
                    ControlSource ="PCODE"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT CODE, NAME + N' ' + CODE AS name, CODE AS Expr1, RADAH FROM STUF_DEF WHER"
                        "E (RADAH > 1) ORDER BY NAME + N' ' + CODE"
                    ColumnWidths ="0;2268;567"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =13506
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
                    Left =172
                    Top =1365
                    Width =14385
                    Height =6570
                    TabIndex =7
                    Name ="DTL_MANF_SUB1_DAY"
                    SourceObject ="Form.DTL_MANF_SUB1_DAY"
                    LinkChildFields ="PRGID"
                    LinkMasterFields ="PRGID"
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
                    Left =7240
                    Top =8004
                    Width =1596
                    Height =315
                    ColumnOrder =5
                    TabIndex =8
                    BackColor =16777164
                    Name ="Text14"
                    ControlSource ="=[Forms]![HEAD_MANF]![MAVADR]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =8923
                            Top =8010
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
                    Left =2715
                    Top =7995
                    Width =2421
                    Height =675
                    ColumnOrder =6
                    FontSize =16
                    FontWeight =700
                    TabIndex =9
                    BackColor =13434828
                    Name ="Text16"
                    ControlSource ="=[Forms]![HEAD_MANF]![MAVADR]+[Forms]![HEAD_MANF]![DASTR]+[Forms]![HEAD_MANF]![S"
                        "ARR]"
                    Format ="#,###"
                    FontName ="B Titr"
                    Tag ="=[DTL_MANF_SUB1_DAY].[Form]![SMABK]+[IMBIBE_MANF]+[IMBIBE_SAR]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =5220
                            Top =7995
                            Width =1905
                            Height =300
                            Name ="Label17"
                            Caption ="قيمت تمام شده استاندارد:"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =93
                    TextFontCharSet =0
                    Left =3174
                    Top =170
                    Width =1140
                    Height =465
                    FontWeight =700
                    TabIndex =10
                    ForeColor =255
                    Name ="Command19"
                    Caption ="چاپ "
                    OnClick ="[Event Procedure]"
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =5524
                    Top =60
                    Width =1071
                    Height =330
                    ColumnWidth =1560
                    ColumnOrder =9
                    TabIndex =3
                    Name ="SA_HOUR"
                    DefaultValue ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6706
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
                    Left =5524
                    Top =430
                    Width =1071
                    Height =330
                    ColumnWidth =2220
                    ColumnOrder =10
                    TabIndex =4
                    Name ="SA_NHOU"
                    DefaultValue ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6721
                            Top =430
                            Width =1950
                            Height =240
                            Name ="Label27"
                            Caption ="نرخ ساعت كار استاندارد:"
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
                    Left =163
                    Top =56
                    Width =2226
                    Height =1020
                    ColumnWidth =5400
                    ColumnOrder =8
                    TabIndex =11
                    ForeColor =8388608
                    Name ="TOZIH"
                    ControlTipText ="هرگونه توضيحي راجع به فرمول"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextFontCharSet =0
                            Left =2478
                            Top =56
                            Width =735
                            Height =240
                            Name ="Label33"
                            Caption ="توضيحات:"
                        End
                    End
                End
                Begin OptionGroup
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =8565
                    Top =878
                    Width =3227
                    Height =388
                    TabIndex =12
                    Name ="Frame34"
                    DefaultValue ="1"
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =215
                            Left =11240
                            Top =765
                            Width =495
                            Height =240
                            BackColor =-2147483633
                            Name ="Label35"
                            Caption ="قيمت"
                        End
                        Begin OptionButton
                            OverlapFlags =87
                            ReadingOrder =1
                            Left =9555
                            Top =953
                            OptionValue =2
                            Name ="Option39"
                            Begin
                                Begin Label
                                    OverlapFlags =119
                                    Left =8730
                                    Top =925
                                    Width =825
                                    Height =240
                                    Name ="Label40"
                                    Caption ="آخرين خريد"
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =87
                            ReadingOrder =1
                            Left =10950
                            Top =983
                            OptionValue =1
                            Name ="Option37"
                            Begin
                                Begin Label
                                    OverlapFlags =119
                                    Left =9915
                                    Top =955
                                    Width =1035
                                    Height =240
                                    Name ="Label38"
                                    Caption ="ميانگين موزون"
                                End
                            End
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "HEAD_MANF_DAY.cls"
