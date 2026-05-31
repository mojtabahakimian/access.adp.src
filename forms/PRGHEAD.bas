Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    DividingLines = NotDefault
    OrderByOn = NotDefault
    DefaultView =0
    AllowUpdating =4
    ScrollBars =0
    ViewsAllowed =1
    TabularCharSet =178
    TabularFamily =42
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =15307
    DatasheetFontHeight =10
    ItemSuffix =60
    Left =2970
    Top =780
    Right =18270
    Bottom =9390
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xf58c6aa66ea8e240
    End
    RecordSource ="PRGHEAD"
    Caption ="برنامه ريزي مواد اوليه"
    OnCurrent ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial (Arabic)"
    OnLoad ="[Event Procedure]"
    OnError ="[Event Procedure]"
    AllowDatasheetView =0
    RecordSourceQualifier ="dbo"
    Begin
        Begin Label
            BackStyle =0
            TextFontCharSet =178
            FontName ="Tahoma"
        End
        Begin Line
            Width =1701
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
        Begin Tab
            TextFontCharSet =178
            Width =5103
            Height =3402
            FontName ="Tahoma"
        End
        Begin Page
            Width =1701
            Height =1701
        End
        Begin Section
            CanGrow = NotDefault
            Height =9765
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =12260
                    Top =82
                    Height =315
                    Name ="PRGID"
                    ControlSource ="PRGID"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            TextAlign =1
                            Left =14049
                            Top =89
                            Width =1035
                            Height =300
                            Name ="Label1"
                            Caption ="شماره برنامه"
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
                    Left =12256
                    Top =450
                    Height =315
                    TabIndex =1
                    Name ="PRG_DATE"
                    ControlSource ="PRG_DATE"
                    BeforeUpdate ="[Event Procedure]"
                    DefaultValue ="=FARSIDATE(Date())"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            TextAlign =1
                            Left =14045
                            Top =450
                            Width =990
                            Height =300
                            Name ="Label3"
                            Caption ="تاريخ برنامه"
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
                    Left =9659
                    Top =454
                    Height =315
                    TabIndex =3
                    Name ="NUM"
                    ControlSource ="NUM"
                    ValidationRule ="Not Is Null"
                    OnKeyPress ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            TextAlign =1
                            Left =11448
                            Top =454
                            Width =735
                            Height =300
                            Name ="Label7"
                            Caption ="تعداد كالا"
                        End
                    End
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =3
                    Left =6969
                    Top =86
                    Width =4395
                    Height =315
                    TabIndex =2
                    BackColor =13434828
                    Name ="PCODE"
                    ControlSource ="PCODE"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT CODE, NAME + N' ' + CODE AS nam, CODE AS Expr1 FROM STUF_DEF WHERE (RADAH"
                        " <> 1)"
                    ColumnWidths ="0;3402;567"
                    OnNotInList ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            TextAlign =1
                            Left =11452
                            Top =86
                            Width =540
                            Height =300
                            Name ="Label5"
                            Caption ="كد كالا"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontCharSet =0
                    Left =170
                    Top =56
                    Width =1365
                    Height =360
                    FontWeight =700
                    TabIndex =4
                    Name ="Command22"
                    Caption ="شروع محاسبه"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Find Next"
                End
                Begin Tab
                    OverlapFlags =85
                    TextFontCharSet =0
                    Left =60
                    Top =855
                    Width =15240
                    Height =8865
                    FontWeight =700
                    TabIndex =6
                    Name ="TabCtl10"
                    Begin
                        Begin Page
                            OverlapFlags =215
                            Left =194
                            Top =1260
                            Width =14971
                            Height =8325
                            Name ="Page11"
                            Caption ="اجزاء تشكيل دهنده"
                            Begin
                                Begin Subform
                                    OverlapFlags =247
                                    Left =194
                                    Top =1320
                                    Width =14700
                                    Height =7530
                                    Name ="PROGPAS_SUB"
                                    SourceObject ="Form.PROGPAS_SUB"
                                    LinkChildFields ="PRGID"
                                    LinkMasterFields ="PRGID"
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =215
                            Left =194
                            Top =1260
                            Width =14971
                            Height =8325
                            Name ="Page12"
                            Caption ="خالص مواد اوليه مصرفي"
                            Begin
                                Begin Subform
                                    OverlapFlags =247
                                    Left =194
                                    Top =1320
                                    Width =14700
                                    Height =7530
                                    Name ="Child13"
                                    SourceObject ="Form.QPROGPAS2_SUB"
                                    LinkChildFields ="PRGID"
                                    LinkMasterFields ="PRGID"
                                End
                                Begin TextBox
                                    DecimalPlaces =0
                                    OverlapFlags =215
                                    TextFontCharSet =0
                                    TextAlign =3
                                    NumeralShapes =2
                                    Left =2610
                                    Top =9240
                                    Width =2931
                                    Height =315
                                    TabIndex =1
                                    BackColor =13434828
                                    Name ="Text14"
                                    ControlSource ="=[Child13].[Form]![SMABLK]*[Forms]![PRGHEAD]![NUM]"
                                    Format ="#,###"
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            TextFontCharSet =0
                                            TextAlign =1
                                            Left =5632
                                            Top =9243
                                            Width =1695
                                            Height =300
                                            FontWeight =700
                                            Name ="Label15"
                                            Caption ="جمع مواد مصرفي"
                                        End
                                    End
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =247
                            Left =195
                            Top =1260
                            Width =14970
                            Height =8325
                            Name ="Page31"
                            Caption ="سفارش مواد اوليه و موجودي"
                            Begin
                                Begin Subform
                                    OverlapFlags =247
                                    Left =224
                                    Top =1320
                                    Width =14670
                                    Height =7530
                                    Name ="Child32"
                                    SourceObject ="Form.QPROGPAS3_SUB"
                                    LinkChildFields ="PRGID"
                                    LinkMasterFields ="PRGID"
                                End
                                Begin TextBox
                                    DecimalPlaces =0
                                    OverlapFlags =247
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =1
                                    KeyboardLanguage =43
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =2493
                                    Top =9240
                                    Width =2496
                                    Height =315
                                    FontWeight =700
                                    TabIndex =1
                                    BackColor =13434828
                                    Name ="Text33"
                                    ControlSource ="=[Child13].[Form]![SKOLMAV]"
                                    Format ="#,###"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            TextAlign =1
                                            Left =5080
                                            Top =9243
                                            Width =1785
                                            Height =300
                                            FontWeight =700
                                            Name ="Label34"
                                            Caption ="جمع مواد مصرفي"
                                        End
                                    End
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =247
                            Left =195
                            Top =1260
                            Width =14970
                            Height =8325
                            Name ="Page16"
                            Caption ="جذب هزينه تبديل"
                            Begin
                                Begin Subform
                                    OverlapFlags =247
                                    Left =3974
                                    Top =1320
                                    Width =11070
                                    Height =8265
                                    Name ="Child17"
                                    SourceObject ="Form.QPASED3"
                                    LinkChildFields ="PRGID"
                                    LinkMasterFields ="PRGID"
                                End
                                Begin TextBox
                                    OverlapFlags =247
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =1
                                    KeyboardLanguage =43
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =340
                                    Top =2518
                                    Width =1476
                                    Height =315
                                    TabIndex =1
                                    BackColor =13434828
                                    Name ="SIMBIBE_MANF1"
                                    ControlSource ="=[Child17].[Form]![SIMBIBE_MANF1]*[num]"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            Left =1945
                                            Top =2518
                                            Width =1845
                                            Height =300
                                            Name ="Label21"
                                            Caption ="هزينه دستمزد  :"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =247
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =1
                                    KeyboardLanguage =43
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =342
                                    Top =3287
                                    Width =1476
                                    Height =315
                                    TabIndex =2
                                    BackColor =13434828
                                    Name ="SIMBIBE_SAR1"
                                    ControlSource ="=[Child17].[Form]![SIMBIBE_SAR1]*[num]"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            Left =1947
                                            Top =3292
                                            Width =1785
                                            Height =300
                                            Name ="Label23"
                                            Caption =" هزينه سربار  :"
                                        End
                                    End
                                End
                                Begin Line
                                    OverlapFlags =119
                                    Left =356
                                    Top =4485
                                    Width =3393
                                    Name ="Line24"
                                End
                                Begin TextBox
                                    SpecialEffect =0
                                    OverlapFlags =247
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =1
                                    KeyboardLanguage =43
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =356
                                    Top =4560
                                    Width =1476
                                    Height =315
                                    TabIndex =3
                                    Name ="SIMBIBE_MANF"
                                    ControlSource ="=[SIMBIBE_SAR1]+[sm]"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            Left =1961
                                            Top =4565
                                            Width =1785
                                            Height =300
                                            Name ="Label26"
                                            Caption ="جذب هزينه تبديل :"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =247
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =1
                                    KeyboardLanguage =43
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =341
                                    Top =5451
                                    Width =1476
                                    Height =315
                                    TabIndex =4
                                    Name ="Text44"
                                    ControlSource ="=([Child17].[Form]![sanieh1]+[hm])*[num]"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            Left =1945
                                            Top =5456
                                            Width =1860
                                            Height =405
                                            Name ="Label45"
                                            Caption ="زمان توليد كل:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =247
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =1
                                    KeyboardLanguage =43
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =341
                                    Top =5055
                                    Width =1476
                                    Height =315
                                    TabIndex =5
                                    BackColor =13434828
                                    Name ="hm"
                                    DefaultValue ="0"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            Left =1946
                                            Top =5060
                                            Width =1785
                                            Height =300
                                            Name ="Label47"
                                            Caption ="زمان توليد خو دكالا:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =247
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =1
                                    KeyboardLanguage =43
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =342
                                    Top =2097
                                    Width =1476
                                    Height =315
                                    TabIndex =6
                                    BackColor =13434828
                                    Name ="dm"
                                    DefaultValue ="0"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            Left =1947
                                            Top =2097
                                            Width =1845
                                            Height =300
                                            Name ="Label49"
                                            Caption ="هزينه دستمزد  خود كالا:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =247
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =1
                                    KeyboardLanguage =43
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =342
                                    Top =3748
                                    Width =1476
                                    Height =315
                                    TabIndex =7
                                    BackColor =13434828
                                    Name ="sm"
                                    DefaultValue ="0"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            Left =1947
                                            Top =3753
                                            Width =1785
                                            Height =300
                                            Name ="Label51"
                                            Caption =" هزينه سربار خود كالا :"
                                        End
                                    End
                                End
                                Begin TextBox
                                    SpecialEffect =0
                                    OverlapFlags =247
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =1
                                    KeyboardLanguage =43
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =342
                                    Top =2891
                                    Width =1476
                                    Height =315
                                    TabIndex =8
                                    Name ="Text55"
                                    ControlSource ="=[SIMBIBE_MANF1]+[dm]"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            Left =1947
                                            Top =2896
                                            Width =1785
                                            Height =300
                                            Name ="Label56"
                                            Caption ="جذب هزينه دستمزد :"
                                        End
                                    End
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =247
                            Left =194
                            Top =1260
                            Width =14971
                            Height =8325
                            Name ="Page35"
                            Caption ="وضعيت گروه برنامه"
                            Begin
                                Begin Subform
                                    OverlapFlags =247
                                    Left =194
                                    Top =1320
                                    Width =14700
                                    Height =7530
                                    Name ="QPROGPAS3_SUB2"
                                    SourceObject ="Form.QPROGPAS3_SUB2"
                                    LinkChildFields ="GRP"
                                    LinkMasterFields ="GRP"
                                End
                                Begin Subform
                                    OverlapFlags =247
                                    Left =8745
                                    Top =9075
                                    Width =6225
                                    Height =495
                                    TabIndex =1
                                    Name ="SARBAR"
                                    SourceObject ="Form.SARBAR"
                                    LinkChildFields ="GRP"
                                    LinkMasterFields ="GRP"
                                End
                            End
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
                    Left =7009
                    Top =450
                    Height =315
                    TabIndex =7
                    Name ="GRP"
                    ControlSource ="GRP"
                    ValidationRule ="Not Is Null"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            TextAlign =1
                            Left =8764
                            Top =450
                            Width =810
                            Height =285
                            Name ="Label40"
                            Caption ="گروه برنامه"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontCharSet =0
                    Left =165
                    Top =420
                    Width =1365
                    Height =360
                    FontWeight =700
                    TabIndex =5
                    Name ="Command41"
                    Caption ="حذف اجزاء"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Find Next"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =3
                    Left =1644
                    Top =113
                    Width =4311
                    Height =317
                    TabIndex =8
                    BackColor =16777164
                    Name ="formula"
                    ControlSource ="formula"
                    RowSourceType ="Table/View/StoredProc"
                    ColumnWidths ="0;3402;567"
                    ValidationRule ="Not Is Null"
                    OnEnter ="[Event Procedure]"
                    OnExit ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6011
                            Top =113
                            Width =945
                            Height =300
                            Name ="Label57"
                            Caption ="فرمول"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontCharSet =0
                    Left =1923
                    Top =453
                    Width =1365
                    Height =360
                    FontWeight =700
                    TabIndex =9
                    Name ="Command58"
                    Caption ="شروع تجزيه"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Find Next"
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontCharSet =0
                    Left =3397
                    Top =453
                    Width =2265
                    Height =360
                    FontWeight =700
                    TabIndex =10
                    Name ="Command59"
                    Caption ="قيمت گذاري آخرين خريد"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Find Next"
                End
            End
        End
    End
End
CodeBehindForm
' See "PRGHEAD.cls"
