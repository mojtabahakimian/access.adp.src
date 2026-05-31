Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    OrderByOn = NotDefault
    AllowEdits = NotDefault
    DefaultView =0
    AllowUpdating =4
    ViewsAllowed =1
    TabularCharSet =178
    TabularFamily =42
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =15840
    DatasheetFontHeight =10
    ItemSuffix =168
    Left =1260
    Right =17085
    Bottom =8670
    HelpContextId =6001
    DatasheetGridlinesColor =12632256
    MaxRecords =1000000
    RecSrcDt = Begin
        0x2e57b5c0d4efe540
    End
    RecordSource ="SELECT *, 2 AS dtag, TAG AS HTAG FROM TR_HEAD_LST WHERE (TAG = 13) ORDER BY NUMB"
        "ER, UP_DATE, UP_TIME"
    Caption ="فاكتور فروش"
    HelpFile ="dena.hlp"
    DatasheetFontName ="Arial (Arabic)"
    AllowDatasheetView =0
    RecordSourceQualifier ="dbo"
    Begin
        Begin Label
            BackStyle =0
            TextFontCharSet =178
            FontName ="Tahoma"
        End
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
            Width =850
            Height =850
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
        Begin UnboundObjectFrame
            SpecialEffect =2
            OldBorderStyle =1
            Width =4536
            Height =2835
        End
        Begin PageBreak
            Width =283
        End
        Begin CustomControl
            SpecialEffect =2
            Width =4536
            Height =2835
        End
        Begin ToggleButton
            TextFontCharSet =178
            Width =283
            Height =283
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Tahoma"
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
            Height =12405
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin Tab
                    OverlapFlags =85
                    TextFontCharSet =0
                    Style =1
                    Width =15840
                    Height =12405
                    FontWeight =700
                    Name ="TabCtl56"
                    Begin
                        Begin Page
                            OverlapFlags =215
                            Left =135
                            Top =345
                            Width =15570
                            Height =11925
                            Name ="Page57"
                            Caption ="&روي فاكتور"
                            Begin
                                Begin Subform
                                    OverlapFlags =247
                                    Left =162
                                    Top =2487
                                    Width =14355
                                    Height =4935
                                    Name ="INVO_LST_sub"
                                    SourceObject ="Form.TR_INVO_LST_sub2"
                                    LinkChildFields ="NUMBER;TAG;UP_DATE;UP_TIME"
                                    LinkMasterFields ="NUMBER;DTAG;UP_DATE;UP_TIME"
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    TabStop = NotDefault
                                    DecimalPlaces =0
                                    OverlapFlags =223
                                    TextAlign =3
                                    TextFontFamily =2
                                    ReadingOrder =2
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =757
                                    Top =7425
                                    Width =1821
                                    Height =315
                                    FontWeight =700
                                    TabIndex =1
                                    BackColor =13434828
                                    Name ="JJKOL"
                                    ControlSource ="=[INVO_LST_sub].[Form]![SMABLK]"
                                    Format ="#,##0.00\" ريال \";-#,##0.00\" ريال \""
                                    FontName ="Titr"
                                    Begin
                                        Begin Label
                                            OverlapFlags =223
                                            TextFontCharSet =0
                                            Left =2677
                                            Top =7425
                                            Width =1260
                                            Height =300
                                            Name ="Label61"
                                            Caption ="جمع كل فاكتور:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    Visible = NotDefault
                                    Locked = NotDefault
                                    TabStop = NotDefault
                                    DecimalPlaces =0
                                    OverlapFlags =215
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =1
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =3997
                                    Top =7440
                                    Width =1251
                                    Height =315
                                    TabIndex =2
                                    BackColor =16777164
                                    Name ="MOGU"
                                    Begin
                                        Begin Label
                                            Visible = NotDefault
                                            OverlapFlags =215
                                            TextFontCharSet =0
                                            Left =5347
                                            Top =7440
                                            Width =750
                                            Height =300
                                            Name ="Label63"
                                            Caption ="موجودي:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    TabStop = NotDefault
                                    DecimalPlaces =0
                                    SpecialEffect =0
                                    OverlapFlags =215
                                    TextAlign =3
                                    TextFontFamily =2
                                    ReadingOrder =1
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =692
                                    Top =7755
                                    Width =1881
                                    Height =300
                                    FontWeight =700
                                    TabIndex =3
                                    BackColor =16777164
                                    Name ="Text121"
                                    ControlSource ="=[GHABEL]-[NPAR]"
                                    Format ="#,##0.00;#,##0.00-"
                                    FontName ="Titr"
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            TextFontCharSet =0
                                            Left =2673
                                            Top =7762
                                            Width =525
                                            Height =300
                                            Name ="Label122"
                                            Caption ="مانده:"
                                        End
                                    End
                                End
                                Begin CheckBox
                                    Visible = NotDefault
                                    OverlapFlags =215
                                    ReadingOrder =1
                                    Left =12812
                                    Top =8444
                                    Width =340
                                    Height =218
                                    TabIndex =4
                                    Name ="TKHF"
                                    ControlSource ="TKHF"
                                    Begin
                                        Begin Label
                                            Visible = NotDefault
                                            OverlapFlags =215
                                            TextFontCharSet =0
                                            Left =13202
                                            Top =8390
                                            Width =1590
                                            Height =240
                                            ForeColor =255
                                            Name ="Label144"
                                            Caption ="تخفيف به ريز كالا باشد"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =215
                                    TextFontCharSet =0
                                    TextAlign =1
                                    ReadingOrder =2
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    IMESentenceMode =3
                                    Left =7937
                                    Top =7940
                                    Width =1431
                                    Height =345
                                    FontSize =10
                                    FontWeight =700
                                    TabIndex =5
                                    BackColor =13408767
                                    Name ="UP_DATE"
                                    ControlSource ="UP_DATE"
                                    InputMask ="####/##/##"
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            Left =9469
                                            Top =7937
                                            Width =900
                                            Height =240
                                            Name ="Label148"
                                            Caption ="تاريخ اصلاح:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =215
                                    TextFontCharSet =0
                                    TextAlign =1
                                    ReadingOrder =2
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    IMESentenceMode =3
                                    Left =7954
                                    Top =8387
                                    Width =2841
                                    Height =345
                                    FontSize =10
                                    FontWeight =700
                                    TabIndex =6
                                    BackColor =13408767
                                    Name ="UP_TIME"
                                    ControlSource ="UP_TIME"
                                    Format ="General Date"
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            Left =10911
                                            Top =8389
                                            Width =1050
                                            Height =240
                                            Name ="Label149"
                                            Caption ="ساعت اصلاح:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =215
                                    TextAlign =3
                                    ReadingOrder =2
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    IMESentenceMode =3
                                    Left =7937
                                    Top =8827
                                    Height =345
                                    TabIndex =7
                                    BackColor =13408767
                                    Name ="UP_USER_NAME"
                                    ControlSource ="UP_USER_NAME"
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            Left =9689
                                            Top =8827
                                            Width =1290
                                            Height =240
                                            Name ="Label150"
                                            Caption ="كاربر اصلاح كننده:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =215
                                    TextFontCharSet =0
                                    TextAlign =1
                                    TextFontFamily =18
                                    ReadingOrder =2
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    IMESentenceMode =3
                                    Left =10531
                                    Top =7987
                                    Width =1431
                                    Height =345
                                    FontSize =10
                                    FontWeight =700
                                    TabIndex =8
                                    BackColor =13408767
                                    Name ="PC_NAME"
                                    ControlSource ="PC_NAME"
                                    FontName ="Times New Roman"
                                    InputMask ="##/##/##"
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            Left =12062
                                            Top =7982
                                            Width =1755
                                            Height =240
                                            Name ="Label152"
                                            Caption ="نام كامپيوتر اصلاح كننده:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =215
                                    TextFontCharSet =0
                                    TextAlign =1
                                    TextFontFamily =18
                                    ReadingOrder =2
                                    ScrollBarAlign =2
                                    IMESentenceMode =3
                                    Left =11447
                                    Top =8777
                                    Width =1929
                                    Height =345
                                    FontSize =12
                                    FontWeight =700
                                    TabIndex =9
                                    BackColor =13408767
                                    Name ="IPADD"
                                    ControlSource ="IPADD"
                                    FontName ="Times New Roman"
                                    InputMask ="##/##/##"
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            Left =11098
                                            Top =8829
                                            Width =285
                                            Height =240
                                            Name ="Label155"
                                            Caption ="IP:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    OverlapFlags =215
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =2
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =13224
                                    Top =411
                                    Width =1026
                                    Height =317
                                    TabIndex =10
                                    BackColor =16777164
                                    ForeColor =8388608
                                    Name ="NUMBER1"
                                    ControlSource ="NUMBER1"
                                    DefaultValue ="0"
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            TextFontCharSet =0
                                            Left =14334
                                            Top =411
                                            Width =675
                                            Height =317
                                            ForeColor =8388608
                                            Name ="Label1"
                                            Caption ="شماره :"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =215
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =2
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =11349
                                    Top =411
                                    Width =1161
                                    Height =317
                                    TabIndex =11
                                    ForeColor =8388608
                                    Name ="DATE_N"
                                    ControlSource ="DATE_N"
                                    DefaultValue ="=farsidate(Date())"
                                    InputMask ="####/##/##"
                                    Tag ="840617"
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            TextFontCharSet =0
                                            Left =12549
                                            Top =411
                                            Width =585
                                            Height =300
                                            ForeColor =8388608
                                            Name ="Label9"
                                            Caption ="تاريخ :"
                                        End
                                    End
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    DecimalPlaces =0
                                    OverlapFlags =215
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =2
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =9414
                                    Top =411
                                    Width =1311
                                    Height =317
                                    TabIndex =12
                                    ForeColor =8388608
                                    Name ="USER_NAME"
                                    ControlSource ="USER_NAME"
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            TextFontCharSet =0
                                            Left =10782
                                            Top =411
                                            Width =525
                                            Height =300
                                            ForeColor =8388608
                                            Name ="Label111"
                                            Caption ="كاربر:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    DecimalPlaces =0
                                    OverlapFlags =223
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =2
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =8514
                                    Top =411
                                    Width =336
                                    Height =317
                                    TabIndex =13
                                    ForeColor =8388608
                                    Name ="MAS"
                                    ControlSource ="MAS"
                                    DefaultValue ="0"
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            TextFontCharSet =0
                                            Left =8874
                                            Top =411
                                            Width =450
                                            Height =300
                                            ForeColor =8388608
                                            Name ="Label106"
                                            Caption ="مدت:"
                                        End
                                    End
                                End
                                Begin ComboBox
                                    LimitToList = NotDefault
                                    OldBorderStyle =0
                                    OverlapFlags =223
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =2
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =5979
                                    Top =416
                                    Width =1146
                                    Height =317
                                    TabIndex =14
                                    ForeColor =8388608
                                    Name ="NUMBER"
                                    ControlSource ="NUMBER"
                                    RowSourceType ="Table/View/StoredProc"
                                    RowSource ="SELECT NUMBER, TAG FROM HEAD_LST WHERE (TAG = 2) AND (NOT (NUMBER IN (SELECT HEA"
                                        "D_LST.NUMBER FROM HEAD_LST WHERE (((HEAD_LST.TAG) = 13))))) AND (SADER = 0 OR SA"
                                        "DER IS NULL) ORDER BY NUMBER"
                                    Tag ="1"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            TextFontCharSet =0
                                            Left =7164
                                            Top =416
                                            Width =1425
                                            Height =317
                                            ForeColor =8388608
                                            Name ="Label125"
                                            Caption ="شماره حواله انبار :"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =215
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =2
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    IMESentenceMode =3
                                    Left =3909
                                    Top =416
                                    Width =816
                                    Height =317
                                    TabIndex =15
                                    ForeColor =8388608
                                    Name ="FNUMCO"
                                    ControlSource ="FNUMCO"
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            TextFontCharSet =0
                                            Left =4764
                                            Top =416
                                            Width =1140
                                            Height =317
                                            ForeColor =8388608
                                            Name ="Label127"
                                            Caption ="شماره داخلي :"
                                        End
                                    End
                                End
                                Begin ComboBox
                                    LimitToList = NotDefault
                                    OldBorderStyle =0
                                    OverlapFlags =223
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =2
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    ColumnCount =2
                                    Left =12534
                                    Top =855
                                    Width =1716
                                    Height =315
                                    TabIndex =16
                                    ForeColor =8388608
                                    Name ="DEPATMAN"
                                    ControlSource ="DEPATMAN"
                                    RowSourceType ="Table/View/StoredProc"
                                    RowSource ="SELECT DEPART.DEPATMAN, DEPART.DEPNAME FROM DEPART ORDER BY DEPART.DEPNAME;"
                                    ColumnWidths ="0"
                                    DefaultValue ="=[Forms]![DEFAULT]![TFSAZMAN]"
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            TextFontCharSet =0
                                            Left =14334
                                            Top =855
                                            Width =825
                                            Height =300
                                            ForeColor =8388608
                                            Name ="Label107"
                                            Caption ="واحد :"
                                        End
                                    End
                                End
                                Begin ComboBox
                                    LimitToList = NotDefault
                                    OldBorderStyle =0
                                    OverlapFlags =215
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =2
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    ColumnCount =2
                                    Left =10674
                                    Top =855
                                    Width =1191
                                    Height =315
                                    TabIndex =17
                                    ForeColor =8388608
                                    Name ="SHIFT"
                                    ControlSource ="SHIFT"
                                    RowSourceType ="Table/View/StoredProc"
                                    RowSource ="SELECT SHIFT.SHIFT_ID, SHIFT.SHNAME FROM SHIFT ORDER BY SHIFT.SHNAME;"
                                    ColumnWidths ="0"
                                    DefaultValue ="=[Forms]![DEFAULT]![SHIFT]"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            Left =11904
                                            Top =855
                                            Width =780
                                            Height =300
                                            ForeColor =8388608
                                            Name ="Label109"
                                            Caption ="شيفت :"
                                        End
                                    End
                                End
                                Begin ComboBox
                                    LimitToList = NotDefault
                                    OldBorderStyle =0
                                    OverlapFlags =215
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =2
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    ColumnCount =2
                                    Left =7134
                                    Top =855
                                    Width =2436
                                    Height =315
                                    TabIndex =18
                                    ForeColor =8388608
                                    Name ="CUST_KIND"
                                    ControlSource ="CUST_KIND"
                                    RowSourceType ="Table/View/StoredProc"
                                    RowSource ="SELECT CUSTKIND.CUST_COD, CUSTKIND.CUSTKNAME FROM CUSTKIND ORDER BY CUSTKIND.CUS"
                                        "TKNAME;"
                                    ColumnWidths ="0"
                                    ValidationRule ="Not Is Null"
                                    DefaultValue ="[Forms]![BASEKNOW]![DEFTKH]"
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            TextFontCharSet =0
                                            Left =9609
                                            Top =855
                                            Width =1020
                                            Height =300
                                            ForeColor =8388608
                                            Name ="Label113"
                                            Caption ="نوع مشتري :"
                                        End
                                    End
                                End
                                Begin ComboBox
                                    LimitToList = NotDefault
                                    Locked = NotDefault
                                    OldBorderStyle =0
                                    OverlapFlags =215
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =2
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    ColumnCount =3
                                    Left =10029
                                    Top =1306
                                    Width =4221
                                    Height =315
                                    TabIndex =19
                                    BackColor =13434828
                                    ForeColor =8388608
                                    Name ="CUST_NO"
                                    ControlSource ="CUST_NO"
                                    RowSourceType ="Table/View/StoredProc"
                                    RowSource ="SELECT hes, NAME AS nam, hes AS Expr1 FROM CUST_HESAB"
                                    ColumnWidths ="0;2835;567"
                                    Tag ="112-1-1"
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            TextFontCharSet =0
                                            Left =14334
                                            Top =1306
                                            Width =735
                                            Height =300
                                            ForeColor =8388608
                                            Name ="Label19"
                                            Caption =" مشتري:"
                                        End
                                    End
                                End
                                Begin ComboBox
                                    LimitToList = NotDefault
                                    Locked = NotDefault
                                    OldBorderStyle =0
                                    OverlapFlags =223
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =2
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =7134
                                    Top =1306
                                    Width =2796
                                    Height =315
                                    TabIndex =20
                                    BackColor =13434828
                                    ForeColor =8388608
                                    Name ="CUST_NO2"
                                    ControlSource ="CUST_NO"
                                    RowSourceType ="Table/View/StoredProc"
                                    RowSource ="SELECT hes FROM CUST_HESAB"
                                End
                                Begin TextBox
                                    OverlapFlags =215
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =2
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =7134
                                    Top =1655
                                    Width =7125
                                    Height =315
                                    TabIndex =21
                                    ForeColor =8388608
                                    Name ="MOLAH"
                                    ControlSource ="MOLAH"
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            TextFontCharSet =0
                                            Left =14334
                                            Top =1655
                                            Width =840
                                            Height =300
                                            ForeColor =8388608
                                            Name ="Label21"
                                            Caption ="ملاحظات:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    TabStop = NotDefault
                                    DecimalPlaces =0
                                    SpecialEffect =0
                                    OverlapFlags =215
                                    TextFontCharSet =0
                                    TextAlign =2
                                    ReadingOrder =2
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =1575
                                    Top =1605
                                    Width =1416
                                    Height =300
                                    TabIndex =22
                                    BackColor =32768
                                    ForeColor =65535
                                    Name ="N_S"
                                    ControlSource ="N_S"
                                    Begin
                                        Begin Label
                                            BackStyle =1
                                            OverlapFlags =215
                                            TextFontCharSet =0
                                            Left =3000
                                            Top =1605
                                            Width =1485
                                            Height =300
                                            BackColor =32768
                                            ForeColor =65535
                                            Name ="lsanad"
                                            Caption ="ثبت درسند شماره:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    TabStop = NotDefault
                                    DecimalPlaces =0
                                    SpecialEffect =0
                                    OverlapFlags =215
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =2
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =210
                                    Top =1290
                                    Width =2781
                                    Height =300
                                    TabIndex =23
                                    BackColor =32768
                                    ForeColor =65535
                                    Name ="MANDAH"
                                    Begin
                                        Begin Label
                                            BackStyle =1
                                            OverlapFlags =215
                                            TextFontCharSet =0
                                            Left =3000
                                            Top =1290
                                            Width =1485
                                            Height =300
                                            BackColor =32768
                                            ForeColor =65535
                                            Name ="Label69"
                                            Caption ="مانده حساب:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    Visible = NotDefault
                                    TabStop = NotDefault
                                    DecimalPlaces =0
                                    OverlapFlags =247
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =2
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =6928
                                    Top =345
                                    Width =411
                                    Height =345
                                    TabIndex =24
                                    Name ="LETSANAD"
                                    DefaultValue ="False"
                                End
                                Begin TextBox
                                    Visible = NotDefault
                                    OverlapFlags =215
                                    IMESentenceMode =3
                                    Left =3438
                                    Top =880
                                    Width =276
                                    TabIndex =25
                                    Name ="DTAG"
                                    ControlSource ="DTAG"
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    TabStop = NotDefault
                                    DecimalPlaces =0
                                    SpecialEffect =0
                                    OverlapFlags =215
                                    TextFontCharSet =0
                                    TextAlign =2
                                    ReadingOrder =2
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =210
                                    Top =1605
                                    Width =1356
                                    Height =300
                                    TabIndex =26
                                    BackColor =32768
                                    ForeColor =65535
                                    Name ="MABNA"
                                End
                                Begin CheckBox
                                    OverlapFlags =215
                                    ReadingOrder =1
                                    Left =4935
                                    Top =1599
                                    Width =340
                                    Height =218
                                    TabIndex =27
                                    Name ="TICMBAA"
                                    ControlSource ="TICMBAA"
                                    Begin
                                        Begin Label
                                            OverlapFlags =223
                                            Left =5328
                                            Top =1545
                                            Width =1515
                                            Height =240
                                            ForeColor =255
                                            Name ="Label142"
                                            Caption ="محاسبه.ماليات.ب.ا.ا"
                                        End
                                    End
                                End
                                Begin CheckBox
                                    Visible = NotDefault
                                    OverlapFlags =215
                                    ReadingOrder =1
                                    Left =6734
                                    Top =1357
                                    Width =340
                                    Height =218
                                    TabIndex =28
                                    Name ="Check165"
                                    ControlSource ="TKHF"
                                    Begin
                                        Begin Label
                                            Visible = NotDefault
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            Left =7124
                                            Top =1303
                                            Width =1590
                                            Height =240
                                            ForeColor =255
                                            Name ="Label166"
                                            Caption ="تخفيف به ريز كالا باشد"
                                        End
                                    End
                                End
                                Begin CheckBox
                                    Locked = NotDefault
                                    OverlapFlags =215
                                    ReadingOrder =1
                                    Left =4950
                                    Top =1344
                                    Width =340
                                    Height =218
                                    TabIndex =29
                                    Name ="OKF"
                                    ControlSource ="OKF"
                                    DefaultValue ="False"
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            TextFontCharSet =0
                                            Left =5340
                                            Top =1290
                                            Width =975
                                            Height =240
                                            ForeColor =255
                                            Name ="Label146"
                                            Caption ="تائيد فاكتور"
                                        End
                                    End
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    OverlapFlags =215
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =2
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =1625
                                    Top =402
                                    Width =1206
                                    Height =317
                                    TabIndex =30
                                    BackColor =16777164
                                    ForeColor =8388608
                                    Name ="ANBARF"
                                    ControlSource ="ANBARF"
                                    DefaultValue ="0"
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            TextFontCharSet =0
                                            Left =2942
                                            Top =396
                                            Width =915
                                            Height =300
                                            ForeColor =8388608
                                            Name ="Label147"
                                            Caption ="شماره برگه:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    Visible = NotDefault
                                    OverlapFlags =215
                                    IMESentenceMode =3
                                    Left =1227
                                    Top =453
                                    Width =291
                                    TabIndex =31
                                    Name ="SADER"
                                    ControlSource ="SADER"
                                    DefaultValue ="0"
                                End
                                Begin TextBox
                                    Visible = NotDefault
                                    OverlapFlags =223
                                    IMESentenceMode =3
                                    Left =6564
                                    Top =793
                                    Width =411
                                    TabIndex =32
                                    Name ="CDDATE"
                                    ControlSource ="CDDATE"
                                End
                                Begin TextBox
                                    Visible = NotDefault
                                    OverlapFlags =255
                                    IMESentenceMode =3
                                    Left =6621
                                    Top =907
                                    Width =411
                                    TabIndex =33
                                    Name ="CDTIME"
                                    ControlSource ="CDTIME"
                                End
                                Begin TextBox
                                    Visible = NotDefault
                                    OverlapFlags =255
                                    IMESentenceMode =3
                                    Left =6564
                                    Top =850
                                    Width =411
                                    Height =255
                                    TabIndex =34
                                    Name ="OKDATE"
                                    ControlSource ="OKDATE"
                                End
                                Begin TextBox
                                    Visible = NotDefault
                                    OverlapFlags =247
                                    IMESentenceMode =3
                                    Left =6621
                                    Top =850
                                    Width =411
                                    TabIndex =35
                                    Name ="OKTIME"
                                    ControlSource ="OKTIME"
                                End
                                Begin CheckBox
                                    OverlapFlags =215
                                    ReadingOrder =1
                                    Left =4935
                                    Top =1809
                                    Width =340
                                    Height =218
                                    TabIndex =36
                                    Name ="JAY"
                                    ControlSource ="JAY"
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            Left =5328
                                            Top =1755
                                            Width =1515
                                            Height =240
                                            ForeColor =255
                                            Name ="JAYl"
                                            Caption ="محاسبه جايزه"
                                        End
                                    End
                                End
                                Begin ComboBox
                                    LimitToList = NotDefault
                                    Visible = NotDefault
                                    TabStop = NotDefault
                                    OldBorderStyle =0
                                    OverlapFlags =215
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =1
                                    ScrollBarAlign =2
                                    DecimalPlaces =0
                                    NumeralShapes =2
                                    IMESentenceMode =3
                                    ColumnCount =3
                                    Left =12075
                                    Top =2040
                                    Width =2196
                                    Height =330
                                    TabIndex =37
                                    BackColor =13434828
                                    ForeColor =8388608
                                    Name ="MODAT_PPID"
                                    ControlSource ="MODAT_PPID"
                                    RowSourceType ="Table/View/StoredProc"
                                    RowSource ="SELECT PPID, PPAME, MODAT FROM PRICE_PAYNO UNION SELECT 0, 'آزاد', 0"
                                    ColumnWidths ="0;1532;290"
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            TextFontCharSet =0
                                            Left =14340
                                            Top =2040
                                            Width =960
                                            Height =285
                                            ForeColor =8388608
                                            Name ="Label167"
                                            Caption ="نحوه پرداخت:"
                                        End
                                    End
                                End
                                Begin ComboBox
                                    LimitToList = NotDefault
                                    Visible = NotDefault
                                    Locked = NotDefault
                                    OldBorderStyle =0
                                    OverlapFlags =215
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =1
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    ColumnCount =2
                                    Left =6045
                                    Top =2040
                                    Width =4866
                                    Height =330
                                    TabIndex =38
                                    ForeColor =8388608
                                    Name ="PEPID"
                                    ControlSource ="PEPID"
                                    RowSourceType ="Table/View/StoredProc"
                                    RowSource ="SELECT PEPID, PEPNAME, PEPDATE, PEPDEPART FROM PRICE_ELAMIE ORDER BY PEPNAME DES"
                                        "C"
                                    ColumnWidths ="0"
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            TextFontCharSet =0
                                            Left =10920
                                            Top =2040
                                            Width =1125
                                            Height =240
                                            ForeColor =8388608
                                            Name ="Label126"
                                            Caption ="اعلاميه قيمت :"
                                        End
                                    End
                                End
                                Begin ComboBox
                                    LimitToList = NotDefault
                                    Visible = NotDefault
                                    Locked = NotDefault
                                    OldBorderStyle =0
                                    OverlapFlags =215
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =1
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    ColumnCount =2
                                    Left =225
                                    Top =2060
                                    Width =4521
                                    Height =330
                                    TabIndex =39
                                    ForeColor =8388608
                                    Name ="PEID"
                                    ControlSource ="PEID"
                                    RowSourceType ="Table/View/StoredProc"
                                    RowSource ="SELECT PEID, PENAME, PEDATE, PEPDEPART FROM PRICE_ELAMIETF"
                                    ColumnWidths ="0"
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            TextFontCharSet =0
                                            Left =4838
                                            Top =2065
                                            Width =1170
                                            Height =285
                                            ForeColor =8388608
                                            Name ="Label128"
                                            Caption ="اعلاميه تخفيف :"
                                        End
                                    End
                                End
                                Begin Subform
                                    OverlapFlags =247
                                    Left =170
                                    Top =8510
                                    Width =5102
                                    Height =1251
                                    TabIndex =40
                                    Name ="TR_TAKHFIF_APLAY_SUB"
                                    SourceObject ="Form.TR_TAKHFIF_APLAY_SUB"
                                    LinkChildFields ="number;kind"
                                    LinkMasterFields ="number;Dtag"
                                    Begin
                                        Begin Label
                                            OverlapFlags =223
                                            TextAlign =3
                                            Left =3997
                                            Top =8277
                                            Width =1275
                                            Height =240
                                            Name ="Label330"
                                            Caption ="تخفيفات پيشرفته"
                                        End
                                    End
                                End
                                Begin CheckBox
                                    Locked = NotDefault
                                    OverlapFlags =215
                                    ReadingOrder =1
                                    Left =14203
                                    Top =9411
                                    TabIndex =41
                                    Name ="SGN1"
                                    ControlSource ="SGN1"
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            Left =14464
                                            Top =9354
                                            Width =675
                                            Height =240
                                            ForeColor =16711680
                                            Name ="LSGN1"
                                            Caption ="فروش"
                                        End
                                    End
                                End
                                Begin CheckBox
                                    Locked = NotDefault
                                    OverlapFlags =215
                                    ReadingOrder =1
                                    Left =14173
                                    Top =9695
                                    Width =290
                                    TabIndex =42
                                    Name ="SGN2"
                                    ControlSource ="SGN2"
                                    Begin
                                        Begin Label
                                            OverlapFlags =223
                                            Left =14457
                                            Top =9639
                                            Width =975
                                            Height =240
                                            ForeColor =16711680
                                            Name ="LSGN2"
                                            Caption ="حسابداري"
                                        End
                                    End
                                End
                                Begin CheckBox
                                    Locked = NotDefault
                                    OverlapFlags =215
                                    ReadingOrder =1
                                    Left =14203
                                    Top =9954
                                    TabIndex =43
                                    Name ="SGN3"
                                    ControlSource ="SGN3"
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            Left =14457
                                            Top =9891
                                            Width =960
                                            Height =240
                                            ForeColor =16711680
                                            Name ="LSGN3"
                                            Caption ="مدير عامل"
                                        End
                                    End
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =247
                            AccessKey =129
                            Left =122
                            Top =450
                            Width =15583
                            Height =11820
                            Name ="Page58"
                            Caption ="&پشت فاكتور"
                            Begin
                                Begin TextBox
                                    DecimalPlaces =0
                                    OverlapFlags =255
                                    TextAlign =3
                                    TextFontFamily =2
                                    ReadingOrder =2
                                    KeyboardLanguage =43
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =11439
                                    Top =1035
                                    Height =315
                                    FontWeight =700
                                    Name ="M_NAGHD"
                                    ControlSource ="M_NAGHD"
                                    Format ="#,##0.00\" ريال \";-#,##0.00\" ريال \""
                                    FontName ="Titr"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            TextFontCharSet =0
                                            Left =13178
                                            Top =1035
                                            Width =750
                                            Height =300
                                            Name ="Label23"
                                            Caption ="مبلغ نقد:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    DecimalPlaces =0
                                    OverlapFlags =255
                                    TextAlign =3
                                    TextFontFamily =2
                                    ReadingOrder =2
                                    KeyboardLanguage =43
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =11439
                                    Top =1375
                                    Height =315
                                    FontWeight =700
                                    TabIndex =1
                                    Name ="MABL_VAR"
                                    ControlSource ="MABL_VAR"
                                    Format ="#,##0.00\" ريال \";-#,##0.00\" ريال \""
                                    FontName ="Titr"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            TextFontCharSet =0
                                            Left =13178
                                            Top =1380
                                            Width =1065
                                            Height =300
                                            Name ="Label25"
                                            Caption ="مبلغ واريزي:"
                                        End
                                    End
                                End
                                Begin ComboBox
                                    LimitToList = NotDefault
                                    OldBorderStyle =0
                                    OverlapFlags =255
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =1
                                    KeyboardLanguage =43
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    ColumnCount =2
                                    Left =5267
                                    Top =1380
                                    Width =4851
                                    Height =315
                                    TabIndex =2
                                    Name ="MOIN_VAR"
                                    ControlSource ="MOIN_VAR"
                                    RowSourceType ="Table/View/StoredProc"
                                    RowSource ="SELECT  RTRIM(CAST(N_KOL AS nvarchar)) + '-' + RTRIM(CAST(NUMBER AS nvarchar)) +"
                                        " '-' + RTRIM(CAST(TNUMBER AS nvarchar)) AS Expr1, NAME,   N_KOL FROM TDETA_HES W"
                                        "HERE     (N_KOL =112)"
                                    ColumnWidths ="853"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            TextFontCharSet =0
                                            Left =10161
                                            Top =1382
                                            Width =1110
                                            Height =300
                                            Name ="Label27"
                                            Caption ="معين واريزي:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    DecimalPlaces =0
                                    OverlapFlags =255
                                    TextAlign =3
                                    TextFontFamily =2
                                    ReadingOrder =2
                                    KeyboardLanguage =43
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =11439
                                    Top =1725
                                    Height =315
                                    FontWeight =700
                                    TabIndex =3
                                    Name ="MABL_HAV"
                                    ControlSource ="MABL_HAV"
                                    Format ="#,##0.00\" ريال \";-#,##0.00\" ريال \""
                                    FontName ="Titr"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            TextFontCharSet =0
                                            Left =13178
                                            Top =1725
                                            Width =930
                                            Height =300
                                            Name ="Label29"
                                            Caption ="مبلغ حواله:"
                                        End
                                    End
                                End
                                Begin ComboBox
                                    LimitToList = NotDefault
                                    OldBorderStyle =0
                                    OverlapFlags =255
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =1
                                    KeyboardLanguage =43
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    ColumnCount =2
                                    Left =5267
                                    Top =1727
                                    Width =4851
                                    Height =315
                                    TabIndex =4
                                    Name ="MOIN_HAV"
                                    ControlSource ="MOIN_HAV"
                                    RowSourceType ="Table/View/StoredProc"
                                    RowSource ="SELECT  RTRIM(CAST(N_KOL AS nvarchar)) + '-' + RTRIM(CAST(NUMBER AS nvarchar)) +"
                                        " '-' + RTRIM(CAST(TNUMBER AS nvarchar)) AS Expr1, NAME,   N_KOL FROM TDETA_HES W"
                                        "HERE     (N_KOL =112)"
                                    ColumnWidths ="853"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            TextFontCharSet =0
                                            Left =10161
                                            Top =1727
                                            Width =975
                                            Height =300
                                            Name ="Label31"
                                            Caption ="معين حواله:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    DecimalPlaces =0
                                    OverlapFlags =255
                                    TextAlign =3
                                    TextFontFamily =2
                                    ReadingOrder =2
                                    KeyboardLanguage =43
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =11439
                                    Top =2070
                                    Height =315
                                    FontWeight =700
                                    TabIndex =5
                                    Name ="TAKHFIF"
                                    ControlSource ="TAKHFIF"
                                    Format ="#,##0.00\" ريال \";-#,##0.00\" ريال \""
                                    FontName ="Titr"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            TextFontCharSet =0
                                            Left =13178
                                            Top =2085
                                            Width =990
                                            Height =300
                                            Name ="Label37"
                                            Caption ="مبلغ تخفيف:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    OverlapFlags =255
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =1
                                    KeyboardLanguage =43
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =5267
                                    Top =2057
                                    Width =4851
                                    Height =315
                                    TabIndex =6
                                    Name ="takh"
                                    Format ="Fixed"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            TextFontCharSet =0
                                            Left =10161
                                            Top =2072
                                            Width =1020
                                            Height =300
                                            Name ="Label102"
                                            Caption ="درصد تخفيف:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    DecimalPlaces =0
                                    OverlapFlags =255
                                    TextAlign =3
                                    TextFontFamily =2
                                    ReadingOrder =2
                                    KeyboardLanguage =43
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =11437
                                    Top =2430
                                    Height =315
                                    FontWeight =700
                                    TabIndex =7
                                    Name ="MABL_HAZ2"
                                    ControlSource ="MABL_HAZ"
                                    Format ="#,##0.00\" ريال \";-#,##0.00\" ريال \""
                                    FontName ="Titr"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            TextFontCharSet =0
                                            Left =13172
                                            Top =2430
                                            Width =1095
                                            Height =300
                                            Name ="Label115"
                                            Caption ="خدمات:"
                                        End
                                    End
                                End
                                Begin ComboBox
                                    OldBorderStyle =0
                                    OverlapFlags =255
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =1
                                    KeyboardLanguage =43
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    ColumnCount =2
                                    Left =5252
                                    Top =2416
                                    Width =4851
                                    Height =315
                                    TabIndex =8
                                    Name ="MOIN_HAZ"
                                    ControlSource ="MOIN_HAZ"
                                    RowSourceType ="Table/View/StoredProc"
                                    RowSource ="SELECT     RTRIM(CAST(N_KOL AS nvarchar)) + '-' + RTRIM(CAST(NUMBER AS nvarchar)"
                                        ") + '-' + RTRIM(CAST(TNUMBER AS nvarchar)) AS Expr1, NAME, RTRIM(CAST(N_KOL AS n"
                                        "varchar)) + '-' + RTRIM(CAST(NUMBER AS nvarchar)) + '-' + RTRIM(CAST(TNUMBER AS "
                                        "nvarchar)) AS Expr2 FROM TDETA_HES WHERE     (N_KOL = 521)"
                                    ColumnWidths ="853"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            TextFontCharSet =0
                                            Left =10151
                                            Top =2415
                                            Width =1200
                                            Height =300
                                            Name ="Label117"
                                            Caption ="معين خدمات:"
                                        End
                                    End
                                End
                                Begin Rectangle
                                    OverlapFlags =255
                                    Left =632
                                    Top =495
                                    Width =13815
                                    Height =371
                                    Name ="Box80"
                                End
                                Begin Rectangle
                                    OverlapFlags =255
                                    Left =5162
                                    Top =1005
                                    Width =9285
                                    Height =2171
                                    Name ="Box81"
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    TabStop = NotDefault
                                    SpecialEffect =0
                                    OverlapFlags =247
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =1
                                    KeyboardLanguage =43
                                    ScrollBarAlign =2
                                    BackStyle =0
                                    NumeralShapes =2
                                    Left =11956
                                    Top =546
                                    Width =1116
                                    Height =300
                                    TabIndex =9
                                    Name ="Text78"
                                    ControlSource ="NUMBER"
                                    DefaultValue ="0"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            Left =13081
                                            Top =544
                                            Width =1110
                                            Height =300
                                            Name ="Label79"
                                            Caption ="شماره فاكتور:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    TabStop = NotDefault
                                    SpecialEffect =0
                                    OverlapFlags =247
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =1
                                    KeyboardLanguage =43
                                    ScrollBarAlign =2
                                    BackStyle =0
                                    NumeralShapes =2
                                    Left =9722
                                    Top =546
                                    Width =1131
                                    Height =300
                                    TabIndex =10
                                    Name ="Text73"
                                    ControlSource ="DATE_N"
                                    InputMask ="####/##/##"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            TextFontCharSet =0
                                            Left =10862
                                            Top =540
                                            Width =1065
                                            Height =300
                                            Name ="Label74"
                                            Caption ="تاريخ فاكتور:"
                                        End
                                    End
                                End
                                Begin ComboBox
                                    LimitToList = NotDefault
                                    Locked = NotDefault
                                    TabStop = NotDefault
                                    SpecialEffect =0
                                    OldBorderStyle =0
                                    OverlapFlags =247
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =1
                                    KeyboardLanguage =43
                                    ScrollBarAlign =2
                                    BackStyle =0
                                    NumeralShapes =2
                                    ColumnCount =2
                                    Left =5237
                                    Top =546
                                    Width =3621
                                    Height =300
                                    TabIndex =11
                                    Name ="Combo75"
                                    ControlSource ="CUST_NO"
                                    RowSourceType ="Table/View/StoredProc"
                                    RowSource ="SELECT RTRIM(CAST(N_KOL AS nvarchar)) + '-' + RTRIM(CAST(NUMBER AS nvarchar)) + "
                                        "'-' + RTRIM(CAST(TNUMBER AS nvarchar)) AS hes, NAME AS nam, RTRIM(CAST(N_KOL AS "
                                        "nvarchar)) + '-' + RTRIM(CAST(NUMBER AS nvarchar)) + '-' + RTRIM(CAST(TNUMBER AS"
                                        " nvarchar)) AS Expr1 FROM TDETA_HES"
                                    ColumnWidths ="0;2835;567"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            TextFontCharSet =0
                                            Left =8867
                                            Top =540
                                            Width =795
                                            Height =300
                                            Name ="Label76"
                                            Caption =" مشتري:"
                                        End
                                    End
                                End
                                Begin Subform
                                    OverlapFlags =247
                                    Left =122
                                    Top =3495
                                    Width =14280
                                    Height =4170
                                    TabIndex =12
                                    Name ="PAY_GETD_SUB"
                                    SourceObject ="Form.TR_PAY_GETD_SUB2"
                                    LinkChildFields ="NUMBER;TAG;UP_DATE;UP_TIME"
                                    LinkMasterFields ="NUMBER;DTAG;UP_DATE;UP_TIME"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            TextFontCharSet =0
                                            Left =13052
                                            Top =3225
                                            Width =1200
                                            Height =300
                                            Name ="PAY_GETD_SUB Label"
                                            Caption ="چكهاي دريافتي"
                                            EventProcPrefix ="PAY_GETD_SUB_Label"
                                        End
                                    End
                                End
                                Begin TextBox
                                    TabStop = NotDefault
                                    DecimalPlaces =0
                                    SpecialEffect =0
                                    OverlapFlags =255
                                    TextAlign =3
                                    TextFontFamily =2
                                    ReadingOrder =1
                                    KeyboardLanguage =43
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =732
                                    Top =2310
                                    Width =2106
                                    Height =300
                                    FontWeight =700
                                    TabIndex =13
                                    Name ="NPAR"
                                    ControlSource ="=[M_NAGHD]+[MABL_VAR]+[MABL_HAV]+[NCHK]"
                                    Format ="#,##0.00;-#,##0.00"
                                    FontName ="Titr"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            TextFontCharSet =0
                                            Left =2938
                                            Top =2317
                                            Width =1545
                                            Height =300
                                            Name ="Label86"
                                            Caption ="جمع مبالغ پرداختي:"
                                        End
                                    End
                                End
                                Begin Rectangle
                                    OverlapFlags =255
                                    Left =625
                                    Top =907
                                    Width =3912
                                    Height =2154
                                    Name ="Box95"
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    TabStop = NotDefault
                                    DecimalPlaces =0
                                    SpecialEffect =0
                                    OverlapFlags =247
                                    TextAlign =3
                                    TextFontFamily =2
                                    ReadingOrder =1
                                    KeyboardLanguage =43
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =733
                                    Top =990
                                    Width =2106
                                    Height =300
                                    FontWeight =700
                                    TabIndex =14
                                    Name ="JF"
                                    ControlSource ="=[INVO_LST_sub].[Form]![SMABLK]"
                                    Format ="#,##0.00;-#,##0.00"
                                    FontName ="Titr"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            Left =2943
                                            Top =990
                                            Width =1260
                                            Height =300
                                            Name ="Label88"
                                            Caption ="جمع كل فاكتور:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    TabStop = NotDefault
                                    DecimalPlaces =0
                                    SpecialEffect =0
                                    OverlapFlags =247
                                    TextAlign =3
                                    TextFontFamily =2
                                    ReadingOrder =1
                                    KeyboardLanguage =43
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =732
                                    Top =1320
                                    Width =2106
                                    Height =300
                                    FontWeight =700
                                    TabIndex =15
                                    Name ="HKH"
                                    ControlSource ="=[MABL_HAZ]"
                                    Format ="#,##0.00;-#,##0.00"
                                    FontName ="Titr"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            Left =2937
                                            Top =1320
                                            Width =1110
                                            Height =240
                                            Name ="Label92"
                                            Caption ="خدمات:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    TabStop = NotDefault
                                    DecimalPlaces =0
                                    SpecialEffect =0
                                    OverlapFlags =255
                                    TextAlign =3
                                    TextFontFamily =2
                                    ReadingOrder =1
                                    KeyboardLanguage =43
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =732
                                    Top =1635
                                    Width =2106
                                    Height =345
                                    FontWeight =700
                                    TabIndex =16
                                    Name ="NTKHFIF"
                                    ControlSource ="=[TAKHFIF]"
                                    Format ="#,##0.00;-#,##0.00"
                                    FontName ="Titr"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            Left =2937
                                            Top =1635
                                            Width =705
                                            Height =240
                                            Name ="Label94"
                                            Caption ="تخفيفات:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    TabStop = NotDefault
                                    DecimalPlaces =0
                                    SpecialEffect =0
                                    OverlapFlags =247
                                    TextAlign =3
                                    TextFontFamily =2
                                    ReadingOrder =1
                                    KeyboardLanguage =43
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =732
                                    Top =1965
                                    Width =2106
                                    Height =300
                                    FontWeight =700
                                    TabIndex =17
                                    BackColor =16777164
                                    Name ="GHABEL"
                                    ControlSource ="=[JF]+[HKH]-[NTKHFIF]+[MBAA]"
                                    Format ="#,##0.00;-#,##0.00"
                                    FontName ="Titr"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            Left =2937
                                            Top =1972
                                            Width =1440
                                            Height =300
                                            Name ="Label97"
                                            Caption ="مبلغ قابل پرداخت:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    TabStop = NotDefault
                                    DecimalPlaces =0
                                    SpecialEffect =0
                                    OverlapFlags =247
                                    TextAlign =3
                                    TextFontFamily =2
                                    ReadingOrder =1
                                    KeyboardLanguage =43
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =732
                                    Top =2640
                                    Width =2106
                                    Height =300
                                    FontWeight =700
                                    TabIndex =18
                                    BackColor =16777164
                                    Name ="MAN"
                                    ControlSource ="=[GHABEL]-[NPAR]"
                                    Format ="#,##0.00;#,##0.00-"
                                    FontName ="Titr"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            Left =2938
                                            Top =2647
                                            Width =525
                                            Height =300
                                            Name ="Label99"
                                            Caption ="مانده:"
                                        End
                                    End
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
                                    Left =1815
                                    Top =7823
                                    Width =2106
                                    Height =315
                                    TabIndex =19
                                    Name ="NCHK"
                                    ControlSource ="=IIf(IsError([PAY_GETD_SUB].[Form]![SCMAB]),0,[PAY_GETD_SUB].[Form]![SCMAB])"
                                    Format ="#,##0.00\" ريال \";-#,##0.00\" ريال \""
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            Left =4014
                                            Top =7823
                                            Width =1260
                                            Height =300
                                            Name ="Label90"
                                            Caption ="جمع مبلغ چكها:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    DecimalPlaces =0
                                    OverlapFlags =247
                                    TextAlign =3
                                    TextFontFamily =2
                                    ReadingOrder =2
                                    KeyboardLanguage =43
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =11432
                                    Top =2775
                                    Height =315
                                    FontWeight =700
                                    TabIndex =20
                                    Name ="MBAA"
                                    ControlSource ="MBAA"
                                    Format ="#,##0.00\" ريال \";-#,##0.00\" ريال \""
                                    FontName ="Titr"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            Left =13172
                                            Top =2775
                                            Width =1200
                                            Height =285
                                            Name ="Label136"
                                            Caption ="ماليات و عوارض:"
                                        End
                                    End
                                End
                                Begin ComboBox
                                    LimitToList = NotDefault
                                    OldBorderStyle =0
                                    OverlapFlags =247
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =1
                                    KeyboardLanguage =43
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    ColumnCount =2
                                    Left =5252
                                    Top =2761
                                    Width =4851
                                    Height =315
                                    TabIndex =21
                                    Name ="HMBAA"
                                    ControlSource ="HMBAA"
                                    RowSourceType ="Table/View/StoredProc"
                                    RowSource ="SELECT RTRIM(CAST(N_KOL AS nvarchar)) + '-' + RTRIM(CAST(NUMBER AS nvarchar)) + "
                                        "'-' + RTRIM(CAST(TNUMBER AS nvarchar)) AS Expr1, NAME, RTRIM(CAST(N_KOL AS nvarc"
                                        "har)) + '-' + RTRIM(CAST(NUMBER AS nvarchar)) + '-' + RTRIM(CAST(TNUMBER AS nvar"
                                        "char)) AS Expr2 FROM TDETA_HES"
                                    ColumnWidths ="853"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            Left =10151
                                            Top =2760
                                            Width =1200
                                            Height =300
                                            Name ="Label138"
                                            Caption ="معين ماليات:"
                                        End
                                    End
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =247
                            Left =135
                            Top =450
                            Width =15570
                            Height =11820
                            Name ="Page155"
                            Caption ="ساير"
                            Begin
                                Begin Subform
                                    OverlapFlags =247
                                    Left =420
                                    Top =1185
                                    Width =14265
                                    Height =2700
                                    Name ="VISITOR_DTL_SUB"
                                    SourceObject ="Form.TR_VISITOR_DTL_SUB"
                                    LinkChildFields ="NUMBER;TAG;UP_DATE;UP_TIME"
                                    LinkMasterFields ="NUMBER;dTAG;UP_DATE;UP_TIME"
                                End
                                Begin Rectangle
                                    OverlapFlags =255
                                    Left =195
                                    Top =450
                                    Width =14580
                                    Height =371
                                    Name ="Box170"
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    TabStop = NotDefault
                                    SpecialEffect =0
                                    OverlapFlags =247
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =1
                                    KeyboardLanguage =43
                                    ScrollBarAlign =2
                                    BackStyle =0
                                    NumeralShapes =2
                                    Left =12329
                                    Top =501
                                    Width =1116
                                    Height =300
                                    TabIndex =1
                                    Name ="Text171"
                                    ControlSource ="NUMBER"
                                    DefaultValue ="0"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            Left =13454
                                            Top =499
                                            Width =1110
                                            Height =300
                                            Name ="Label172"
                                            Caption ="شماره فاكتور:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    TabStop = NotDefault
                                    SpecialEffect =0
                                    OverlapFlags =247
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =1
                                    KeyboardLanguage =43
                                    ScrollBarAlign =2
                                    BackStyle =0
                                    NumeralShapes =2
                                    Left =10095
                                    Top =501
                                    Width =1131
                                    Height =300
                                    TabIndex =2
                                    Name ="Text173"
                                    ControlSource ="DATE_N"
                                    InputMask ="####/##/##"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            Left =11235
                                            Top =495
                                            Width =1065
                                            Height =300
                                            Name ="Label174"
                                            Caption ="تاريخ فاكتور:"
                                        End
                                    End
                                End
                                Begin ComboBox
                                    LimitToList = NotDefault
                                    Locked = NotDefault
                                    TabStop = NotDefault
                                    SpecialEffect =0
                                    OldBorderStyle =0
                                    OverlapFlags =247
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =1
                                    KeyboardLanguage =43
                                    ScrollBarAlign =2
                                    BackStyle =0
                                    NumeralShapes =2
                                    ColumnCount =2
                                    Left =1845
                                    Top =501
                                    Width =7386
                                    Height =300
                                    TabIndex =3
                                    Name ="Combo175"
                                    ControlSource ="CUST_NO"
                                    RowSourceType ="Table/View/StoredProc"
                                    RowSource ="SELECT RTRIM(CAST(N_KOL AS nvarchar)) + '-' + RTRIM(CAST(NUMBER AS nvarchar)) + "
                                        "'-' + RTRIM(CAST(TNUMBER AS nvarchar)) AS hes, NAME AS nam, RTRIM(CAST(N_KOL AS "
                                        "nvarchar)) + '-' + RTRIM(CAST(NUMBER AS nvarchar)) + '-' + RTRIM(CAST(TNUMBER AS"
                                        " nvarchar)) AS Expr1 FROM TDETA_HES"
                                    ColumnWidths ="0;2835;567"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            Left =9240
                                            Top =495
                                            Width =795
                                            Height =300
                                            Name ="Label176"
                                            Caption =" مشتري:"
                                        End
                                    End
                                End
                                Begin Subform
                                    OverlapFlags =247
                                    OldBorderStyle =0
                                    SpecialEffect =0
                                    Left =4830
                                    Top =5160
                                    Width =9885
                                    Height =2700
                                    TabIndex =4
                                    Name ="OTHER_DTL_SUB"
                                    SourceObject ="Form.TR_OTHER_DTL_SUB1"
                                    LinkChildFields ="NUMBER;TAG"
                                    LinkMasterFields ="NUMBER;DTAG"
                                End
                                Begin Rectangle
                                    OverlapFlags =255
                                    Left =195
                                    Top =4815
                                    Width =14575
                                    Height =3345
                                    Name ="Box179"
                                End
                                Begin Label
                                    BackStyle =1
                                    OverlapFlags =247
                                    TextAlign =2
                                    Left =12315
                                    Top =4710
                                    Width =1590
                                    Height =240
                                    BackColor =-2147483633
                                    Name ="OTHER_DTL_SUB Label"
                                    Caption ="ساير اطلاعات"
                                    EventProcPrefix ="OTHER_DTL_SUB_Label"
                                End
                                Begin Rectangle
                                    OverlapFlags =255
                                    Left =195
                                    Top =1023
                                    Width =14568
                                    Height =3345
                                    Name ="Box182"
                                End
                                Begin Label
                                    BackStyle =1
                                    OverlapFlags =247
                                    TextAlign =2
                                    Left =12600
                                    Top =915
                                    Width =1245
                                    Height =240
                                    BackColor =-2147483633
                                    Name ="VISITOR_DTL_SUB Label"
                                    Caption ="پورسانت"
                                    EventProcPrefix ="VISITOR_DTL_SUB_Label"
                                End
                                Begin Rectangle
                                    OverlapFlags =255
                                    Left =450
                                    Top =5205
                                    Width =3912
                                    Height =2154
                                    Name ="Box177"
                                End
                                Begin TextBox
                                    TabStop = NotDefault
                                    DecimalPlaces =0
                                    SpecialEffect =0
                                    OverlapFlags =247
                                    TextAlign =3
                                    TextFontFamily =2
                                    ReadingOrder =1
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =557
                                    Top =6608
                                    Width =2106
                                    Height =300
                                    FontWeight =700
                                    TabIndex =5
                                    Name ="Text178"
                                    ControlSource ="=[M_NAGHD]+[MABL_VAR]+[MABL_HAV]+[NCHK]"
                                    Format ="#,##0.00;-#,##0.00"
                                    FontName ="Titr"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            Left =2763
                                            Top =6615
                                            Width =1545
                                            Height =300
                                            Name ="Label179"
                                            Caption ="جمع مبالغ پرداختي:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    TabStop = NotDefault
                                    DecimalPlaces =0
                                    SpecialEffect =0
                                    OverlapFlags =247
                                    TextAlign =3
                                    TextFontFamily =2
                                    ReadingOrder =1
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =558
                                    Top =5288
                                    Width =2106
                                    Height =300
                                    FontWeight =700
                                    TabIndex =6
                                    Name ="Text180"
                                    ControlSource ="=IIf(IsError([INVO_LST_sub].[Form]![SMABLK]) Or IsNull([INVO_LST_sub].[Form]![SM"
                                        "ABLK]),0,[INVO_LST_sub].[Form]![SMABLK])"
                                    Format ="#,##0.00;-#,##0.00"
                                    FontName ="Titr"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            Left =2763
                                            Top =5288
                                            Width =1260
                                            Height =300
                                            Name ="Label181"
                                            Caption ="جمع كل فاكتور:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    TabStop = NotDefault
                                    DecimalPlaces =0
                                    SpecialEffect =0
                                    OverlapFlags =247
                                    TextAlign =3
                                    TextFontFamily =2
                                    ReadingOrder =1
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =557
                                    Top =5618
                                    Width =2106
                                    Height =300
                                    FontWeight =700
                                    TabIndex =7
                                    Name ="Text182"
                                    ControlSource ="=[MABL_HAZ]"
                                    Format ="#,##0.00;-#,##0.00"
                                    FontName ="Titr"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            Left =2762
                                            Top =5618
                                            Width =1080
                                            Height =300
                                            Name ="Label183"
                                            Caption ="هزينه خدمات:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    TabStop = NotDefault
                                    DecimalPlaces =0
                                    SpecialEffect =0
                                    OverlapFlags =247
                                    TextAlign =3
                                    TextFontFamily =2
                                    ReadingOrder =1
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =557
                                    Top =5948
                                    Width =2106
                                    Height =300
                                    FontWeight =700
                                    TabIndex =8
                                    Name ="Text184"
                                    ControlSource ="=[TAKHFIF]"
                                    Format ="#,##0.00;-#,##0.00"
                                    FontName ="Titr"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            Left =2762
                                            Top =5948
                                            Width =735
                                            Height =300
                                            Name ="Label185"
                                            Caption ="تخفيفات:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    TabStop = NotDefault
                                    DecimalPlaces =0
                                    SpecialEffect =0
                                    OverlapFlags =247
                                    TextAlign =3
                                    TextFontFamily =2
                                    ReadingOrder =1
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =557
                                    Top =6278
                                    Width =2106
                                    Height =300
                                    FontWeight =700
                                    TabIndex =9
                                    BackColor =16777164
                                    Name ="Text186"
                                    ControlSource ="=[JF]+[HKH]-[NTKHFIF]+[MBAA]"
                                    Format ="#,##0.00;-#,##0.00"
                                    FontName ="Titr"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            Left =2762
                                            Top =6285
                                            Width =1440
                                            Height =300
                                            Name ="Label187"
                                            Caption ="مبلغ قابل پرداخت:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    TabStop = NotDefault
                                    DecimalPlaces =0
                                    SpecialEffect =0
                                    OverlapFlags =247
                                    TextAlign =3
                                    TextFontFamily =2
                                    ReadingOrder =1
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =557
                                    Top =6938
                                    Width =2106
                                    Height =300
                                    FontWeight =700
                                    TabIndex =10
                                    BackColor =16777164
                                    Name ="Text188"
                                    ControlSource ="=[GHABEL]-[NPAR]"
                                    Format ="#,##0.00;#,##0.00-"
                                    FontName ="Titr"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            Left =2763
                                            Top =6945
                                            Width =525
                                            Height =300
                                            Name ="Label189"
                                            Caption ="مانده:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    DecimalPlaces =0
                                    OverlapFlags =247
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =2
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =570
                                    Top =3960
                                    Width =2106
                                    Height =315
                                    TabIndex =11
                                    BackColor =13408767
                                    Name ="Text190"
                                    ControlSource ="=IIf(IsNull([VISITOR_DTL_SUB].[Form]![SCMAB]) Or IsError([VISITOR_DTL_SUB].[Form"
                                        "]![SCMAB]),0,[VISITOR_DTL_SUB].[Form]![SCMAB])"
                                    Format ="#,##0.00\" ريال \";-#,##0.00\" ريال \""
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            Left =2776
                                            Top =3960
                                            Width =1260
                                            Height =300
                                            Name ="Label191"
                                            Caption ="جمع پورسانت:"
                                        End
                                    End
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
' See "TR_HEAD_LST_FROOSH2.cls"
