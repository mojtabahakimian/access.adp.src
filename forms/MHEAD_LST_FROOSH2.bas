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
    ScrollBars =0
    ViewsAllowed =1
    TabularCharSet =178
    TabularFamily =42
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =14116
    DatasheetFontHeight =10
    ItemSuffix =179
    Left =495
    Top =780
    Right =14610
    Bottom =9045
    HelpContextId =6001
    DatasheetGridlinesColor =12632256
    MaxRecords =100000
    RecSrcDt = Begin
        0x224c3d3b9ee7e340
    End
    RecordSource ="HAVALAH_FROOSH"
    Caption ="فاكتور فروش"
    OnCurrent ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    HelpFile ="dena.hlp"
    DatasheetFontName ="Arial (Arabic)"
    OnError ="[Event Procedure]"
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
            Height =8445
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin Tab
                    OverlapFlags =85
                    TextFontCharSet =0
                    Style =1
                    Width =14085
                    Height =8445
                    FontWeight =700
                    Name ="TabCtl56"
                    Begin
                        Begin Page
                            OverlapFlags =215
                            Left =60
                            Top =345
                            Width =13894
                            Height =7965
                            Name ="Page57"
                            Caption ="&روي فاكتور"
                            Begin
                                Begin TextBox
                                    Locked = NotDefault
                                    OverlapFlags =215
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =1
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =12004
                                    Top =411
                                    Width =1026
                                    Height =317
                                    BackColor =16777164
                                    ForeColor =8388608
                                    Name ="NUMBER1"
                                    ControlSource ="NUMBER1"
                                    DefaultValue ="0"
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            TextFontCharSet =0
                                            Left =13114
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
                                    ReadingOrder =1
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =10129
                                    Top =411
                                    Width =1161
                                    Height =317
                                    TabIndex =1
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
                                            Left =11329
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
                                    ReadingOrder =1
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =8194
                                    Top =411
                                    Width =1311
                                    Height =317
                                    TabIndex =2
                                    ForeColor =8388608
                                    Name ="USER_NAME"
                                    ControlSource ="USER_NAME"
                                    DefaultValue ="UCurrentUser()"
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            TextFontCharSet =0
                                            Left =9562
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
                                    ReadingOrder =1
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =7294
                                    Top =411
                                    Width =336
                                    Height =317
                                    TabIndex =3
                                    ForeColor =8388608
                                    Name ="MAS"
                                    ControlSource ="MAS"
                                    DefaultValue ="0"
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            TextFontCharSet =0
                                            Left =7654
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
                                    ReadingOrder =1
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =4770
                                    Top =416
                                    Width =1146
                                    Height =317
                                    TabIndex =4
                                    ForeColor =8388608
                                    Name ="NUMBER"
                                    ControlSource ="NUMBER"
                                    RowSourceType ="Table/View/StoredProc"
                                    RowSource ="SELECT NUMBER, TAG FROM HEAD_LST WHERE (TAG = 2) AND (NOT (NUMBER IN (SELECT HEA"
                                        "D_LST.NUMBER FROM HEAD_LST WHERE (((HEAD_LST.TAG) = 13))))) ORDER BY NUMBER"
                                    Tag ="1"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            TextFontCharSet =0
                                            Left =5955
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
                                    OverlapFlags =223
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =1
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    IMESentenceMode =3
                                    Left =2700
                                    Top =416
                                    Width =816
                                    Height =317
                                    TabIndex =5
                                    ForeColor =8388608
                                    Name ="FNUMCO"
                                    ControlSource ="FNUMCO"
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            TextFontCharSet =0
                                            Left =3555
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
                                    Locked = NotDefault
                                    OldBorderStyle =0
                                    OverlapFlags =223
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =1
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    ColumnCount =2
                                    Left =11314
                                    Top =855
                                    Width =1716
                                    Height =315
                                    TabIndex =6
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
                                            Left =13114
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
                                    OldBorderStyle =0
                                    OverlapFlags =215
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =1
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    ColumnCount =2
                                    Left =9454
                                    Top =855
                                    Width =1191
                                    Height =315
                                    TabIndex =7
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
                                            Left =10684
                                            Top =855
                                            Width =660
                                            Height =300
                                            ForeColor =8388608
                                            Name ="Label109"
                                            Caption ="شيفت :"
                                        End
                                    End
                                End
                                Begin ComboBox
                                    OldBorderStyle =0
                                    OverlapFlags =215
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =1
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    ColumnCount =2
                                    Left =6784
                                    Top =855
                                    Width =1566
                                    Height =315
                                    TabIndex =8
                                    ForeColor =8388608
                                    Name ="CUST_KIND"
                                    ControlSource ="CUST_KIND"
                                    RowSourceType ="Table/View/StoredProc"
                                    RowSource ="SELECT CUSTKIND.CUST_COD, CUSTKIND.CUSTKNAME FROM CUSTKIND ORDER BY CUSTKIND.CUS"
                                        "TKNAME;"
                                    ColumnWidths ="0"
                                    ValidationRule ="Not Is Null"
                                    DefaultValue ="1"
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            TextFontCharSet =0
                                            Left =8389
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
                                    ReadingOrder =1
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    ColumnCount =3
                                    Left =8805
                                    Top =1260
                                    Width =4221
                                    Height =315
                                    TabIndex =9
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
                                            Left =13110
                                            Top =1260
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
                                    OverlapFlags =215
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =1
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =6780
                                    Top =1260
                                    Width =1926
                                    Height =315
                                    TabIndex =10
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
                                    ReadingOrder =1
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =6784
                                    Top =1655
                                    Width =6255
                                    Height =315
                                    TabIndex =11
                                    ForeColor =8388608
                                    Name ="MOLAH"
                                    ControlSource ="MOLAH"
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            TextFontCharSet =0
                                            Left =13114
                                            Top =1655
                                            Width =840
                                            Height =300
                                            ForeColor =8388608
                                            Name ="Label21"
                                            Caption ="ملاحظات:"
                                        End
                                    End
                                End
                                Begin Subform
                                    OverlapFlags =247
                                    Left =60
                                    Top =2097
                                    Width =13845
                                    Height =5370
                                    TabIndex =12
                                    Name ="INVO_LST_sub"
                                    SourceObject ="Form.MINVO_LST_SUB2"
                                    LinkChildFields ="NUMBER;TAG"
                                    LinkMasterFields ="NUMBER;DTAG"
                                End
                                Begin ComboBox
                                    Visible = NotDefault
                                    OldBorderStyle =0
                                    OverlapFlags =215
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =1
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    ColumnCount =2
                                    Left =4010
                                    Top =7930
                                    Width =3891
                                    Height =315
                                    TabIndex =13
                                    Name ="MOIN_HAZ"
                                    ControlSource ="MOIN_HAZ"
                                    RowSourceType ="Table/View/StoredProc"
                                    RowSource ="SELECT     RTRIM(CAST(N_KOL AS nvarchar)) + '-' + RTRIM(CAST(NUMBER AS nvarchar)"
                                        ") + '-' + RTRIM(CAST(TNUMBER AS nvarchar)) AS Expr1, NAME, RTRIM(CAST(N_KOL AS n"
                                        "varchar)) + '-' + RTRIM(CAST(NUMBER AS nvarchar)) + '-' + RTRIM(CAST(TNUMBER AS "
                                        "nvarchar)) AS Expr2 FROM TDETA_HES WHERE     (N_KOL = 521)"
                                    ColumnWidths ="0;2438"
                                    Begin
                                        Begin Label
                                            Visible = NotDefault
                                            OverlapFlags =215
                                            TextFontCharSet =0
                                            Left =7949
                                            Top =7929
                                            Width =1335
                                            Height =300
                                            Name ="Label35"
                                            Caption ="حساب سرويس:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    Visible = NotDefault
                                    DecimalPlaces =0
                                    OverlapFlags =223
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =1
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =2063
                                    Top =7936
                                    Width =621
                                    Height =315
                                    TabIndex =14
                                    Name ="SPER"
                                    Format ="Fixed"
                                    InputMask ="##\\%"
                                End
                                Begin TextBox
                                    Visible = NotDefault
                                    DecimalPlaces =0
                                    OverlapFlags =223
                                    TextAlign =3
                                    TextFontFamily =2
                                    ReadingOrder =2
                                    KeyboardLanguage =43
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =185
                                    Top =7935
                                    Width =1821
                                    Height =315
                                    FontWeight =700
                                    TabIndex =15
                                    Name ="MABL_HAZ"
                                    ControlSource ="MABL_HAZ"
                                    Format ="#,##0.00\" ريال \";-#,##0.00\" ريال \""
                                    FontName ="Titr"
                                    Begin
                                        Begin Label
                                            Visible = NotDefault
                                            OverlapFlags =215
                                            TextFontCharSet =0
                                            Left =2738
                                            Top =7936
                                            Width =1200
                                            Height =300
                                            Name ="Label33"
                                            Caption =" درصد سرويس:"
                                        End
                                    End
                                End
                                Begin CommandButton
                                    OverlapFlags =215
                                    AccessKey =223
                                    TextFontCharSet =0
                                    Left =1125
                                    Top =737
                                    Width =1545
                                    Height =465
                                    FontWeight =700
                                    TabIndex =16
                                    Name ="Command120"
                                    Caption ="چاپ فا&كتوركوچك"
                                    OnClick ="[Event Procedure]"
                                End
                                Begin CommandButton
                                    TabStop = NotDefault
                                    OverlapFlags =215
                                    AccessKey =141
                                    TextFontCharSet =0
                                    Left =60
                                    Top =737
                                    Width =1065
                                    Height =465
                                    FontWeight =700
                                    TabIndex =17
                                    Name ="Command100"
                                    Caption ="&چاپ فاكتور"
                                    OnClick ="[Event Procedure]"
                                End
                                Begin TextBox
                                    Visible = NotDefault
                                    TabStop = NotDefault
                                    DecimalPlaces =0
                                    OverlapFlags =247
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =1
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =6236
                                    Top =458
                                    Width =591
                                    Height =315
                                    TabIndex =18
                                    Name ="hTAG"
                                    ControlSource ="htag"
                                    DefaultValue ="13"
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    TabStop = NotDefault
                                    DecimalPlaces =0
                                    OverlapFlags =215
                                    TextAlign =3
                                    TextFontFamily =2
                                    ReadingOrder =2
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =190
                                    Top =7554
                                    Width =1821
                                    Height =315
                                    FontWeight =700
                                    TabIndex =19
                                    BackColor =13434828
                                    Name ="JJKOL"
                                    ControlSource ="=[INVO_LST_sub].[Form]![SMABLK]"
                                    Format ="#,##0.00\" ريال \";-#,##0.00\" ريال \""
                                    FontName ="Titr"
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            TextFontCharSet =0
                                            Left =2110
                                            Top =7554
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
                                    Left =3430
                                    Top =7569
                                    Width =1251
                                    Height =315
                                    TabIndex =20
                                    BackColor =16777164
                                    Name ="MOGU"
                                    Begin
                                        Begin Label
                                            Visible = NotDefault
                                            OverlapFlags =215
                                            TextFontCharSet =0
                                            Left =4780
                                            Top =7569
                                            Width =750
                                            Height =300
                                            Name ="Label63"
                                            Caption ="موجودي:"
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
                                    Left =1440
                                    Top =1530
                                    Width =1416
                                    Height =300
                                    TabIndex =21
                                    BackColor =32768
                                    ForeColor =65535
                                    Name ="N_S"
                                    ControlSource ="N_S"
                                    OnDblClick ="[Event Procedure]"
                                    Begin
                                        Begin Label
                                            BackStyle =1
                                            OverlapFlags =215
                                            TextFontCharSet =0
                                            Left =2865
                                            Top =1530
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
                                    Left =75
                                    Top =1215
                                    Width =2781
                                    Height =300
                                    TabIndex =22
                                    BackColor =32768
                                    ForeColor =65535
                                    Name ="MANDAH"
                                    Begin
                                        Begin Label
                                            BackStyle =1
                                            OverlapFlags =215
                                            TextFontCharSet =0
                                            Left =2865
                                            Top =1215
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
                                    TabStop = NotDefault
                                    DecimalPlaces =0
                                    SpecialEffect =0
                                    OverlapFlags =247
                                    TextAlign =3
                                    TextFontFamily =2
                                    ReadingOrder =1
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =170
                                    Top =7914
                                    Width =1881
                                    Height =300
                                    FontWeight =700
                                    TabIndex =23
                                    BackColor =16777164
                                    Name ="Text121"
                                    ControlSource ="=[GHABEL]-[NPAR]"
                                    Format ="#,##0.00;#,##0.00-"
                                    FontName ="Titr"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            Left =2139
                                            Top =7937
                                            Width =525
                                            Height =300
                                            Name ="Label122"
                                            Caption ="مانده:"
                                        End
                                    End
                                End
                                Begin CommandButton
                                    OverlapFlags =215
                                    TextFontCharSet =0
                                    Left =2670
                                    Top =737
                                    Width =396
                                    Height =465
                                    TabIndex =24
                                    Name ="Command123"
                                    Caption ="Command123"
                                    OnClick ="[Event Procedure]"
                                    PictureData = Begin
                                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                                        0xffff0000ffffff00dadadadadadadadaad00000000000dadd0888888888080da ,
                                        0x000000000000080d0888888bbb88000a088888877788080d0000000000000880 ,
                                        0x0888888888808080d000000000080800ad0ffffffff08080dad0f00000f0000a ,
                                        0xada0ffffffff0daddada0f00000f0adaadad0ffffffff0addadad000000000da ,
                                        0xadadadadadadadad000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x00000000000000000000000000000000000000000000000000000000
                                    End
                                    ObjectPalette = Begin
                                        0x0003100000000000800000000080000080800000000080008000800000808000 ,
                                        0x80808000c0c0c000ff000000c0c0c000ffff00000000ff00c0c0c00000ffff00 ,
                                        0xffffff0000000000
                                    End
                                    Tag ="چاپ فاكتور كوچك"
                                    ControlTipText ="Print Report"
                                End
                                Begin TextBox
                                    Visible = NotDefault
                                    TabStop = NotDefault
                                    DecimalPlaces =0
                                    OverlapFlags =247
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =1
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =5719
                                    Top =345
                                    Width =411
                                    Height =345
                                    TabIndex =25
                                    Name ="LETSANAD"
                                    DefaultValue ="False"
                                End
                                Begin TextBox
                                    Visible = NotDefault
                                    OverlapFlags =215
                                    IMESentenceMode =3
                                    Left =3288
                                    Top =822
                                    Width =276
                                    TabIndex =26
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
                                    Left =75
                                    Top =1530
                                    Width =1356
                                    Height =300
                                    TabIndex =27
                                    BackColor =32768
                                    ForeColor =65535
                                    Name ="MABNA"
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =247
                            AccessKey =129
                            Left =135
                            Top =450
                            Width =13815
                            Height =7860
                            Name ="Page58"
                            Caption ="&پشت فاكتور"
                            Begin
                                Begin TextBox
                                    DecimalPlaces =0
                                    OverlapFlags =255
                                    TextAlign =3
                                    TextFontFamily =2
                                    ReadingOrder =2
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =10597
                                    Top =1200
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
                                            Left =12336
                                            Top =1200
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
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =10597
                                    Top =1540
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
                                            Left =12336
                                            Top =1545
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
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    ColumnCount =2
                                    Left =4486
                                    Top =1545
                                    Width =4792
                                    Height =315
                                    TabIndex =2
                                    Name ="MOIN_VAR"
                                    ControlSource ="MOIN_VAR"
                                    RowSourceType ="Table/View/StoredProc"
                                    RowSource ="SELECT  RTRIM(CAST(N_KOL AS nvarchar)) + '-' + RTRIM(CAST(NUMBER AS nvarchar)) +"
                                        " '-' + RTRIM(CAST(TNUMBER AS nvarchar)) AS Expr1, NAME,   N_KOL FROM TDETA_HES W"
                                        "HERE     (N_KOL =112)"
                                    ColumnWidths ="851"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            TextFontCharSet =0
                                            Left =9319
                                            Top =1547
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
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =10597
                                    Top =1890
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
                                            Left =12336
                                            Top =1890
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
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    ColumnCount =2
                                    Left =4486
                                    Top =1892
                                    Width =4792
                                    Height =315
                                    TabIndex =4
                                    Name ="MOIN_HAV"
                                    ControlSource ="MOIN_HAV"
                                    RowSourceType ="Table/View/StoredProc"
                                    RowSource ="SELECT  RTRIM(CAST(N_KOL AS nvarchar)) + '-' + RTRIM(CAST(NUMBER AS nvarchar)) +"
                                        " '-' + RTRIM(CAST(TNUMBER AS nvarchar)) AS Expr1, NAME,   N_KOL FROM TDETA_HES W"
                                        "HERE     (N_KOL =112)"
                                    ColumnWidths ="851"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            TextFontCharSet =0
                                            Left =9319
                                            Top =1892
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
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =10597
                                    Top =2235
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
                                            Left =12336
                                            Top =2250
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
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =4486
                                    Top =2222
                                    Width =4792
                                    Height =315
                                    TabIndex =6
                                    Name ="takh"
                                    Format ="Fixed"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            TextFontCharSet =0
                                            Left =9319
                                            Top =2237
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
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =10595
                                    Top =2595
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
                                            Left =12330
                                            Top =2595
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
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    ColumnCount =3
                                    Left =4486
                                    Top =2581
                                    Width =4792
                                    Height =315
                                    TabIndex =8
                                    Name ="MOIN_HAZ2"
                                    ControlSource ="MOIN_HAZ"
                                    RowSourceType ="Table/View/StoredProc"
                                    RowSource ="SELECT RTRIM(CAST(N_KOL AS nvarchar)) + '-' + RTRIM(CAST(NUMBER AS nvarchar)) + "
                                        "'-' + RTRIM(CAST(TNUMBER AS nvarchar)) AS Expr1, NAME, RTRIM(CAST(N_KOL AS nvarc"
                                        "har)) + '-' + RTRIM(CAST(NUMBER AS nvarchar)) + '-' + RTRIM(CAST(TNUMBER AS nvar"
                                        "char)) AS Expr2 FROM TDETA_HES"
                                    ColumnWidths ="0"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            TextFontCharSet =0
                                            Left =9309
                                            Top =2580
                                            Width =1200
                                            Height =300
                                            Name ="Label117"
                                            Caption ="معين عوارض:"
                                        End
                                    End
                                End
                                Begin Rectangle
                                    OverlapFlags =255
                                    Left =169
                                    Top =495
                                    Width =13425
                                    Height =371
                                    Name ="Box80"
                                End
                                Begin Rectangle
                                    OverlapFlags =255
                                    Left =4384
                                    Top =1005
                                    Width =9210
                                    Height =2336
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
                                    Left =11283
                                    Top =546
                                    Width =1116
                                    Height =300
                                    TabIndex =9
                                    Name ="Text78"
                                    ControlSource ="NUMBER"
                                    DefaultValue ="0"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            TextFontCharSet =0
                                            Left =12408
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
                                    Left =9049
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
                                            Left =10189
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
                                    Left =289
                                    Top =546
                                    Width =7896
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
                                            Left =8194
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
                                    Left =170
                                    Top =3631
                                    Width =13515
                                    Height =3900
                                    TabIndex =12
                                    Name ="PAY_GETD_SUB"
                                    SourceObject ="Form.PAY_GETD_SUB"
                                    LinkChildFields ="NUMBER;TAG"
                                    LinkMasterFields ="NUMBER;DTAG"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            TextFontCharSet =0
                                            Left =12406
                                            Top =3344
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
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =281
                                    Top =2366
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
                                            Left =2487
                                            Top =2373
                                            Width =1545
                                            Height =300
                                            Name ="Label86"
                                            Caption ="جمع مبالغ پرداختي:"
                                        End
                                    End
                                End
                                Begin Rectangle
                                    OverlapFlags =255
                                    Left =174
                                    Top =963
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
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =282
                                    Top =1046
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
                                            Left =2492
                                            Top =1046
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
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =281
                                    Top =1376
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
                                            Left =2486
                                            Top =1376
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
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =281
                                    Top =1691
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
                                            Left =2486
                                            Top =1691
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
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =281
                                    Top =2021
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
                                            Left =2486
                                            Top =2028
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
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =281
                                    Top =2696
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
                                            Left =2487
                                            Top =2703
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
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =450
                                    Top =7590
                                    Width =2106
                                    Height =315
                                    TabIndex =19
                                    Name ="NCHK"
                                    ControlSource ="=IIf(IsNull([PAY_GETD_SUB].[Form]![SCMAB]) Or IsError([PAY_GETD_SUB].[Form]![SCM"
                                        "AB]),0,[PAY_GETD_SUB].[Form]![SCMAB])"
                                    Format ="#,##0.00\" ريال \";-#,##0.00\" ريال \""
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            Left =2649
                                            Top =7590
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
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =10602
                                    Top =2963
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
                                            Left =12342
                                            Top =2963
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
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    ColumnCount =2
                                    Left =4486
                                    Top =2949
                                    Width =4792
                                    Height =315
                                    TabIndex =21
                                    Name ="HMBAA"
                                    ControlSource ="HMBAA"
                                    RowSourceType ="Table/View/StoredProc"
                                    RowSource ="SELECT RTRIM(CAST(N_KOL AS nvarchar)) + '-' + RTRIM(CAST(NUMBER AS nvarchar)) + "
                                        "'-' + RTRIM(CAST(TNUMBER AS nvarchar)) AS Expr1, NAME, RTRIM(CAST(N_KOL AS nvarc"
                                        "har)) + '-' + RTRIM(CAST(NUMBER AS nvarchar)) + '-' + RTRIM(CAST(TNUMBER AS nvar"
                                        "char)) AS Expr2 FROM TDETA_HES"
                                    ColumnWidths ="854"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            Left =9321
                                            Top =2948
                                            Width =1200
                                            Height =300
                                            Name ="Label138"
                                            Caption ="معين ماليات:"
                                        End
                                    End
                                End
                            End
                        End
                    End
                End
                Begin CheckBox
                    Visible = NotDefault
                    OverlapFlags =255
                    ReadingOrder =1
                    Left =12545
                    Top =7597
                    TabIndex =1
                    Name ="SGN1"
                    ControlSource ="SGN1"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =12806
                            Top =7540
                            Width =675
                            Height =240
                            ForeColor =16711680
                            Name ="LSGN1"
                            Caption ="فروش"
                        End
                    End
                End
                Begin CheckBox
                    Visible = NotDefault
                    OverlapFlags =255
                    ReadingOrder =1
                    Left =12515
                    Top =7851
                    Width =290
                    TabIndex =2
                    Name ="SGN2"
                    ControlSource ="SGN2"
                    Begin
                        Begin Label
                            OverlapFlags =255
                            Left =12799
                            Top =7795
                            Width =975
                            Height =240
                            ForeColor =16711680
                            Name ="LSGN2"
                            Caption ="حسابداري"
                        End
                    End
                End
                Begin CheckBox
                    Visible = NotDefault
                    OverlapFlags =247
                    ReadingOrder =1
                    Left =12545
                    Top =8110
                    TabIndex =3
                    Name ="SGN3"
                    ControlSource ="SGN3"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =12799
                            Top =8047
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
    End
End
CodeBehindForm
' See "MHEAD_LST_FROOSH2.cls"
