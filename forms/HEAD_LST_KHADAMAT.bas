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
    Width =14310
    DatasheetFontHeight =10
    ItemSuffix =128
    Left =2265
    Top =390
    Right =16575
    Bottom =9390
    HelpContextId =6001
    DatasheetGridlinesColor =12632256
    MaxRecords =100000000
    BeforeDelConfirm ="[Event Procedure]"
    RecSrcDt = Begin
        0x231680701dd1e340
    End
    ServerFilter ="TAG = 14"
    RecordSource ="SELECT HEAD_LST.* FROM HEAD_LST ORDER BY NUMBER"
    Caption ="فاكتور خدمات"
    OnCurrent ="[Event Procedure]"
    BeforeInsert ="[Event Procedure]"
    OnDelete ="[Event Procedure]"
    BeforeUpdate ="[Event Procedure]"
    AfterUpdate ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    HelpFile ="dena.hlp"
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
            Height =9990
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin Tab
                    OverlapFlags =85
                    TextFontCharSet =0
                    Style =1
                    Width =14310
                    Height =9990
                    FontWeight =700
                    Name ="TabCtl56"
                    Begin
                        Begin Page
                            OverlapFlags =215
                            AccessKey =209
                            Left =113
                            Top =405
                            Width =14062
                            Height =9454
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
                                    Left =11820
                                    Top =405
                                    Width =1371
                                    Height =315
                                    BackColor =16777164
                                    ForeColor =8388608
                                    Name ="NUMBER"
                                    ControlSource ="NUMBER"
                                    DefaultValue ="0"
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            TextFontCharSet =0
                                            Left =13275
                                            Top =410
                                            Width =675
                                            Height =300
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
                                    Left =9780
                                    Top =405
                                    Width =1371
                                    Height =315
                                    TabIndex =1
                                    ForeColor =8388608
                                    Name ="DATE_N"
                                    ControlSource ="DATE_N"
                                    BeforeUpdate ="[Event Procedure]"
                                    DefaultValue ="=FARSIDATE(Date())"
                                    InputMask ="####/##/##"
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            TextFontCharSet =0
                                            Left =11190
                                            Top =405
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
                                    Left =7845
                                    Top =423
                                    Width =1311
                                    Height =315
                                    TabIndex =2
                                    ForeColor =8388608
                                    Name ="USER_NAME"
                                    ControlSource ="USER_NAME"
                                    DefaultValue ="UCurrentUser()"
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            TextFontCharSet =0
                                            Left =9213
                                            Top =420
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
                                    OverlapFlags =215
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =1
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =6585
                                    Top =405
                                    Width =696
                                    Height =315
                                    TabIndex =3
                                    ForeColor =8388608
                                    Name ="MAS"
                                    ControlSource ="MAS"
                                    DefaultValue ="0"
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            TextFontCharSet =0
                                            Left =7305
                                            Top =405
                                            Width =450
                                            Height =300
                                            ForeColor =8388608
                                            Name ="Label106"
                                            Caption ="مدت:"
                                        End
                                    End
                                End
                                Begin ComboBox
                                    Locked = NotDefault
                                    OldBorderStyle =0
                                    OverlapFlags =215
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =1
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    ColumnCount =2
                                    Left =11475
                                    Top =780
                                    Width =1716
                                    Height =315
                                    TabIndex =4
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
                                            Left =13275
                                            Top =780
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
                                    Left =9615
                                    Top =780
                                    Width =1191
                                    Height =315
                                    TabIndex =5
                                    ForeColor =8388608
                                    Name ="SHIFT"
                                    ControlSource ="SHIFT"
                                    RowSourceType ="Table/View/StoredProc"
                                    RowSource ="SELECT SHIFT.SHIFT_ID, SHIFT.SHNAME FROM SHIFT ORDER BY SHIFT.SHNAME;"
                                    ColumnWidths ="0"
                                    DefaultValue ="=[Forms]![DEFAULT]![SHIFT]"
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            TextFontCharSet =0
                                            Left =10845
                                            Top =780
                                            Width =585
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
                                    Left =6945
                                    Top =780
                                    Width =1566
                                    Height =315
                                    TabIndex =6
                                    ForeColor =8388608
                                    Name ="CUST_KIND"
                                    ControlSource ="CUST_KIND"
                                    RowSourceType ="Table/View/StoredProc"
                                    RowSource ="SELECT CUSTKIND.CUST_COD, CUSTKIND.CUSTKNAME FROM CUSTKIND ORDER BY CUSTKIND.CUS"
                                        "TKNAME;"
                                    ColumnWidths ="0"
                                    ValidationRule ="Not Is Null"
                                    AfterUpdate ="[Event Procedure]"
                                    OnDblClick ="[Event Procedure]"
                                    DefaultValue ="[Forms]![BASEKNOW]![DEFTKH]"
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            TextFontCharSet =0
                                            Left =8550
                                            Top =780
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
                                    OldBorderStyle =0
                                    OverlapFlags =215
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =1
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    ColumnCount =2
                                    Left =8970
                                    Top =1141
                                    Width =4221
                                    Height =315
                                    TabIndex =7
                                    BackColor =13434828
                                    ForeColor =8388608
                                    Name ="CUST_NO"
                                    ControlSource ="CUST_NO"
                                    RowSourceType ="Table/View/StoredProc"
                                    RowSource ="SELECT hes, NAME AS nam, hes AS Expr1 FROM CUST_HESAB"
                                    ColumnWidths ="0;2835;567"
                                    AfterUpdate ="[Event Procedure]"
                                    OnDblClick ="[Event Procedure]"
                                    OnNotInList ="[Event Procedure]"
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            TextFontCharSet =0
                                            Left =13275
                                            Top =1141
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
                                    OldBorderStyle =0
                                    OverlapFlags =215
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =1
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =6945
                                    Top =1141
                                    Width =1926
                                    Height =315
                                    TabIndex =8
                                    BackColor =13434828
                                    ForeColor =8388608
                                    Name ="CUST_NO2"
                                    ControlSource ="CUST_NO"
                                    RowSourceType ="Table/View/StoredProc"
                                    RowSource ="SELECT hes FROM CUST_HESAB"
                                    AfterUpdate ="[Event Procedure]"
                                    OnNotInList ="[Event Procedure]"
                                End
                                Begin TextBox
                                    OverlapFlags =215
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =1
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =6945
                                    Top =1490
                                    Width =6255
                                    Height =315
                                    TabIndex =9
                                    ForeColor =8388608
                                    Name ="MOLAH"
                                    ControlSource ="MOLAH"
                                    AfterUpdate ="[Event Procedure]"
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            TextFontCharSet =0
                                            Left =13275
                                            Top =1490
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
                                    Left =113
                                    Top =1862
                                    Width =14010
                                    Height =6135
                                    TabIndex =10
                                    Name ="INVO_LST_KHAD"
                                    SourceObject ="Form.INVO_LST_KHAD"
                                    LinkChildFields ="NUMBER;TAG"
                                    LinkMasterFields ="NUMBER;htag"
                                    OnExit ="[Event Procedure]"
                                End
                                Begin ComboBox
                                    OldBorderStyle =0
                                    OverlapFlags =215
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =1
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    ColumnCount =3
                                    Left =2891
                                    Top =8901
                                    Width =3891
                                    Height =315
                                    TabIndex =11
                                    Name ="MOIN_HAZ"
                                    ControlSource ="MOIN_HAZ"
                                    RowSourceType ="Table/View/StoredProc"
                                    ColumnWidths ="0;2438"
                                    OnExit ="[Event Procedure]"
                                    Begin
                                        Begin Label
                                            OverlapFlags =223
                                            TextFontCharSet =0
                                            Left =6830
                                            Top =8900
                                            Width =1440
                                            Height =300
                                            Name ="Label35"
                                            Caption ="حساب سرويس:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    DecimalPlaces =0
                                    OverlapFlags =215
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =1
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =2233
                                    Top =8432
                                    Width =621
                                    Height =315
                                    TabIndex =12
                                    Name ="SPER"
                                    AfterUpdate ="[Event Procedure]"
                                    DefaultValue ="0"
                                    InputMask ="##\\%"
                                End
                                Begin TextBox
                                    DecimalPlaces =0
                                    OverlapFlags =215
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =1
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =355
                                    Top =8431
                                    Width =1821
                                    Height =315
                                    TabIndex =13
                                    Name ="MABL_HAZ"
                                    ControlSource ="MABL_HAZ"
                                    Format ="#,##0.00\" ريال \";-#,##0.00\" ريال \""
                                    AfterUpdate ="[Event Procedure]"
                                    OnKeyPress ="[Event Procedure]"
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            TextFontCharSet =0
                                            Left =2908
                                            Top =8432
                                            Width =1200
                                            Height =300
                                            Name ="Label33"
                                            Caption =" درصد سرويس:"
                                        End
                                    End
                                End
                                Begin CommandButton
                                    OverlapFlags =223
                                    AccessKey =223
                                    TextFontCharSet =0
                                    Left =1291
                                    Top =566
                                    Width =1530
                                    Height =465
                                    FontWeight =700
                                    TabIndex =14
                                    Name ="Command120"
                                    Caption ="چاپ فا&كتوركوچك"
                                    OnClick ="[Event Procedure]"
                                End
                                Begin TextBox
                                    Visible = NotDefault
                                    TabStop = NotDefault
                                    DecimalPlaces =0
                                    OverlapFlags =215
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =1
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =1808
                                    Top =694
                                    Width =291
                                    Height =315
                                    TabIndex =15
                                    Name ="hTAG"
                                    ControlSource ="TAG"
                                    DefaultValue ="14"
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    TabStop = NotDefault
                                    DecimalPlaces =0
                                    OverlapFlags =215
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =1
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =360
                                    Top =8050
                                    Width =1821
                                    Height =315
                                    TabIndex =16
                                    BackColor =13434828
                                    Name ="JJKOL"
                                    ControlSource ="=[INVO_LST_KHAD].[Form]![SMABLK]"
                                    Format ="#,##0.00\" ريال \";-#,##0.00\" ريال \""
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            TextFontCharSet =0
                                            Left =2280
                                            Top =8050
                                            Width =1260
                                            Height =300
                                            Name ="Label61"
                                            Caption ="جمع كل فاكتور:"
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
                                    Left =195
                                    Top =1425
                                    Width =2781
                                    Height =300
                                    TabIndex =17
                                    BackColor =32768
                                    ForeColor =65535
                                    Name ="N_S"
                                    ControlSource ="N_S"
                                    OnDblClick ="[Event Procedure]"
                                    Begin
                                        Begin Label
                                            BackStyle =1
                                            OverlapFlags =223
                                            TextFontCharSet =0
                                            Left =2985
                                            Top =1425
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
                                    Left =195
                                    Top =1110
                                    Width =2781
                                    Height =300
                                    TabIndex =18
                                    BackColor =32768
                                    ForeColor =65535
                                    Name ="MANDAH"
                                    OnDblClick ="[Event Procedure]"
                                    Begin
                                        Begin Label
                                            BackStyle =1
                                            OverlapFlags =223
                                            TextFontCharSet =0
                                            Left =2985
                                            Top =1110
                                            Width =1485
                                            Height =300
                                            BackColor =32768
                                            ForeColor =65535
                                            Name ="Label69"
                                            Caption ="مانده حساب:"
                                        End
                                    End
                                End
                                Begin CommandButton
                                    TabStop = NotDefault
                                    OverlapFlags =215
                                    AccessKey =141
                                    TextFontCharSet =0
                                    Left =226
                                    Top =566
                                    Width =1065
                                    Height =465
                                    FontWeight =700
                                    TabIndex =19
                                    Name ="Command100"
                                    Caption ="&چاپ فاكتور"
                                    OnClick ="[Event Procedure]"
                                End
                                Begin TextBox
                                    TabStop = NotDefault
                                    DecimalPlaces =0
                                    SpecialEffect =0
                                    OverlapFlags =215
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =1
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =340
                                    Top =8815
                                    Width =1866
                                    Height =300
                                    TabIndex =20
                                    BackColor =16777164
                                    Name ="Text121"
                                    ControlSource ="=[GHABEL]-[NPAR]"
                                    Format ="#,##0.00;#,##0.00-"
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            TextFontCharSet =0
                                            Left =2306
                                            Top =8822
                                            Width =525
                                            Height =300
                                            Name ="Label122"
                                            Caption ="مانده:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    Visible = NotDefault
                                    TabStop = NotDefault
                                    DecimalPlaces =0
                                    OverlapFlags =223
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =1
                                    KeyboardLanguage =43
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =2935
                                    Top =592
                                    Width =411
                                    Height =345
                                    TabIndex =21
                                    Name ="LETSANAD"
                                    DefaultValue ="False"
                                End
                                Begin CheckBox
                                    OverlapFlags =215
                                    ReadingOrder =1
                                    Left =4762
                                    Top =1471
                                    Width =340
                                    Height =218
                                    TabIndex =22
                                    Name ="TICMBAA"
                                    ControlSource ="TICMBAA"
                                    AfterUpdate ="[Event Procedure]"
                                    Begin
                                        Begin Label
                                            OverlapFlags =223
                                            Left =5155
                                            Top =1417
                                            Width =1515
                                            Height =240
                                            ForeColor =255
                                            Name ="Label142"
                                            Caption ="محاسبه.ماليات.ب.ا.ا"
                                        End
                                    End
                                End
                                Begin CheckBox
                                    Locked = NotDefault
                                    OverlapFlags =215
                                    ReadingOrder =1
                                    Left =4755
                                    Top =1081
                                    Width =340
                                    Height =218
                                    TabIndex =23
                                    Name ="OKF"
                                    ControlSource ="OKF"
                                    DefaultValue ="True"
                                    Begin
                                        Begin Label
                                            OverlapFlags =223
                                            TextFontCharSet =0
                                            Left =5145
                                            Top =1027
                                            Width =975
                                            Height =240
                                            ForeColor =255
                                            Name ="Label146"
                                            Caption ="تائيد فاكتور"
                                        End
                                    End
                                End
                                Begin CommandButton
                                    TabStop = NotDefault
                                    OverlapFlags =223
                                    TextFontCharSet =0
                                    Left =5209
                                    Top =510
                                    Width =1155
                                    Height =465
                                    FontWeight =700
                                    TabIndex =24
                                    Name ="ESLAH"
                                    Caption ="اصلاح فاكتور"
                                    OnClick ="[Event Procedure]"
                                End
                                Begin CommandButton
                                    TabStop = NotDefault
                                    OverlapFlags =255
                                    AccessKey =134
                                    TextFontCharSet =0
                                    Left =2835
                                    Top =570
                                    Width =1170
                                    Height =465
                                    FontWeight =700
                                    TabIndex =25
                                    Name ="Command139"
                                    Caption ="&چاپ فاكتور1"
                                    OnClick ="[Event Procedure]"
                                    UnicodeAccessKey =1670
                                End
                                Begin Subform
                                    OverlapFlags =247
                                    Left =8280
                                    Top =8288
                                    Width =5882
                                    Height =1251
                                    TabIndex =26
                                    Name ="TAKHFIF_APLAY_SUB_khad"
                                    SourceObject ="Form.TAKHFIF_APLAY_SUB_khad"
                                    LinkChildFields ="number;kind"
                                    LinkMasterFields ="number;htag"
                                    Begin
                                        Begin Label
                                            OverlapFlags =223
                                            TextAlign =3
                                            Left =12887
                                            Top =8055
                                            Width =1275
                                            Height =240
                                            Name ="Label330"
                                            Caption ="تخفيفات پيشرفته"
                                        End
                                    End
                                End
                                Begin TextBox
                                    Visible = NotDefault
                                    OverlapFlags =247
                                    IMESentenceMode =3
                                    Left =3004
                                    Top =793
                                    TabIndex =27
                                    Name ="CDDATE"
                                    ControlSource ="CDDATE"
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            Left =5671
                                            Top =793
                                            Width =735
                                            Height =240
                                            Name ="Label123"
                                            Caption ="CDDATE:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    Visible = NotDefault
                                    OverlapFlags =215
                                    IMESentenceMode =3
                                    Left =3004
                                    Top =1133
                                    TabIndex =28
                                    Name ="CDTIME"
                                    ControlSource ="CDTIME"
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            Left =5701
                                            Top =1133
                                            Width =705
                                            Height =240
                                            Name ="Label124"
                                            Caption ="CDTIME:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    Visible = NotDefault
                                    OverlapFlags =215
                                    IMESentenceMode =3
                                    Left =3004
                                    Top =1474
                                    TabIndex =29
                                    Name ="OKDATE"
                                    ControlSource ="OKDATE"
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            Left =5671
                                            Top =1474
                                            Width =735
                                            Height =240
                                            Name ="Label125"
                                            Caption ="OKDATE:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    Visible = NotDefault
                                    OverlapFlags =223
                                    IMESentenceMode =3
                                    Left =3004
                                    Top =1814
                                    TabIndex =30
                                    Name ="OKTIME"
                                    ControlSource ="OKTIME"
                                    Begin
                                        Begin Label
                                            OverlapFlags =223
                                            Left =5701
                                            Top =1814
                                            Width =705
                                            Height =240
                                            Name ="Label126"
                                            Caption ="OKTIME:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    Visible = NotDefault
                                    OverlapFlags =215
                                    TextFontCharSet =0
                                    IMESentenceMode =3
                                    Left =8280
                                    Top =9619
                                    TabIndex =31
                                    Name ="UID"
                                    ControlSource ="UID"
                                    DefaultValue ="=UCurrentUserIDD()"
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =247
                            AccessKey =129
                            Left =60
                            Top =450
                            Width =14115
                            Height =9405
                            Name ="Page58"
                            Caption ="&پشت فاكتور"
                            Begin
                                Begin TextBox
                                    DecimalPlaces =0
                                    OverlapFlags =255
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =2
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =11108
                                    Top =1140
                                    Width =1521
                                    Height =315
                                    Name ="M_NAGHD"
                                    ControlSource ="M_NAGHD"
                                    Format ="#,##0.00\" ريال \";-#,##0.00\" ريال \""
                                    AfterUpdate ="[Event Procedure]"
                                    OnKeyPress ="[Event Procedure]"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            TextFontCharSet =0
                                            Left =12666
                                            Top =1140
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
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =2
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =11108
                                    Top =1480
                                    Width =1521
                                    Height =315
                                    TabIndex =1
                                    Name ="MABL_VAR"
                                    ControlSource ="MABL_VAR"
                                    Format ="#,##0.00\" ريال \";-#,##0.00\" ريال \""
                                    AfterUpdate ="[Event Procedure]"
                                    OnKeyPress ="[Event Procedure]"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            TextFontCharSet =0
                                            Left =12666
                                            Top =1485
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
                                    Left =4590
                                    Top =1485
                                    Width =5137
                                    Height =315
                                    TabIndex =2
                                    Name ="MOIN_VAR"
                                    ControlSource ="MOIN_VAR"
                                    RowSourceType ="Table/View/StoredProc"
                                    RowSource ="SELECT hes, NAME, hes FROM CUST_HESAB ORDER BY NAME, hes"
                                    ColumnWidths ="851"
                                    OnExit ="[Event Procedure]"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            TextFontCharSet =0
                                            Left =9769
                                            Top =1487
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
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =2
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =11108
                                    Top =1830
                                    Width =1521
                                    Height =315
                                    TabIndex =3
                                    Name ="MABL_HAV"
                                    ControlSource ="MABL_HAV"
                                    Format ="#,##0.00\" ريال \";-#,##0.00\" ريال \""
                                    AfterUpdate ="[Event Procedure]"
                                    OnKeyPress ="[Event Procedure]"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            TextFontCharSet =0
                                            Left =12666
                                            Top =1830
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
                                    Left =4590
                                    Top =1832
                                    Width =5137
                                    Height =315
                                    TabIndex =4
                                    Name ="MOIN_HAV"
                                    ControlSource ="MOIN_HAV"
                                    RowSourceType ="Table/View/StoredProc"
                                    RowSource ="SELECT hes, NAME, hes FROM CUST_HESAB ORDER BY NAME, hes"
                                    ColumnWidths ="851"
                                    OnExit ="[Event Procedure]"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            TextFontCharSet =0
                                            Left =9769
                                            Top =1832
                                            Width =975
                                            Height =300
                                            Name ="Label31"
                                            Caption ="معين حواله:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    DecimalPlaces =0
                                    OverlapFlags =255
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =2
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =11108
                                    Top =2175
                                    Width =1521
                                    Height =315
                                    TabIndex =5
                                    Name ="TAKHFIF"
                                    ControlSource ="TAKHFIF"
                                    Format ="#,##0.00\" ريال \";-#,##0.00\" ريال \""
                                    AfterUpdate ="[Event Procedure]"
                                    OnKeyPress ="[Event Procedure]"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            TextFontCharSet =0
                                            Left =12666
                                            Top =2190
                                            Width =990
                                            Height =300
                                            Name ="Label37"
                                            Caption ="مبلغ تخفيف:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =255
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =1
                                    KeyboardLanguage =43
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =4590
                                    Top =2162
                                    Width =5137
                                    Height =315
                                    TabIndex =6
                                    Name ="takh"
                                    Format ="Fixed"
                                    AfterUpdate ="[Event Procedure]"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            TextFontCharSet =0
                                            Left =9769
                                            Top =2177
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
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =2
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =11108
                                    Top =2535
                                    Width =1521
                                    Height =315
                                    TabIndex =7
                                    Name ="MABL_HAZ2"
                                    ControlSource ="MABL_HAZ"
                                    Format ="#,##0.00\" ريال \";-#,##0.00\" ريال \""
                                    AfterUpdate ="[Event Procedure]"
                                    OnKeyPress ="[Event Procedure]"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            TextFontCharSet =0
                                            Left =12660
                                            Top =2535
                                            Width =1095
                                            Height =300
                                            Name ="Label115"
                                            Caption =" مبلغ سرويس:"
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
                                    ColumnCount =3
                                    Left =4590
                                    Top =2521
                                    Width =5137
                                    Height =315
                                    TabIndex =8
                                    Name ="MOIN_HAZ2"
                                    ControlSource ="MOIN_HAZ"
                                    RowSourceType ="Table/View/StoredProc"
                                    RowSource ="SELECT hes, NAME, hes FROM CUST_HESAB ORDER BY NAME, hes"
                                    ColumnWidths ="0"
                                    OnExit ="[Event Procedure]"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            TextFontCharSet =0
                                            Left =9759
                                            Top =2520
                                            Width =1215
                                            Height =240
                                            Name ="Label117"
                                            Caption ="حساب سرويس:"
                                        End
                                    End
                                End
                                Begin Rectangle
                                    OverlapFlags =255
                                    Left =180
                                    Top =495
                                    Width =13575
                                    Height =371
                                    Name ="Box80"
                                End
                                Begin Rectangle
                                    OverlapFlags =255
                                    Left =4485
                                    Top =1005
                                    Width =9690
                                    Height =2276
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
                                    Left =9210
                                    Top =546
                                    Width =1131
                                    Height =300
                                    TabIndex =9
                                    Name ="Text73"
                                    ControlSource ="DATE_N"
                                    InputMask ="####/##/##"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            Left =10350
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
                                    Left =345
                                    Top =546
                                    Width =8031
                                    Height =300
                                    TabIndex =10
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
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            Left =8385
                                            Top =540
                                            Width =765
                                            Height =300
                                            Name ="Label76"
                                            Caption =" مشتري:"
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
                                    Left =11444
                                    Top =546
                                    Width =1116
                                    Height =300
                                    TabIndex =11
                                    Name ="Text78"
                                    ControlSource ="NUMBER"
                                    DefaultValue ="0"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            Left =12569
                                            Top =544
                                            Width =1110
                                            Height =300
                                            Name ="Label79"
                                            Caption ="شماره فاكتور:"
                                        End
                                    End
                                End
                                Begin Subform
                                    OverlapFlags =247
                                    Left =60
                                    Top =3751
                                    Width =13680
                                    Height =4980
                                    TabIndex =12
                                    Name ="PAY_GETD_KHAD_SUB"
                                    SourceObject ="Form.PAY_GETD_KHAD_SUB"
                                    LinkChildFields ="NUMBER;TAG"
                                    LinkMasterFields ="NUMBER;htag"
                                    OnExit ="[Event Procedure]"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            Left =12462
                                            Top =3401
                                            Width =1200
                                            Height =300
                                            Name ="PAY_GETD_SUB Label"
                                            Caption ="چكهاي دريافتي"
                                            EventProcPrefix ="PAY_GETD_SUB_Label"
                                        End
                                    End
                                End
                                Begin TextBox
                                    DecimalPlaces =0
                                    SpecialEffect =0
                                    OverlapFlags =255
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =2
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =277
                                    Top =2423
                                    Width =2106
                                    Height =300
                                    TabIndex =13
                                    Name ="NPAR"
                                    ControlSource ="=[M_NAGHD]+[MABL_VAR]+[MABL_HAV]+[NCHK]"
                                    Format ="#,##0.00;-#,##0.00"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            TextFontCharSet =0
                                            Left =2489
                                            Top =2430
                                            Width =1425
                                            Height =240
                                            Name ="Label86"
                                            Caption ="جمع مبالغ پرداختي:"
                                        End
                                    End
                                End
                                Begin Rectangle
                                    OverlapFlags =255
                                    Left =170
                                    Top =1020
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
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =2
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =278
                                    Top =1103
                                    Width =2106
                                    Height =300
                                    TabIndex =14
                                    Name ="JF"
                                    ControlSource ="=[INVO_LST_KHAD].[Form]![SMABLK]"
                                    Format ="#,##0.00;-#,##0.00"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            Left =2489
                                            Top =1103
                                            Width =1110
                                            Height =240
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
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =2
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =277
                                    Top =1433
                                    Width =2106
                                    Height =300
                                    TabIndex =15
                                    Name ="HKH"
                                    ControlSource ="=[MABL_HAZ]"
                                    Format ="#,##0.00;-#,##0.00"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            Left =2489
                                            Top =1433
                                            Width =1140
                                            Height =240
                                            Name ="Label92"
                                            Caption ="هزينه سرويس:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    TabStop = NotDefault
                                    DecimalPlaces =0
                                    SpecialEffect =0
                                    OverlapFlags =247
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =2
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =270
                                    Top =1785
                                    Width =2106
                                    Height =300
                                    TabIndex =16
                                    Name ="NTKHFIF"
                                    ControlSource ="=[TAKHFIF]"
                                    Format ="#,##0.00;-#,##0.00"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            Left =2489
                                            Top =1785
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
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =2
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =270
                                    Top =2115
                                    Width =2106
                                    Height =300
                                    TabIndex =17
                                    BackColor =16777164
                                    Name ="GHABEL"
                                    ControlSource ="=[JF]+[HKH]-[NTKHFIF]+[MBAA]"
                                    Format ="#,##0.00;-#,##0.00"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            Left =2489
                                            Top =2122
                                            Width =1305
                                            Height =240
                                            Name ="Label97"
                                            Caption ="مبلغ قابل پرداخت:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    DecimalPlaces =0
                                    SpecialEffect =0
                                    OverlapFlags =247
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =2
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =277
                                    Top =2753
                                    Width =2106
                                    Height =300
                                    TabIndex =18
                                    BackColor =16777164
                                    Name ="MAN"
                                    ControlSource ="=[GHABEL]-[NPAR]"
                                    Format ="#,##0.00;#,##0.00-"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            Left =2489
                                            Top =2760
                                            Width =480
                                            Height =240
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
                                    ReadingOrder =2
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =340
                                    Top =8787
                                    Width =2106
                                    Height =315
                                    TabIndex =19
                                    Name ="NCHK"
                                    ControlSource ="=IIf(IsNull([PAY_GETD_KHAD_SUB].[Form]![SCMAB]),0,[PAY_GETD_KHAD_SUB].[Form]![SC"
                                        "MAB])"
                                    Format ="#,##0.00\" ريال \";-#,##0.00\" ريال \""
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            Left =2539
                                            Top =8787
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
                                    Left =11108
                                    Top =2895
                                    Width =1521
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
                                            Left =12675
                                            Top =2895
                                            Width =1290
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
                                    Left =4590
                                    Top =2896
                                    Width =5137
                                    Height =315
                                    TabIndex =21
                                    Name ="HMBAA"
                                    ControlSource ="HMBAA"
                                    RowSourceType ="Table/View/StoredProc"
                                    RowSource ="SELECT hes, NAME, hes FROM CUST_HESAB ORDER BY NAME, hes"
                                    ColumnWidths ="854"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            Left =9729
                                            Top =2895
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
            End
        End
    End
End
CodeBehindForm
' See "HEAD_LST_KHADAMAT.cls"
