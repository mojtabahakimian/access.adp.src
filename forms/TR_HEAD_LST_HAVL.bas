Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    KeyPreview = NotDefault
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
    Width =14456
    DatasheetFontHeight =10
    ItemSuffix =110
    Left =135
    Right =14340
    Bottom =7995
    HelpContextId =6002
    DatasheetGridlinesColor =12632256
    MaxRecords =10000000
    Filter ="((NUMBER=2431))"
    OrderBy ="UP_TIME DESC"
    RecSrcDt = Begin
        0xf4c8f8d531d7e340
    End
    RecordSource ="SELECT TOP 100 PERCENT NUMBER, TAG AS htag, TAG AS DTAG, ANBAR, NUMBER1, DATE_N,"
        " TAH, MAS, VAS, N_S, CUST_NO, MOLAH, M_NAGHD, MABL_VAR, MOIN_VAR, MABL_HAV, MOIN"
        "_HAV, MABL_HAZ, MOIN_HAZ, TAKHFIF, MOIN_KHF, ANBARF, FNUMCO, DEPATMAN, SHIFT, CU"
        "ST_KIND, USER_NAME, SGN1, SGN2, SGN3, SGN4, MBAA, HMBAA, TAMIR, TICMBAA, TKHF, O"
        "KF, UP_DATE, UP_TIME, UP_USER_NAME, PC_NAME, IPADD FROM TR_HEAD_LST WHERE (TAG ="
        " 2) ORDER BY NUMBER, UP_DATE, UP_TIME"
    Caption ="حواله  انبار فروش"
    HelpFile ="dena.hlp"
    DatasheetFontName ="Arial (Arabic)"
    OnKeyPress ="[Event Procedure]"
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
            Height =9524
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
                    Left =11685
                    Top =236
                    Width =1356
                    Height =315
                    ColumnOrder =22
                    BackColor =16777164
                    ForeColor =8388608
                    Name ="NUMBER"
                    ControlSource ="NUMBER"
                    DefaultValue ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =13185
                            Top =241
                            Width =1110
                            Height =300
                            Name ="Label1"
                            Caption ="شماره حواله:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =9825
                    Top =237
                    Width =951
                    Height =315
                    ColumnOrder =21
                    TabIndex =1
                    BackColor =13434828
                    ForeColor =8388608
                    Name ="DATE_N"
                    ControlSource ="DATE_N"
                    DefaultValue ="=farsidate(Date())"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextFontCharSet =0
                            Left =10785
                            Top =237
                            Width =855
                            Height =240
                            Name ="Label9"
                            Caption ="تاريخ حواله:"
                        End
                    End
                End
                Begin Subform
                    OverlapFlags =87
                    Left =56
                    Top =1485
                    Width =14265
                    Height =7185
                    TabIndex =7
                    Name ="INVO_LST_HAVL_SUB"
                    SourceObject ="Form.tr_INVO_LST_HAVL_SUB"
                    LinkChildFields ="NUMBER;TAG;UP_DATE;UP_TIME"
                    LinkMasterFields ="NUMBER;htag;UP_DATE;UP_TIME"
                End
                Begin TextBox
                    Visible = NotDefault
                    TabStop = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =6349
                    Width =291
                    Height =315
                    ColumnOrder =15
                    TabIndex =8
                    Name ="hTAG"
                    ControlSource ="htag"
                    DefaultValue ="2"
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =850
                    Top =8794
                    Width =966
                    Height =315
                    ColumnOrder =14
                    TabIndex =9
                    BackColor =13434828
                    Name ="Text59"
                    ControlSource ="=[INVO_LST_HAVL_SUB].[Form]![SMABLK]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =1915
                            Top =8794
                            Width =1020
                            Height =300
                            Name ="Label61"
                            Caption ="جمع تحويلي:"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4338
                    Top =8799
                    Width =1881
                    Height =315
                    ColumnOrder =13
                    TabIndex =10
                    BackColor =16777164
                    Name ="MOGU"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =6318
                            Top =8799
                            Width =750
                            Height =300
                            Name ="Label63"
                            Caption ="موجودي:"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =93
                    TextFontCharSet =0
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3348
                    Width =291
                    Height =315
                    ColumnOrder =12
                    TabIndex =11
                    Name ="N_S"
                    ControlSource ="N_S"
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =9993
                    Top =615
                    Width =3066
                    Height =315
                    ColumnOrder =19
                    TabIndex =3
                    BackColor =13434828
                    ForeColor =8388608
                    Name ="TAH"
                    ControlSource ="TAH"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT HEAD_LST.TAH FROM HEAD_LST GROUP BY HEAD_LST.TAH ORDER BY HEAD_LST.TAH;"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =13200
                            Top =615
                            Width =1170
                            Height =300
                            Name ="Label105"
                            Caption ="تحويل دهنده:"
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
                    Left =5223
                    Top =634
                    Width =3480
                    Height =315
                    ColumnOrder =18
                    TabIndex =4
                    BackColor =13434828
                    ForeColor =8388608
                    Name ="MOLAH"
                    ControlSource ="MOLAH"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT MOLAH FROM HEAD_LST GROUP BY MOLAH ORDER BY MOLAH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =8748
                            Top =630
                            Width =1200
                            Height =300
                            Name ="Label108"
                            Caption ="توسط:"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =8145
                    Top =228
                    Width =1041
                    Height =315
                    ColumnOrder =23
                    TabIndex =12
                    ForeColor =8388608
                    Name ="USER_NAME"
                    ControlSource ="USER_NAME"
                    DefaultValue ="=uCurrentUser()"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextFontCharSet =0
                            Left =9243
                            Top =225
                            Width =525
                            Height =300
                            ForeColor =8388608
                            Name ="Label111"
                            Caption ="كاربر:"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =8793
                    Top =990
                    Width =4266
                    Height =315
                    ColumnOrder =17
                    TabIndex =5
                    BackColor =13434828
                    ForeColor =8388608
                    Name ="CUST_NO"
                    ControlSource ="CUST_NO"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT RTRIM(CAST(N_KOL AS nvarchar)) + '-' + RTRIM(CAST(NUMBER AS nvarchar)) + "
                        "'-' + RTRIM(CAST(TNUMBER AS nvarchar)) AS hes, NAME AS nam, RTRIM(CAST(N_KOL AS "
                        "nvarchar)) + '-' + RTRIM(CAST(NUMBER AS nvarchar)) + '-' + RTRIM(CAST(TNUMBER AS"
                        " nvarchar)) AS Expr1 FROM TDETA_HES"
                    ColumnWidths ="0;2835;567"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =13203
                            Top =990
                            Width =780
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
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =6783
                    Top =1005
                    Width =1926
                    Height =315
                    ColumnOrder =16
                    TabIndex =6
                    BackColor =13434828
                    ForeColor =8388608
                    Name ="CUST_NO2"
                    ControlSource ="CUST_NO"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT RTRIM(CAST(N_KOL AS nvarchar)) + '-' + RTRIM(CAST(NUMBER AS nvarchar)) + "
                        "'-' + RTRIM(CAST(TNUMBER AS nvarchar)) AS Expr1 FROM TDETA_HES"
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =255
                    TextFontCharSet =0
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5955
                    Top =225
                    Width =1011
                    Height =315
                    ColumnOrder =20
                    TabIndex =2
                    BackColor =16777164
                    ForeColor =8388608
                    Name ="FNUMCO"
                    ControlSource ="FNUMCO"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            ReadingOrder =2
                            NumeralShapes =2
                            Left =6987
                            Top =225
                            Width =1065
                            Height =300
                            Name ="Label103"
                            Caption ="شماره داخلي:"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    Visible = NotDefault
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =247
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =8223
                    Width =1716
                    Height =315
                    ColumnOrder =10
                    TabIndex =13
                    ForeColor =8388608
                    Name ="DEPATMAN"
                    ControlSource ="DEPATMAN"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT DEPART.DEPATMAN, DEPART.DEPNAME FROM DEPART ORDER BY DEPART.DEPNAME;"
                    ColumnWidths ="0"
                    DefaultValue ="=[Forms]![DEFAULT]![TFSAZMAN]"
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =247
                            TextFontCharSet =0
                            Left =10023
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
                    Visible = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =255
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =3121
                    Width =1191
                    Height =315
                    ColumnOrder =11
                    TabIndex =14
                    ForeColor =8388608
                    Name ="SHIFT"
                    ControlSource ="SHIFT"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT SHIFT.SHIFT_ID, SHIFT.SHNAME FROM SHIFT ORDER BY SHIFT.SHNAME;"
                    ColumnWidths ="0"
                    DefaultValue ="=[Forms]![DEFAULT]![SHIFT]"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextFontCharSet =0
                            Left =4351
                            Width =585
                            Height =300
                            ForeColor =8388608
                            Name ="Label109"
                            Caption ="شيفت :"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    Visible = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =3798
                    Top =1020
                    Width =1566
                    Height =315
                    ColumnOrder =9
                    TabIndex =15
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
                            OverlapFlags =93
                            TextFontCharSet =0
                            Left =5403
                            Top =1020
                            Width =1020
                            Height =300
                            ForeColor =8388608
                            Name ="Label113"
                            Caption ="نوع مشتري :"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    ReadingOrder =1
                    Left =10480
                    Top =8731
                    ColumnOrder =8
                    TabIndex =16
                    Name ="SGN1"
                    ControlSource ="SGN1"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =10741
                            Top =8674
                            Width =675
                            Height =240
                            ForeColor =16711680
                            Name ="LSGN1"
                            Caption ="فروش"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    ReadingOrder =1
                    Left =10480
                    Top =8985
                    ColumnOrder =6
                    TabIndex =17
                    Name ="SGN2"
                    ControlSource ="SGN2"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =10734
                            Top =8929
                            Width =975
                            Height =240
                            ForeColor =16711680
                            Name ="LSGN2"
                            Caption ="انبار"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    ReadingOrder =1
                    Left =10480
                    Top =9244
                    ColumnOrder =7
                    TabIndex =18
                    Name ="SGN3"
                    ControlSource ="SGN3"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =10734
                            Top =9181
                            Width =960
                            Height =240
                            ForeColor =16711680
                            Name ="LSGN3"
                            Caption ="مدير عامل"
                        End
                    End
                End
                Begin CheckBox
                    Locked = NotDefault
                    OverlapFlags =93
                    ReadingOrder =1
                    Left =4025
                    Top =677
                    Width =340
                    Height =218
                    ColumnOrder =5
                    TabIndex =19
                    Name ="OKF"
                    ControlSource ="OKF"
                    DefaultValue ="False"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =4415
                            Top =623
                            Width =795
                            Height =240
                            ForeColor =255
                            Name ="Label146"
                            Caption ="تائيد حواله"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =1
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =75
                    Top =168
                    Width =1431
                    Height =345
                    ColumnOrder =0
                    FontSize =10
                    FontWeight =700
                    TabIndex =20
                    BackColor =13408767
                    Name ="UP_DATE"
                    ControlSource ="UP_DATE"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1607
                            Top =165
                            Width =900
                            Height =240
                            Name ="Label148"
                            Caption ="تاريخ اصلاح:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =1
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =92
                    Top =615
                    Width =2841
                    Height =345
                    ColumnOrder =4
                    FontSize =10
                    FontWeight =700
                    TabIndex =21
                    BackColor =13408767
                    Name ="UP_TIME"
                    ControlSource ="UP_TIME"
                    Format ="General Date"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =3049
                            Top =617
                            Width =1050
                            Height =240
                            Name ="Label149"
                            Caption ="ساعت اصلاح:"
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
                    Left =75
                    Top =1055
                    Height =345
                    ColumnOrder =3
                    TabIndex =22
                    BackColor =13408767
                    Name ="UP_USER_NAME"
                    ControlSource ="UP_USER_NAME"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1827
                            Top =1055
                            Width =1290
                            Height =240
                            Name ="Label150"
                            Caption ="كاربر اصلاح كننده:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =247
                    TextFontCharSet =0
                    TextAlign =1
                    TextFontFamily =18
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2669
                    Top =215
                    Width =1431
                    Height =345
                    ColumnOrder =2
                    FontSize =10
                    FontWeight =700
                    TabIndex =23
                    BackColor =13408767
                    Name ="PC_NAME"
                    ControlSource ="PC_NAME"
                    FontName ="Times New Roman"
                    InputMask ="##/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =4200
                            Top =210
                            Width =1755
                            Height =240
                            Name ="Label152"
                            Caption ="نام كامپيوتر اصلاح كننده:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =247
                    TextFontCharSet =0
                    TextAlign =1
                    TextFontFamily =18
                    ReadingOrder =2
                    ScrollBarAlign =2
                    IMESentenceMode =3
                    Left =3510
                    Top =1005
                    Width =1929
                    Height =345
                    ColumnOrder =1
                    FontSize =12
                    FontWeight =700
                    TabIndex =24
                    BackColor =13408767
                    Name ="IPADD"
                    ControlSource ="IPADD"
                    FontName ="Times New Roman"
                    InputMask ="##/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3161
                            Top =1057
                            Width =285
                            Height =240
                            Name ="Label155"
                            Caption ="IP:"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "TR_HEAD_LST_HAVL.cls"
