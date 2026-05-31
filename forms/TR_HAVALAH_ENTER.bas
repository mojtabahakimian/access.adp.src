Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
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
    GridX =60
    GridY =60
    Width =14327
    DatasheetFontHeight =10
    ItemSuffix =109
    Left =210
    Right =14280
    Bottom =7995
    HelpContextId =6001
    DatasheetGridlinesColor =12632256
    MaxRecords =1000000
    RecSrcDt = Begin
        0x02ad0e1d53cde340
    End
    ServerFilter ="tag=9"
    RecordSource ="TR_HEAD_LST"
    Caption ="برگه ورود كالاي ساخته شده به انبار"
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
            Height =8862
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =11565
                    Top =254
                    Width =1326
                    Height =315
                    ColumnOrder =9
                    Name ="NUMBER"
                    ControlSource ="NUMBER"
                    DefaultValue ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =13035
                            Top =259
                            Width =1110
                            Height =300
                            Name ="Label1"
                            Caption ="شماره برگه:"
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
                    Left =9052
                    Top =226
                    Width =1221
                    Height =315
                    ColumnOrder =8
                    TabIndex =1
                    Name ="DATE_N"
                    ControlSource ="DATE_N"
                    DefaultValue ="=FARSIDATE(Date())"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =10417
                            Top =226
                            Width =1065
                            Height =300
                            Name ="Label9"
                            Caption ="تاريخ برگه:"
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
                    Left =9525
                    Top =604
                    Width =3366
                    Height =315
                    ColumnOrder =7
                    TabIndex =3
                    BackColor =16777164
                    Name ="CUST_NO"
                    ControlSource ="CUST_NO"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT RTRIM(CAST(N_KOL AS nvarchar)) + '-' + RTRIM(CAST(NUMBER AS nvarchar)) + "
                        "'-' + RTRIM(CAST(TNUMBER AS nvarchar)) AS hes, NAME AS nam FROM TDETA_HES"
                    ColumnWidths ="0;2835;567"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =13035
                            Top =600
                            Width =1230
                            Height =300
                            Name ="Label19"
                            Caption =" مسئول شيفت:"
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
                    Left =7290
                    Top =604
                    Width =2196
                    Height =315
                    ColumnOrder =6
                    TabIndex =4
                    BackColor =16777164
                    Name ="CUST_NO2"
                    ControlSource ="CUST_NO"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT RTRIM(CAST(N_KOL AS nvarchar)) + '-' + RTRIM(CAST(NUMBER AS nvarchar)) + "
                        "'-' + RTRIM(CAST(TNUMBER AS nvarchar)) AS Expr1 FROM TDETA_HES"
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =7305
                    Top =964
                    Width =5595
                    Height =315
                    ColumnOrder =5
                    TabIndex =5
                    Name ="MOLAH"
                    ControlSource ="MOLAH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =13050
                            Top =964
                            Width =795
                            Height =300
                            Name ="Label21"
                            Caption ="ملاحظات:"
                        End
                    End
                End
                Begin Subform
                    OverlapFlags =85
                    Left =53
                    Top =1810
                    Width =14190
                    Height =6450
                    TabIndex =6
                    Name ="HAVALAH_ENTER_SUB"
                    SourceObject ="Form.tr_HAVALAH_ENTER_SUB"
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
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3723
                    Top =274
                    Width =411
                    Height =315
                    ColumnOrder =4
                    TabIndex =7
                    Name ="hTAG"
                    ControlSource ="TAG"
                    DefaultValue ="9"
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    DecimalPlaces =0
                    SpecialEffect =0
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =9911
                    Top =1351
                    Width =2856
                    Height =300
                    ColumnOrder =2
                    TabIndex =8
                    BackColor =32768
                    ForeColor =65535
                    Name ="N_S"
                    ControlSource ="N_S"
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =12776
                            Top =1351
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
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =7290
                    Top =229
                    Width =1146
                    Height =315
                    ColumnOrder =10
                    TabIndex =9
                    ForeColor =8388608
                    Name ="USER_NAME"
                    ControlSource ="USER_NAME"
                    DefaultValue ="UCurrentUser()"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =8493
                            Top =226
                            Width =525
                            Height =300
                            ForeColor =8388608
                            Name ="Label111"
                            Caption ="كاربر:"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    TabStop = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3043
                    Top =274
                    Width =411
                    Height =345
                    ColumnOrder =1
                    TabIndex =10
                    Name ="LETSANAD"
                    DefaultValue ="False"
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =7320
                    Top =1365
                    Width =1131
                    Height =285
                    ColumnOrder =0
                    TabIndex =2
                    Name ="FNUMCO"
                    ControlSource ="FNUMCO"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextFontCharSet =0
                            Left =8472
                            Top =1365
                            Width =1740
                            Height =285
                            Name ="Label107"
                            Caption ="شماره برگه داخلي:"
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
                    Left =566
                    Top =297
                    Width =1431
                    Height =345
                    FontSize =10
                    FontWeight =700
                    TabIndex =11
                    BackColor =13408767
                    Name ="UP_DATE"
                    ControlSource ="UP_DATE"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2098
                            Top =294
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
                    Left =583
                    Top =744
                    Width =2841
                    Height =345
                    FontSize =10
                    FontWeight =700
                    TabIndex =12
                    BackColor =13408767
                    Name ="UP_TIME"
                    ControlSource ="UP_TIME"
                    Format ="General Date"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3540
                            Top =746
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
                    Left =566
                    Top =1184
                    Height =345
                    TabIndex =13
                    BackColor =13408767
                    Name ="UP_USER_NAME"
                    ControlSource ="UP_USER_NAME"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2318
                            Top =1184
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
                    Left =3160
                    Top =344
                    Width =1431
                    Height =345
                    FontSize =10
                    FontWeight =700
                    TabIndex =14
                    BackColor =13408767
                    Name ="PC_NAME"
                    ControlSource ="PC_NAME"
                    FontName ="Times New Roman"
                    InputMask ="##/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4691
                            Top =339
                            Width =1755
                            Height =240
                            Name ="Label152"
                            Caption ="نام كامپيوتر اصلاح كننده:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =1
                    TextFontFamily =18
                    ReadingOrder =2
                    ScrollBarAlign =2
                    IMESentenceMode =3
                    Left =4076
                    Top =1134
                    Width =1929
                    Height =345
                    FontSize =12
                    FontWeight =700
                    TabIndex =15
                    BackColor =13408767
                    Name ="IPADD"
                    ControlSource ="IPADD"
                    FontName ="Times New Roman"
                    InputMask ="##/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3727
                            Top =1186
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
' See "TR_HAVALAH_ENTER.cls"
