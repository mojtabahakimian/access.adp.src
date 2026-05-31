Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    DividingLines = NotDefault
    DefaultView =0
    AllowUpdating =4
    ViewsAllowed =1
    TabularCharSet =178
    TabularFamily =42
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =60
    GridY =60
    Width =15326
    DatasheetFontHeight =10
    ItemSuffix =111
    Left =3195
    Top =1125
    Right =18285
    Bottom =10245
    HelpContextId =6001
    DatasheetGridlinesColor =12632256
    BeforeDelConfirm ="[Event Procedure]"
    RecSrcDt = Begin
        0x40c829a38de2e440
    End
    RecordSource ="SELECT HEAD_LST.* FROM HEAD_LST WHERE (TAG = 9) ORDER BY NUMBER"
    Caption ="برگه ورود كالاي ساخته شده به انبار"
    OnCurrent ="[Event Procedure]"
    OnDelete ="[Event Procedure]"
    BeforeUpdate ="[Event Procedure]"
    AfterUpdate ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
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
            Height =9137
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
                    Left =12568
                    Top =254
                    Width =1326
                    Height =315
                    ColumnOrder =9
                    Name ="NUMBER"
                    ControlSource ="NUMBER"
                    OnDblClick ="[Event Procedure]"
                    DefaultValue ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =14038
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
                    Left =10470
                    Top =225
                    Width =1221
                    Height =315
                    ColumnOrder =8
                    TabIndex =1
                    Name ="DATE_N"
                    ControlSource ="DATE_N"
                    BeforeUpdate ="[Event Procedure]"
                    DefaultValue ="=FARSIDATE(Date())"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =11730
                            Top =225
                            Width =780
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
                    ColumnCount =3
                    Left =7528
                    Top =604
                    Width =6366
                    Height =315
                    ColumnOrder =7
                    TabIndex =3
                    BackColor =16777164
                    Name ="CUST_NO"
                    ControlSource ="CUST_NO"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT     hes, NAME AS nam, hes AS Expr1\015\012FROM         CUST_HESAB"
                    ColumnWidths ="0;4536;567"
                    OnExit ="[Event Procedure]"
                    OnNotInList ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =14038
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
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5308
                    Top =604
                    Width =2166
                    Height =315
                    ColumnOrder =6
                    TabIndex =4
                    BackColor =16777164
                    Name ="CUST_NO2"
                    ControlSource ="CUST_NO"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT hes FROM CUST_HESAB"
                    OnExit ="[Event Procedure]"
                    OnNotInList ="[Event Procedure]"
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =10348
                    Top =964
                    Width =3555
                    Height =315
                    ColumnOrder =5
                    TabIndex =5
                    Name ="MOLAH"
                    ControlSource ="MOLAH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =14053
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
                    Left =111
                    Top =1360
                    Width =15135
                    Height =6825
                    TabIndex =6
                    Name ="HAVALAH_ENTER_SUB"
                    SourceObject ="Form.HAVALAH_ENTER_SUB"
                    LinkChildFields ="NUMBER;TAG"
                    LinkMasterFields ="NUMBER;htag"
                    OnExit ="[Event Procedure]"
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
                    Left =7326
                    Top =226
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
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3481
                    Top =8494
                    Width =1881
                    Height =315
                    ColumnOrder =3
                    TabIndex =8
                    Name ="MOGU"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =5461
                            Top =8494
                            Width =750
                            Height =300
                            Name ="Label63"
                            Caption ="موجودي:"
                        End
                    End
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    AccessKey =209
                    TextFontCharSet =0
                    Left =476
                    Top =151
                    Width =1365
                    Height =465
                    FontWeight =700
                    TabIndex =9
                    Name ="Command103"
                    Caption ="چاپ &رسيد انبار"
                    OnClick ="[Event Procedure]"
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
                    Left =1873
                    Top =976
                    Width =1431
                    Height =300
                    ColumnOrder =2
                    TabIndex =10
                    BackColor =32768
                    ForeColor =65535
                    Name ="N_S"
                    ControlSource ="N_S"
                    OnDblClick ="[Event Procedure]"
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =3313
                            Top =976
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
                    Left =7965
                    Top =228
                    Width =1956
                    Height =315
                    ColumnOrder =10
                    TabIndex =11
                    ForeColor =8388608
                    Name ="USER_NAME"
                    ControlSource ="USER_NAME"
                    DefaultValue ="UCurrentUser()"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =9978
                            Top =225
                            Width =450
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
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =6646
                    Top =226
                    Width =411
                    Height =345
                    ColumnOrder =1
                    TabIndex =12
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
                    Left =5293
                    Top =226
                    Width =1131
                    Height =285
                    ColumnOrder =0
                    TabIndex =2
                    Name ="FNUMCO"
                    ControlSource ="FNUMCO"
                    OnDblClick ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextFontCharSet =0
                            Left =6445
                            Top =226
                            Width =1395
                            Height =285
                            Name ="Label107"
                            Caption ="شماره برگه داخلي:"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    DecimalPlaces =0
                    SpecialEffect =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =448
                    Top =661
                    Width =2631
                    Height =300
                    TabIndex =13
                    BackColor =128
                    ForeColor =65535
                    Name ="NUMBER1"
                    ControlSource ="NUMBER1"
                    OnDblClick ="[Event Procedure]"
                    Begin
                        Begin Label
                            Visible = NotDefault
                            BackStyle =1
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =3088
                            Top =661
                            Width =1710
                            Height =300
                            BackColor =128
                            ForeColor =65535
                            Name ="LKH"
                            Caption =" شماره برگه خروج مواد:"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    DecimalPlaces =0
                    SpecialEffect =0
                    OverlapFlags =87
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =448
                    Top =976
                    Width =1431
                    Height =300
                    TabIndex =14
                    BackColor =32768
                    ForeColor =65535
                    Name ="MABNA"
                    OnDblClick ="[Event Procedure]"
                End
                Begin CheckBox
                    Visible = NotDefault
                    OverlapFlags =93
                    ReadingOrder =1
                    Left =14097
                    Top =8277
                    TabIndex =15
                    Name ="SGN1"
                    ControlSource ="SGN1"
                    Begin
                        Begin Label
                            OverlapFlags =87
                            Left =14358
                            Top =8220
                            Width =675
                            Height =240
                            ForeColor =16711680
                            Name ="LSGN1"
                            Caption ="انبار"
                        End
                    End
                End
                Begin CheckBox
                    Visible = NotDefault
                    OverlapFlags =95
                    ReadingOrder =1
                    Left =14097
                    Top =8531
                    TabIndex =16
                    Name ="SGN2"
                    ControlSource ="SGN2"
                    OnClick ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =87
                            Left =14351
                            Top =8475
                            Width =975
                            Height =240
                            ForeColor =16711680
                            Name ="LSGN2"
                            Caption ="مدير توليد"
                        End
                    End
                End
                Begin CheckBox
                    Visible = NotDefault
                    OverlapFlags =87
                    ReadingOrder =1
                    Left =14097
                    Top =8790
                    TabIndex =17
                    Name ="SGN3"
                    ControlSource ="SGN3"
                    Begin
                        Begin Label
                            OverlapFlags =87
                            Left =14351
                            Top =8727
                            Width =960
                            Height =240
                            ForeColor =16711680
                            Name ="LSGN3"
                            Caption ="كنترل كيفيت"
                        End
                    End
                End
                Begin CheckBox
                    Locked = NotDefault
                    OverlapFlags =85
                    ReadingOrder =1
                    Left =2895
                    Top =279
                    Width =280
                    Height =218
                    TabIndex =18
                    Name ="OKF"
                    ControlSource ="OKF"
                    DefaultValue ="True"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =3225
                            Top =225
                            Width =795
                            Height =240
                            ForeColor =255
                            Name ="Label146"
                            Caption ="تائيد حواله"
                        End
                    End
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    Left =4065
                    Top =150
                    Width =1155
                    Height =465
                    FontWeight =700
                    TabIndex =19
                    Name ="ESLAH"
                    Caption ="اصلاح برگه"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    Left =1845
                    Top =150
                    Width =1035
                    Height =465
                    TabIndex =20
                    Name ="Command111"
                    Caption ="چاپ 2"
                    OnClick ="[Event Procedure]"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =5313
                    Top =964
                    Width =4371
                    Height =315
                    TabIndex =21
                    ForeColor =8388608
                    Name ="DEPATMAN"
                    ControlSource ="DEPATMAN"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT DEPART.DEPATMAN, DEPART.DEPNAME FROM DEPART ORDER BY DEPART.DEPNAME;"
                    ColumnWidths ="0"
                    DefaultValue ="=[Forms]![DEFAULT]![TFSAZMAN]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =9768
                            Top =964
                            Width =495
                            Height =240
                            ForeColor =8388608
                            Name ="Label110"
                            Caption ="واحد :"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "HAVALAH_ENTER.cls"
