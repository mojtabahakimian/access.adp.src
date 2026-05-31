Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    DividingLines = NotDefault
    KeyPreview = NotDefault
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
    Width =14220
    DatasheetFontHeight =10
    ItemSuffix =111
    Left =4710
    Top =2010
    Right =18585
    Bottom =7890
    HelpContextId =6005
    DatasheetGridlinesColor =12632256
    MaxRecords =100000
    RecSrcDt = Begin
        0x3d19ce0052e2e540
    End
    RecordSource ="HEAD_LST_KBKazad"
    Caption ="فاكتور برگشت  از خريد آزاد"
    OnCurrent ="[Event Procedure]"
    BeforeUpdate ="[Event Procedure]"
    AfterUpdate ="[Event Procedure]"
    HelpFile ="dena.hlp"
    DatasheetFontName ="Arial (Arabic)"
    OnKeyPress ="[Event Procedure]"
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
            Height =9807
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin Tab
                    OverlapFlags =85
                    TextFontCharSet =0
                    Style =1
                    Width =14220
                    Height =9765
                    FontWeight =700
                    Name ="TabCtl56"
                    Begin
                        Begin Page
                            OverlapFlags =215
                            AccessKey =209
                            Left =120
                            Top =450
                            Width =13965
                            Height =9180
                            Name ="Page57"
                            Caption ="&روي فاكتور"
                            Begin
                                Begin TextBox
                                    Locked = NotDefault
                                    TabStop = NotDefault
                                    OverlapFlags =223
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =1
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =10921
                                    Top =637
                                    Width =1326
                                    Height =315
                                    Name ="NUMBER1"
                                    ControlSource ="NUMBER1"
                                    DefaultValue ="0"
                                    Tag ="0"
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            TextFontCharSet =0
                                            Left =12391
                                            Top =638
                                            Width =1635
                                            Height =300
                                            Name ="Label1"
                                            Caption ="شماره فاكتور برگشت:"
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
                                    Left =8154
                                    Top =637
                                    Width =1296
                                    Height =315
                                    TabIndex =1
                                    Name ="NUMBER"
                                    ControlSource ="NUMBER"
                                    RowSourceType ="Table/View/StoredProc"
                                    RowSource ="SELECT NUMBER FROM HEAD_LST WHERE (TAG = 26) AND (NOT (NUMBER IN (SELECT NUMBER "
                                        "FROM HEAD_LST WHERE (TAG = 27)))) ORDER BY NUMBER"
                                    BeforeUpdate ="[Event Procedure]"
                                    AfterUpdate ="[Event Procedure]"
                                    Tag ="1"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            Left =9536
                                            Top =644
                                            Width =1485
                                            Height =300
                                            Name ="Label104"
                                            Caption ="شماره حواله انبار:"
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
                                    Left =5697
                                    Top =652
                                    Width =1221
                                    Height =315
                                    TabIndex =2
                                    Name ="DATE_N"
                                    ControlSource ="DATE_N"
                                    BeforeUpdate ="[Event Procedure]"
                                    AfterUpdate ="[Event Procedure]"
                                    DefaultValue ="=FARSIDATE(Date())"
                                    InputMask ="####/##/##"
                                    Tag ="1425/06/19"
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            TextFontCharSet =0
                                            Left =7062
                                            Top =652
                                            Width =1065
                                            Height =300
                                            Name ="Label9"
                                            Caption ="تاريخ فاكتور:"
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
                                    ReadingOrder =1
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    ColumnCount =2
                                    Left =10992
                                    Top =1020
                                    Width =1716
                                    Height =315
                                    TabIndex =3
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
                                            Left =12867
                                            Top =1020
                                            Width =750
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
                                    ReadingOrder =1
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    ColumnCount =2
                                    Left =8562
                                    Top =1020
                                    Width =1761
                                    Height =315
                                    TabIndex =4
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
                                            Left =10362
                                            Top =1020
                                            Width =690
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
                                    ReadingOrder =1
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    ColumnCount =2
                                    Left =5850
                                    Top =1020
                                    Width =1581
                                    Height =315
                                    TabIndex =5
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
                                            Left =7470
                                            Top =1020
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
                                    TabStop = NotDefault
                                    OldBorderStyle =0
                                    OverlapFlags =215
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =1
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    ColumnCount =2
                                    Left =7888
                                    Top =1384
                                    Width =4821
                                    Height =315
                                    TabIndex =6
                                    BackColor =13434828
                                    Name ="CUST_NO"
                                    ControlSource ="CUST_NO"
                                    RowSourceType ="Table/View/StoredProc"
                                    RowSource ="SELECT hes, NAME AS nam, hes AS Expr1 FROM CUST_HESAB"
                                    ColumnWidths ="0;2835;567"
                                    OnDblClick ="[Event Procedure]"
                                    Tag ="0"
                                    OnNotInList ="[Event Procedure]"
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            TextFontCharSet =0
                                            Left =12853
                                            Top =1384
                                            Width =780
                                            Height =300
                                            Name ="Label19"
                                            Caption =" فروشنده:"
                                        End
                                    End
                                End
                                Begin ComboBox
                                    LimitToList = NotDefault
                                    Locked = NotDefault
                                    TabStop = NotDefault
                                    OldBorderStyle =0
                                    OverlapFlags =215
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =1
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =5830
                                    Top =1365
                                    Width =2001
                                    Height =315
                                    TabIndex =7
                                    BackColor =13434828
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
                                    Left =5848
                                    Top =1744
                                    Width =6870
                                    Height =315
                                    TabIndex =8
                                    Name ="MOLAH"
                                    ControlSource ="MOLAH"
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            TextFontCharSet =0
                                            Left =12868
                                            Top =1744
                                            Width =795
                                            Height =300
                                            Name ="Label21"
                                            Caption ="ملاحظات:"
                                        End
                                    End
                                End
                                Begin Subform
                                    Locked = NotDefault
                                    OverlapFlags =247
                                    Left =120
                                    Top =2144
                                    Width =13860
                                    Height =6075
                                    TabIndex =9
                                    Name ="INVO_LST_KH_BACK_azad_SUB"
                                    SourceObject ="Form.INVO_LST_KH_BACK_azad_SUB"
                                    LinkChildFields ="NUMBER;tag"
                                    LinkMasterFields ="NUMBER;dtag"
                                    OnExit ="[Event Procedure]"
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
                                    Left =4365
                                    Top =585
                                    Width =291
                                    Height =315
                                    TabIndex =10
                                    Name ="hTAG"
                                    ControlSource ="htag"
                                    DefaultValue ="27"
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    TabStop = NotDefault
                                    DecimalPlaces =0
                                    OverlapFlags =223
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =1
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =648
                                    Top =8325
                                    Width =2106
                                    Height =315
                                    TabIndex =11
                                    Name ="Text59"
                                    ControlSource ="=[INVO_LST_KH_BACK_azad_SUB].[Form]![SMABLK]"
                                    Format ="#,##0.00\" ريال \";-#,##0.00\" ريال \""
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            TextFontCharSet =0
                                            Left =2853
                                            Top =8325
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
                                    OverlapFlags =215
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =1
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =6637
                                    Top =8351
                                    Width =1881
                                    Height =315
                                    TabIndex =12
                                    Name ="MOGU"
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            TextFontCharSet =0
                                            Left =8617
                                            Top =8351
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
                                    OverlapFlags =223
                                    TextFontCharSet =0
                                    TextAlign =2
                                    ReadingOrder =1
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =1500
                                    Top =1737
                                    Width =1401
                                    Height =300
                                    TabIndex =13
                                    BackColor =8421376
                                    ForeColor =65535
                                    Name ="N_S"
                                    ControlSource ="N_S"
                                    OnDblClick ="[Event Procedure]"
                                    Begin
                                        Begin Label
                                            BackStyle =1
                                            OverlapFlags =215
                                            TextFontCharSet =0
                                            Left =2910
                                            Top =1737
                                            Width =1485
                                            Height =300
                                            BackColor =8421376
                                            ForeColor =65535
                                            Name ="lsanad"
                                            Caption ="ثبت در سند شماره:"
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
                                    Left =180
                                    Top =1422
                                    Width =2721
                                    Height =300
                                    TabIndex =14
                                    BackColor =8421376
                                    ForeColor =65535
                                    Name ="MANDAH"
                                    OnDblClick ="[Event Procedure]"
                                    Begin
                                        Begin Label
                                            BackStyle =1
                                            OverlapFlags =215
                                            TextFontCharSet =0
                                            Left =2910
                                            Top =1422
                                            Width =1485
                                            Height =300
                                            BackColor =8421376
                                            ForeColor =65535
                                            Name ="Label69"
                                            Caption ="مانده حساب:"
                                        End
                                    End
                                End
                                Begin CommandButton
                                    Enabled = NotDefault
                                    TabStop = NotDefault
                                    OverlapFlags =215
                                    AccessKey =141
                                    TextFontCharSet =0
                                    Left =226
                                    Top =793
                                    Width =1065
                                    Height =465
                                    FontWeight =700
                                    TabIndex =15
                                    Name ="Command100"
                                    Caption ="&چاپ فاكتور"
                                    OnClick ="[Event Procedure]"
                                End
                                Begin TextBox
                                    TabStop = NotDefault
                                    DecimalPlaces =0
                                    OverlapFlags =215
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =1
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =4162
                                    Top =8336
                                    Width =1146
                                    Height =315
                                    TabIndex =16
                                    Name ="TEDADM"
                                    ControlSource ="=[INVO_LST_KH_BACK_azad_SUB].[Form]![MEGHMAR]"
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            TextFontCharSet =0
                                            Left =5407
                                            Top =8336
                                            Width =1200
                                            Height =300
                                            Name ="Label106"
                                            Caption ="جمع مرجوعات:"
                                        End
                                    End
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
                                    Left =4756
                                    Top =585
                                    Width =411
                                    Height =345
                                    TabIndex =17
                                    Name ="LETSANAD"
                                    DefaultValue ="False"
                                End
                                Begin TextBox
                                    Visible = NotDefault
                                    TabStop = NotDefault
                                    DecimalPlaces =0
                                    OverlapFlags =223
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =2
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =315
                                    Top =8340
                                    Width =291
                                    Height =285
                                    TabIndex =18
                                    Name ="USER_NAME"
                                    ControlSource ="USER_NAME"
                                    DefaultValue ="UCURRENTUSER()"
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    TabStop = NotDefault
                                    DecimalPlaces =0
                                    SpecialEffect =0
                                    OverlapFlags =247
                                    TextFontCharSet =0
                                    TextAlign =2
                                    ReadingOrder =1
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =180
                                    Top =1737
                                    Width =1356
                                    Height =300
                                    TabIndex =19
                                    BackColor =8421376
                                    ForeColor =65535
                                    Name ="MABNA"
                                End
                                Begin CheckBox
                                    Locked = NotDefault
                                    OverlapFlags =215
                                    ReadingOrder =1
                                    Left =120
                                    Top =504
                                    Width =340
                                    Height =218
                                    TabIndex =20
                                    Name ="OKF"
                                    ControlSource ="OKF"
                                    DefaultValue ="False"
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            TextFontCharSet =0
                                            Left =510
                                            Top =450
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
                                    Left =4562
                                    Top =1074
                                    Width =1170
                                    Height =465
                                    FontWeight =700
                                    TabIndex =21
                                    Name ="ESLAH"
                                    Caption ="اصلاح فاكتور"
                                    OnClick ="[Event Procedure]"
                                End
                                Begin CommandButton
                                    TabStop = NotDefault
                                    OverlapFlags =223
                                    AccessKey =205
                                    TextFontCharSet =0
                                    Left =1290
                                    Top =795
                                    Width =1455
                                    Height =465
                                    FontWeight =700
                                    TabIndex =22
                                    Name ="Command106"
                                    Caption ="چاپ &حواله انبار"
                                    OnClick ="[Event Procedure]"
                                End
                                Begin TextBox
                                    Visible = NotDefault
                                    OverlapFlags =215
                                    IMESentenceMode =3
                                    Left =2324
                                    Top =963
                                    ColumnWidth =1620
                                    TabIndex =23
                                    Name ="FNUMCO"
                                    ControlSource ="FNUMCO"
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            Left =4046
                                            Top =963
                                            Width =1680
                                            Height =240
                                            Name ="Label110"
                                            Caption ="شماره فاكتور فروشنده:"
                                        End
                                    End
                                End
                                Begin ComboBox
                                    LimitToList = NotDefault
                                    Visible = NotDefault
                                    RowSourceTypeInt =1
                                    OverlapFlags =215
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =2
                                    KeyboardLanguage =43
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    IMESentenceMode =3
                                    ColumnCount =2
                                    Left =8220
                                    Top =8900
                                    Width =3660
                                    Height =330
                                    TabIndex =24
                                    ConditionalFormat = Begin
                                        0x0100000072000000030000000000000002000000000000000200000001000000 ,
                                        0x00000000ccffcc00000000000200000003000000050000000100000000000000 ,
                                        0xffff9900000000000200000006000000080000000100000000000000ccffff00 ,
                                        0x310000000000320000000000330000000000
                                    End
                                    Name ="PERSONEL"
                                    RowSourceType ="Value List"
                                    RowSource ="1;\"حكيميان\";2;\"آقاي فلاح\";4;\"چم حيدر\";12;\"ملكي\";125;\"محمودي\";127;\"حكي"
                                        "ميان -2\";131;\"انكوباتور\";137;\"تكنيك ها و مثالها\";140;\"كارآموز\";141;\"آبيا"
                                        "ر\";142;\"جعفري\""
                                    ColumnWidths ="0"
                                    AfterUpdate ="[Event Procedure]"
                                    DefaultValue ="[Forms]![baseknow]![USERCOD]"
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            TextFontCharSet =0
                                            Left =11940
                                            Top =8900
                                            Width =720
                                            Height =240
                                            Name ="Label5"
                                            Caption ="ارجاع به :"
                                        End
                                    End
                                End
                                Begin CheckBox
                                    OverlapFlags =215
                                    ReadingOrder =1
                                    Left =12812
                                    Top =8504
                                    TabIndex =25
                                    Name ="SGN1"
                                    ControlSource ="SGN1"
                                    OnClick ="[Event Procedure]"
                                    Begin
                                        Begin Label
                                            OverlapFlags =223
                                            Left =13253
                                            Top =8447
                                            Width =630
                                            Height =240
                                            Name ="Label149"
                                            Caption ="بازرگاني"
                                        End
                                    End
                                End
                                Begin CheckBox
                                    OverlapFlags =215
                                    ReadingOrder =1
                                    Left =12815
                                    Top =8981
                                    TabIndex =26
                                    Name ="SGN3"
                                    ControlSource ="SGN3"
                                    OnClick ="[Event Procedure]"
                                    Begin
                                        Begin Label
                                            OverlapFlags =223
                                            Left =13249
                                            Top =8925
                                            Width =795
                                            Height =240
                                            Name ="Label150"
                                            Caption ="حسابداري"
                                        End
                                    End
                                End
                                Begin CheckBox
                                    OverlapFlags =215
                                    ReadingOrder =1
                                    Left =12815
                                    Top =9240
                                    TabIndex =27
                                    Name ="SGN4"
                                    ControlSource ="SGN4"
                                    OnClick ="[Event Procedure]"
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            Left =13249
                                            Top =9240
                                            Width =780
                                            Height =240
                                            Name ="Label65"
                                            Caption ="مدير عامل"
                                        End
                                    End
                                End
                                Begin CheckBox
                                    OverlapFlags =215
                                    ReadingOrder =1
                                    Left =12812
                                    Top =8741
                                    TabIndex =28
                                    Name ="SGN2"
                                    ControlSource ="SGN2"
                                    OnClick ="[Event Procedure]"
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            Left =13253
                                            Top =8684
                                            Width =495
                                            Height =240
                                            Name ="Label77"
                                            Caption ="انبار"
                                        End
                                    End
                                End
                                Begin TextBox
                                    Visible = NotDefault
                                    OverlapFlags =247
                                    TextFontCharSet =0
                                    IMESentenceMode =3
                                    Left =120
                                    Top =8299
                                    TabIndex =29
                                    Name ="UID"
                                    ControlSource ="UID"
                                    DefaultValue ="=UCurrentUserIDD()"
                                End
                            End
                        End
                        Begin Page
                            Enabled = NotDefault
                            OverlapFlags =247
                            AccessKey =129
                            Left =112
                            Top =450
                            Width =13973
                            Height =9180
                            Name ="Page58"
                            Caption ="&پشت فاكتور"
                            Begin
                                Begin TextBox
                                    Visible = NotDefault
                                    DecimalPlaces =0
                                    OverlapFlags =255
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =2
                                    KeyboardLanguage =43
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =10805
                                    Top =1080
                                    Height =315
                                    Name ="M_NAGHD"
                                    ControlSource ="M_NAGHD"
                                    Format ="#,##0.00\" ريال \";-#,##0.00\" ريال \""
                                    OnKeyPress ="[Event Procedure]"
                                    Begin
                                        Begin Label
                                            Visible = NotDefault
                                            OverlapFlags =255
                                            TextFontCharSet =0
                                            Left =12558
                                            Top =1080
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
                                    KeyboardLanguage =43
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =10805
                                    Top =1420
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
                                            Left =12558
                                            Top =1425
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
                                    Left =4291
                                    Top =1425
                                    Width =5211
                                    Height =315
                                    TabIndex =2
                                    Name ="MOIN_VAR"
                                    ControlSource ="MOIN_VAR"
                                    RowSourceType ="Table/View/StoredProc"
                                    RowSource ="SELECT hes, NAME, hes AS Expr1 FROM CUST_HESAB ORDER BY NAME, hes"
                                    ColumnWidths ="851"
                                    OnExit ="[Event Procedure]"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            TextFontCharSet =0
                                            Left =9649
                                            Top =1427
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
                                    KeyboardLanguage =43
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =10805
                                    Top =1770
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
                                            Left =12558
                                            Top =1770
                                            Width =930
                                            Height =300
                                            Name ="Label29"
                                            Caption ="مبلغ حواله:"
                                        End
                                    End
                                End
                                Begin Rectangle
                                    OverlapFlags =255
                                    Left =112
                                    Top =450
                                    Width =13725
                                    Height =371
                                    Name ="Box80"
                                End
                                Begin Rectangle
                                    OverlapFlags =255
                                    Left =4162
                                    Top =960
                                    Width =9675
                                    Height =2321
                                    Name ="Box81"
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
                                    Left =4291
                                    Top =1772
                                    Width =5211
                                    Height =315
                                    TabIndex =4
                                    Name ="MOIN_HAV"
                                    ControlSource ="MOIN_HAV"
                                    RowSourceType ="Table/View/StoredProc"
                                    RowSource ="SELECT hes, NAME, hes AS Expr1 FROM CUST_HESAB ORDER BY NAME, hes"
                                    ColumnWidths ="851"
                                    OnExit ="[Event Procedure]"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            Left =9649
                                            Top =1772
                                            Width =975
                                            Height =300
                                            Name ="Label31"
                                            Caption ="معين حواله:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    DecimalPlaces =0
                                    OverlapFlags =247
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =2
                                    KeyboardLanguage =43
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =10805
                                    Top =2115
                                    Height =315
                                    TabIndex =5
                                    Name ="TAKHFIF"
                                    ControlSource ="TAKHFIF"
                                    Format ="#,##0.00\" ريال \";-#,##0.00\" ريال \""
                                    AfterUpdate ="[Event Procedure]"
                                    OnKeyPress ="[Event Procedure]"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            Left =12558
                                            Top =2130
                                            Width =990
                                            Height =300
                                            Name ="Label37"
                                            Caption ="مبلغ تخفيف:"
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
                                    Left =4291
                                    Top =2102
                                    Width =5211
                                    Height =315
                                    TabIndex =6
                                    Name ="Text101"
                                    Format ="Fixed"
                                    AfterUpdate ="[Event Procedure]"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            Left =9643
                                            Top =2117
                                            Width =1020
                                            Height =300
                                            Name ="Label102"
                                            Caption ="درصد تخفيف:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    DecimalPlaces =0
                                    OverlapFlags =247
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =2
                                    KeyboardLanguage =43
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =10805
                                    Top =2460
                                    Height =315
                                    TabIndex =7
                                    Name ="MABL_HAZ"
                                    ControlSource ="MABL_HAZ"
                                    Format ="#,##0.00\" ريال \";-#,##0.00\" ريال \""
                                    AfterUpdate ="[Event Procedure]"
                                    OnKeyPress ="[Event Procedure]"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            Left =12558
                                            Top =2460
                                            Width =615
                                            Height =300
                                            Name ="Label33"
                                            Caption ="خدمات:"
                                        End
                                    End
                                End
                                Begin Rectangle
                                    OverlapFlags =255
                                    Left =113
                                    Top =907
                                    Width =3912
                                    Height =2154
                                    Name ="Box95"
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
                                    ColumnCount =3
                                    Left =4291
                                    Top =2462
                                    Width =5211
                                    Height =315
                                    TabIndex =8
                                    Name ="MOIN_HAZ"
                                    ControlSource ="MOIN_HAZ"
                                    RowSourceType ="Table/View/StoredProc"
                                    RowSource ="SELECT hes, NAME, hes AS Expr1 FROM CUST_HESAB ORDER BY NAME, hes"
                                    ColumnWidths ="0;3969"
                                    OnExit ="[Event Procedure]"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            Left =9643
                                            Top =2462
                                            Width =1050
                                            Height =300
                                            Name ="Label35"
                                            Caption ="معين خدمات:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    DecimalPlaces =0
                                    OverlapFlags =247
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =2
                                    KeyboardLanguage =43
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =10806
                                    Top =2850
                                    Height =315
                                    TabIndex =9
                                    Name ="MBAA"
                                    ControlSource ="MBAA"
                                    Format ="#,##0.00\" ريال \";-#,##0.00\" ريال \""
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            Left =12570
                                            Top =2850
                                            Width =1200
                                            Height =330
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
                                    Left =4291
                                    Top =2836
                                    Width =5211
                                    Height =315
                                    TabIndex =10
                                    Name ="HMBAA"
                                    ControlSource ="HMBAA"
                                    RowSourceType ="Table/View/StoredProc"
                                    RowSource ="SELECT hes, NAME, hes AS Expr1 FROM CUST_HESAB ORDER BY NAME, hes"
                                    ColumnWidths ="853"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            Left =9639
                                            Top =2835
                                            Width =1110
                                            Height =300
                                            Name ="Label138"
                                            Caption ="معين ماليات:"
                                        End
                                    End
                                End
                                Begin Subform
                                    OverlapFlags =247
                                    Left =120
                                    Top =3608
                                    Width =13770
                                    Height =3480
                                    TabIndex =11
                                    Name ="PAY_GETP_BACK_SUB"
                                    SourceObject ="Form.PAY_GETP_BACK_SUB"
                                    LinkChildFields ="NUMBER;TAG"
                                    LinkMasterFields ="NUMBER;dtag"
                                    OnExit ="[Event Procedure]"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            Left =12629
                                            Top =3288
                                            Width =1260
                                            Height =300
                                            Name ="PAY_GETD_SUB Label"
                                            Caption ="چكهاي  دريافتي"
                                            EventProcPrefix ="PAY_GETD_SUB_Label"
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
                                    Left =226
                                    Top =7200
                                    Width =2106
                                    Height =315
                                    TabIndex =12
                                    Name ="NCHK"
                                    ControlSource ="=IIf(IsNull([PAY_GETP_BACK_SUB].[Form]![SCMAB]),0,[PAY_GETP_BACK_SUB].[Form]![SC"
                                        "MAB])"
                                    Format ="#,##0.00\" ريال \";-#,##0.00\" ريال \""
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            Left =2425
                                            Top =7200
                                            Width =1260
                                            Height =300
                                            Name ="Label90"
                                            Caption ="جمع مبلغ چكها:"
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
                                    Left =9247
                                    Top =501
                                    Width =1221
                                    Height =300
                                    TabIndex =13
                                    Name ="Text73"
                                    ControlSource ="DATE_N"
                                    DefaultValue ="=FARSIDATE(Date())"
                                    InputMask ="####/##/##"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            TextFontCharSet =0
                                            Left =10477
                                            Top =495
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
                                    Left =232
                                    Top =501
                                    Width =8196
                                    Height =300
                                    TabIndex =14
                                    Name ="Combo75"
                                    ControlSource ="CUST_NO"
                                    RowSourceType ="Table/View/StoredProc"
                                    RowSource ="SELECT hes, NAME AS nam, hes AS Expr1 FROM CUST_HESAB"
                                    ColumnWidths ="0;2835;567"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            TextFontCharSet =0
                                            Left =8437
                                            Top =495
                                            Width =750
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
                                    Left =11571
                                    Top =501
                                    Width =1116
                                    Height =300
                                    TabIndex =15
                                    Name ="Text78"
                                    ControlSource ="NUMBER"
                                    DefaultValue ="0"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            TextFontCharSet =0
                                            Left =12696
                                            Top =499
                                            Width =1110
                                            Height =300
                                            Name ="Label79"
                                            Caption ="شماره فاكتور:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    TabStop = NotDefault
                                    DecimalPlaces =0
                                    SpecialEffect =0
                                    OverlapFlags =247
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =1
                                    KeyboardLanguage =43
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =220
                                    Top =2310
                                    Width =2106
                                    Height =300
                                    TabIndex =16
                                    Name ="NPAR"
                                    ControlSource ="=[M_NAGHD]+[MABL_VAR]+[MABL_HAV]+[NCHK]"
                                    Format ="#,##0.00;-#,##0.00"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            Left =2426
                                            Top =2317
                                            Width =1545
                                            Height =300
                                            Name ="Label86"
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
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =1
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =221
                                    Top =990
                                    Width =2106
                                    Height =300
                                    TabIndex =17
                                    Name ="JF"
                                    ControlSource ="=[INVO_LST_KH_BACK_azad_SUB].[Form]![SMABLK]"
                                    Format ="#,##0.00;-#,##0.00"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            Left =2431
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
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =1
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =220
                                    Top =1320
                                    Width =2106
                                    Height =300
                                    TabIndex =18
                                    Name ="HKH"
                                    ControlSource ="=[MABL_HAZ]"
                                    Format ="#,##0.00;-#,##0.00"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            Left =2425
                                            Top =1320
                                            Width =600
                                            Height =300
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
                                    OverlapFlags =247
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =1
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =225
                                    Top =1650
                                    Width =2106
                                    Height =300
                                    TabIndex =19
                                    Name ="NTKHFIF"
                                    ControlSource ="=[TAKHFIF]"
                                    Format ="#,##0.00;-#,##0.00"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            Left =2430
                                            Top =1650
                                            Width =735
                                            Height =300
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
                                    ReadingOrder =1
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =225
                                    Top =1980
                                    Width =2106
                                    Height =300
                                    TabIndex =20
                                    BackColor =16777164
                                    Name ="GHABEL"
                                    ControlSource ="=[JF]+[HKH]+[mbaa]-[NTKHFIF]"
                                    Format ="#,##0.00;-#,##0.00"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            Left =2430
                                            Top =1987
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
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =1
                                    KeyboardLanguage =43
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =220
                                    Top =2640
                                    Width =2106
                                    Height =300
                                    TabIndex =21
                                    BackColor =16777164
                                    Name ="MAN"
                                    ControlSource ="=[GHABEL]-[NPAR]"
                                    Format ="#,##0.00;#,##0.00-"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            Left =2426
                                            Top =2647
                                            Width =525
                                            Height =300
                                            Name ="Label99"
                                            Caption ="مانده:"
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
' See "HEAD_LST_KH_BACK_azad.cls"
