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
    TabularFamily =0
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =64
    GridY =64
    Width =15045
    DatasheetFontHeight =10
    ItemSuffix =27
    Left =6750
    Top =855
    Right =21795
    Bottom =9810
    DatasheetGridlinesColor =12632256
    OrderBy ="ANBGRD_HEAD.GRD_NUM DESC"
    RecSrcDt = Begin
        0x07d6856739b8e240
    End
    RecordSource ="ANBGRD_HEAD"
    Caption ="انبار گرداني"
    OnDelete ="[Event Procedure]"
    BeforeUpdate ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    DatasheetFontName =""
    OnLoad ="[Event Procedure]"
    OnError ="[Event Procedure]"
    AllowDatasheetView =0
    RecordSourceQualifier ="dbo"
    Begin
        Begin Label
            BackStyle =0
            TextFontCharSet =178
            BackColor =-2147483633
            ForeColor =-2147483630
        End
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
            Width =850
            Height =850
        End
        Begin Line
            Width =1701
        End
        Begin Image
            BackStyle =0
            OldBorderStyle =0
            PictureAlignment =2
            Width =1701
            Height =1701
        End
        Begin CommandButton
            TextFontCharSet =178
            Width =1701
            Height =283
            FontSize =8
            FontWeight =400
            FontName ="MS Sans Serif"
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
            OldBorderStyle =0
            BackStyle =0
            Width =4536
            Height =2835
            LabelX =-1701
        End
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            TextFontCharSet =178
            Width =1701
            LabelX =-1701
            BackColor =-2147483643
            ForeColor =-2147483640
            AsianLineBreak =255
        End
        Begin ListBox
            SpecialEffect =2
            TextFontCharSet =178
            Width =1701
            Height =1417
            LabelX =-1701
            BackColor =-2147483643
            ForeColor =-2147483640
        End
        Begin ComboBox
            SpecialEffect =2
            TextFontCharSet =178
            Width =1701
            LabelX =-1701
            BackColor =-2147483643
            ForeColor =-2147483640
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
        Begin ToggleButton
            TextFontCharSet =178
            Width =283
            Height =283
            FontSize =8
            FontWeight =400
            FontName ="MS Sans Serif"
        End
        Begin Tab
            TextFontCharSet =178
            BackStyle =0
            Width =5103
            Height =3402
        End
        Begin Page
            Width =1701
            Height =1701
        End
        Begin Section
            CanGrow = NotDefault
            Height =8970
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =13659
                    Top =60
                    Width =615
                    Height =300
                    ColumnWidth =615
                    BackColor =6723891
                    ForeColor =0
                    Name ="GRD_NUM"
                    ControlSource ="GRD_NUM"
                    FontName ="Tahoma"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            TextAlign =3
                            Left =14369
                            Top =60
                            Width =585
                            Height =285
                            Name ="GRD_NUM_Label"
                            Caption ="شماره:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =11902
                    Top =60
                    Width =1125
                    Height =300
                    ColumnWidth =2310
                    TabIndex =1
                    BackColor =13434828
                    Name ="GRD_DATE"
                    ControlSource ="GRD_DATE"
                    BeforeUpdate ="[Event Procedure]"
                    FontName ="Tahoma"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            TextAlign =3
                            Left =13117
                            Top =60
                            Width =450
                            Height =285
                            Name ="GRD_DATE_Label"
                            Caption ="تاريخ:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin ComboBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =7043
                    Top =60
                    Width =4365
                    Height =300
                    ColumnWidth =615
                    TabIndex =2
                    BackColor =13434828
                    Name ="GRD_ANBAR"
                    ControlSource ="GRD_ANBAR"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT CODE, NAMES FROM TCOD_ANBAR"
                    ColumnWidths ="0"
                    FontName ="Tahoma"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            TextAlign =3
                            Left =11453
                            Top =60
                            Width =390
                            Height =285
                            Name ="GRD_ANBAR_Label"
                            Caption ="انبار:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    FELineBreak = NotDefault
                    DecimalPlaces =0
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5235
                    Top =495
                    Width =846
                    Height =300
                    TabIndex =4
                    BackColor =32768
                    ForeColor =65535
                    Name ="N_S"
                    ControlSource ="N_S"
                    OnDblClick ="[Event Procedure]"
                    FontName ="Tahoma"
                    AsianLineBreak =0
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =6090
                            Top =495
                            Width =1485
                            Height =300
                            BackColor =32768
                            ForeColor =65535
                            Name ="lsanad"
                            Caption ="ثبت درسند شماره:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =7628
                    Top =520
                    Width =6531
                    Height =285
                    TabIndex =5
                    Name ="COMMENT"
                    ControlSource ="COMMENT"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="\"\""
                    FontName ="Tahoma"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =14215
                            Top =502
                            Width =750
                            Height =285
                            Name ="Label10"
                            Caption ="توضيحات:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin Tab
                    OverlapFlags =87
                    TextFontCharSet =0
                    Top =825
                    Width =15045
                    Height =8145
                    TabIndex =6
                    Name ="TabCtl11"
                    FontName ="Tahoma"
                    OnClick ="[Event Procedure]"
                    Begin
                        Begin Page
                            OverlapFlags =215
                            Left =60
                            Top =1230
                            Width =14850
                            Height =7605
                            Name ="Page12"
                            Caption ="شمارش اول"
                            Begin
                                Begin Subform
                                    Enabled = NotDefault
                                    OverlapFlags =215
                                    Left =60
                                    Top =1755
                                    Width =14790
                                    Height =6735
                                    Name ="ANBGRD_LST_SUB"
                                    SourceObject ="Form.ANBGRD_LST_SUB"
                                    LinkChildFields ="GRD_NUM"
                                    LinkMasterFields ="GRD_NUM"
                                End
                                Begin CommandButton
                                    OverlapFlags =215
                                    TextFontCharSet =0
                                    Left =12357
                                    Top =1302
                                    Width =2456
                                    Height =407
                                    FontWeight =700
                                    TabIndex =1
                                    Name ="Command19"
                                    Caption ="دريافت موجودي فعلي"
                                    OnClick ="[Event Procedure]"
                                    FontName ="Tahoma"
                                End
                                Begin CommandButton
                                    OverlapFlags =215
                                    TextFontCharSet =0
                                    Left =9030
                                    Top =1302
                                    Width =3281
                                    Height =408
                                    FontWeight =700
                                    TabIndex =2
                                    Name ="Command20"
                                    Caption ="چاپ لست كالا ها جهت شمارش اول"
                                    OnClick ="[Event Procedure]"
                                    FontName ="Tahoma"
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =215
                            Left =71
                            Top =1230
                            Width =14839
                            Height =7605
                            Name ="Page13"
                            OnClick ="[Event Procedure]"
                            Caption ="شمارش دوم"
                            Begin
                                Begin Subform
                                    Enabled = NotDefault
                                    OverlapFlags =247
                                    Left =71
                                    Top =1755
                                    Width =14775
                                    Height =6765
                                    Name ="ANBGRD_LST_SUB2"
                                    SourceObject ="Form.ANBGRD_LST_SUB2"
                                    LinkChildFields ="GRD_NUM"
                                    LinkMasterFields ="GRD_NUM"
                                    OnEnter ="[Event Procedure]"
                                End
                                Begin CommandButton
                                    OverlapFlags =215
                                    TextFontCharSet =0
                                    Left =11464
                                    Top =1274
                                    Width =3281
                                    Height =407
                                    FontWeight =700
                                    TabIndex =1
                                    Name ="Command21"
                                    Caption ="چاپ لست كالا ها جهت شمارش دوم"
                                    OnClick ="[Event Procedure]"
                                    FontName ="Tahoma"
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =215
                            Left =71
                            Top =1230
                            Width =14839
                            Height =7605
                            Name ="Page14"
                            OnClick ="[Event Procedure]"
                            Caption ="شمارش سوم"
                            Begin
                                Begin Subform
                                    Enabled = NotDefault
                                    OverlapFlags =247
                                    Left =71
                                    Top =1755
                                    Width =14775
                                    Height =6750
                                    Name ="ANBGRD_LST_SUB3"
                                    SourceObject ="Form.ANBGRD_LST_SUB3"
                                    LinkChildFields ="GRD_NUM"
                                    LinkMasterFields ="GRD_NUM"
                                    OnEnter ="[Event Procedure]"
                                End
                                Begin CommandButton
                                    OverlapFlags =215
                                    TextFontCharSet =0
                                    Left =11623
                                    Top =1231
                                    Width =3281
                                    Height =407
                                    FontWeight =700
                                    TabIndex =1
                                    Name ="Command22"
                                    Caption ="چاپ لست كالا ها جهت شمارش سوم"
                                    OnClick ="[Event Procedure]"
                                    FontName ="Tahoma"
                                End
                                Begin CommandButton
                                    OverlapFlags =215
                                    TextFontCharSet =0
                                    Left =8325
                                    Top =1232
                                    Width =3281
                                    Height =407
                                    FontWeight =700
                                    TabIndex =2
                                    Name ="Command23"
                                    Caption ="صدور سند انبار گرداني"
                                    OnClick ="[Event Procedure]"
                                    FontName ="Tahoma"
                                End
                            End
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =3390
                    Top =66
                    Width =1740
                    Height =315
                    ColumnWidth =2310
                    TabIndex =3
                    BackColor =13434828
                    Name ="GRD_HES"
                    ControlSource ="GRD_HES"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT     hes\015\012FROM         CUST_HESAB"
                    FontName ="Tahoma"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            TextAlign =3
                            Left =5175
                            Top =60
                            Width =1770
                            Height =285
                            Name ="GRD_HES_Label"
                            Caption ="حساب كسري و اضافات:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontCharSet =0
                    Left =1545
                    Top =76
                    Width =735
                    Height =405
                    TabIndex =7
                    Name ="Command24"
                    Caption ="چاپ تگ"
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontCharSet =0
                    Left =810
                    Top =76
                    Width =735
                    Height =405
                    TabIndex =8
                    Name ="Command25"
                    Caption ="چاپ "
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    Left =75
                    Top =75
                    Width =735
                    Height =405
                    FontWeight =700
                    TabIndex =9
                    ForeColor =-2147483630
                    Name ="ESLAH"
                    Caption ="اصلاح "
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    FELineBreak = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3390
                    Top =483
                    Width =1311
                    Height =315
                    TabIndex =10
                    BackColor =16777215
                    ForeColor =8388608
                    Name ="USER_NAME"
                    ControlSource ="USER_NAME"
                    DefaultValue ="UCurrentUser()"
                    FontName ="Tahoma"
                    AsianLineBreak =0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =4758
                            Top =480
                            Width =435
                            Height =240
                            BackColor =16777215
                            ForeColor =8388608
                            Name ="Label111"
                            Caption ="كاربر:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    Left =2276
                    Top =70
                    Width =1035
                    Height =405
                    FontWeight =700
                    TabIndex =11
                    ForeColor =-2147483630
                    Name ="Command26"
                    Caption ="حذف "
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"
                End
            End
        End
    End
End
CodeBehindForm
' See "ANBGRD_HEAD.cls"
