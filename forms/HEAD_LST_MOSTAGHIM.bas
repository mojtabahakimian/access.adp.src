Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    DividingLines = NotDefault
    KeyPreview = NotDefault
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
    Width =14626
    DatasheetFontHeight =10
    ItemSuffix =112
    Left =240
    Top =1560
    Right =14625
    Bottom =8040
    HelpContextId =6002
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x81b8d3a57386e440
    End
    RecordSource ="SELECT HEAD_LST.* FROM HEAD_LST WHERE (TAG = 15) ORDER BY NUMBER"
    Caption ="رسيد مستقيم"
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
            Height =8340
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
                    Left =12138
                    Top =133
                    Width =951
                    Height =315
                    BackColor =13434828
                    ForeColor =8388608
                    Name ="NUMBER"
                    ControlSource ="NUMBER"
                    DefaultValue ="0"
                    Tag ="1"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =13233
                            Top =138
                            Width =1110
                            Height =300
                            Name ="Label1"
                            Caption ="شماره فاكتور:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =87
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =9723
                    Top =134
                    Width =1296
                    Height =315
                    TabIndex =1
                    BackColor =16777164
                    ForeColor =8388608
                    Name ="DATE_N"
                    ControlSource ="DATE_N"
                    BeforeUpdate ="[Event Procedure]"
                    DefaultValue ="=FARSIDATE(Date())"
                    InputMask ="####/##/##"
                    Tag ="840524"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextFontCharSet =0
                            Left =11028
                            Top =134
                            Width =1065
                            Height =300
                            Name ="Label9"
                            Caption ="تاريخ فاكتور:"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =6528
                    Top =119
                    Width =1356
                    Height =315
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
                            Left =7905
                            Top =119
                            Width =1755
                            Height =300
                            Name ="Label103"
                            Caption ="شماره فاكتور فروشنده:"
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
                    Left =9918
                    Top =480
                    Width =3366
                    Height =315
                    TabIndex =3
                    BackColor =16777164
                    ForeColor =8388608
                    Name ="CUST_NO"
                    ControlSource ="CUST_NO"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT hes, NAME AS nam, hes AS Expr1 FROM CUST_HESAB"
                    ColumnWidths ="0;2835;567"
                    Tag ="111-1-1"
                    OnNotInList ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =13428
                            Top =480
                            Width =780
                            Height =300
                            Name ="Label19"
                            Caption =" فروشنده:"
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
                    Left =7998
                    Top =480
                    Width =1881
                    Height =315
                    TabIndex =4
                    BackColor =16777164
                    ForeColor =8388608
                    Name ="CUST_NO2"
                    ControlSource ="CUST_NO"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT hes FROM CUST_HESAB"
                    AfterUpdate ="[Event Procedure]"
                    OnNotInList ="[Event Procedure]"
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =8673
                    Top =840
                    Width =4620
                    Height =315
                    TabIndex =5
                    BackColor =16777164
                    ForeColor =8388608
                    Name ="MOLAH"
                    ControlSource ="MOLAH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =13443
                            Top =840
                            Width =795
                            Height =300
                            Name ="Label21"
                            Caption ="ملاحظات:"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =7473
                    Top =855
                    Width =666
                    Height =315
                    TabIndex =6
                    BackColor =16777164
                    ForeColor =8388608
                    Name ="MAS"
                    ControlSource ="MAS"
                    DefaultValue ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =8193
                            Top =855
                            Width =450
                            Height =300
                            Name ="Label106"
                            Caption ="مدت:"
                        End
                    End
                End
                Begin Subform
                    OverlapFlags =85
                    Left =123
                    Top =1635
                    Width =14280
                    Height =6180
                    TabIndex =9
                    Name ="INVO_LST_MO_SUB"
                    SourceObject ="Form.INVO_LST_MO_SUB"
                    LinkChildFields ="NUMBER;TAG"
                    LinkMasterFields ="NUMBER;htag"
                    OnExit ="[Event Procedure]"
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
                    Left =3288
                    Top =7937
                    Width =2106
                    Height =315
                    TabIndex =10
                    BackColor =13434828
                    Name ="Text59"
                    ControlSource ="=[INVO_LST_MO_SUB].[Form]![SMABLK]"
                    Format ="#,##0.00\" ريال \";-#,##0.00\" ريال \""
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =5493
                            Top =7937
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
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =113
                    Top =1017
                    Width =2721
                    Height =300
                    TabIndex =11
                    BackColor =8421376
                    ForeColor =65535
                    Name ="N_S"
                    ControlSource ="N_S"
                    OnDblClick ="[Event Procedure]"
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =2843
                            Top =1017
                            Width =1470
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
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =113
                    Top =702
                    Width =2721
                    Height =300
                    TabIndex =12
                    BackColor =8421376
                    ForeColor =65535
                    Name ="MANDAH"
                    OnDblClick ="[Event Procedure]"
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =2843
                            Top =702
                            Width =1470
                            Height =300
                            BackColor =8421376
                            ForeColor =65535
                            Name ="Label69"
                            Caption ="مانده حساب:"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =6858
                    Top =577
                    Width =531
                    TabIndex =13
                    Name ="USER_NAME"
                    ControlSource ="USER_NAME"
                    DefaultValue ="UCurrentUser()"
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =93
                    AccessKey =209
                    TextFontCharSet =0
                    Left =166
                    Top =170
                    Width =1290
                    Height =465
                    FontWeight =700
                    TabIndex =14
                    Name ="Command106"
                    Caption ="چاپ &رسيد انبار"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =93
                    AccessKey =209
                    TextFontCharSet =0
                    Left =1456
                    Top =170
                    Width =1485
                    Height =465
                    FontWeight =700
                    TabIndex =15
                    Name ="Command108"
                    Caption =" &رسيد انبار2"
                    OnClick ="[Event Procedure]"
                End
                Begin TextBox
                    Visible = NotDefault
                    TabStop = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =7425
                    Top =520
                    Width =306
                    Height =210
                    TabIndex =16
                    Name ="hTAG"
                    ControlSource ="TAG"
                    DefaultValue ="15"
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
                    Left =9738
                    Top =1200
                    Width =3366
                    Height =315
                    TabIndex =7
                    BackColor =16777164
                    ForeColor =8388608
                    Name ="MOIN_HAZ"
                    ControlSource ="MOIN_HAZ"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT hes, NAME AS nam, hes AS Expr1 FROM CUST_HESAB"
                    ColumnWidths ="0;2835;567"
                    OnNotInList ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =13248
                            Top =1200
                            Width =1020
                            Height =300
                            Name ="Label110"
                            Caption =" مركز هزينه:"
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
                    Left =7818
                    Top =1200
                    Width =1881
                    Height =315
                    TabIndex =8
                    BackColor =16777164
                    ForeColor =8388608
                    Name ="MOIN_HAZ1"
                    ControlSource ="MOIN_HAZ"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT hes FROM CUST_HESAB"
                    OnNotInList ="[Event Procedure]"
                End
                Begin TextBox
                    Visible = NotDefault
                    TabStop = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =6517
                    Top =1133
                    Width =411
                    Height =345
                    TabIndex =17
                    Name ="LETSANAD"
                    DefaultValue ="False"
                End
                Begin TextBox
                    Locked = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =12425
                    Top =8025
                    Width =1311
                    Height =315
                    TabIndex =18
                    ForeColor =8388608
                    Name ="Text112"
                    ControlSource ="USER_NAME"
                    DefaultValue ="UCurrentUser()"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =13793
                            Top =8022
                            Width =525
                            Height =300
                            ForeColor =8388608
                            Name ="Label111"
                            Caption ="كاربر:"
                        End
                    End
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    Left =2970
                    Top =165
                    Width =1155
                    Height =465
                    FontWeight =700
                    TabIndex =19
                    Name ="ESLAH"
                    Caption ="اصلاح فاكتور"
                    OnClick ="[Event Procedure]"
                End
                Begin CheckBox
                    Locked = NotDefault
                    OverlapFlags =85
                    ReadingOrder =1
                    Left =4422
                    Top =847
                    Width =340
                    Height =218
                    TabIndex =20
                    Name ="OKF"
                    ControlSource ="OKF"
                    DefaultValue ="False"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =4812
                            Top =793
                            Width =975
                            Height =240
                            ForeColor =255
                            Name ="Label146"
                            Caption ="تائيد فاكتور"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =215
                    TextFontCharSet =0
                    IMESentenceMode =3
                    TabIndex =21
                    Name ="UID"
                    ControlSource ="UID"
                    DefaultValue ="=UCurrentUserIDD()"
                End
            End
        End
    End
End
CodeBehindForm
' See "HEAD_LST_MOSTAGHIM.cls"
