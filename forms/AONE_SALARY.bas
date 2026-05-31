Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    DividingLines = NotDefault
    DefaultView =0
    AllowUpdating =4
    ScrollBars =0
    ViewsAllowed =1
    TabularCharSet =178
    TabularFamily =8
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =64
    GridY =64
    Width =11746
    DatasheetFontHeight =10
    ItemSuffix =50
    Left =1680
    Top =2235
    Right =13425
    Bottom =7890
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xb74244ead9bde240
    End
    RecordSource ="ONE_SALARY"
    Caption ="محاسبه نرخ دستمزد و جذب دستمزد به  كنترل كالاي درجريان ساخت"
    OnCurrent ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    DatasheetFontName =""
    OnLoad ="[Event Procedure]"
    OnError ="[Event Procedure]"
    AllowDatasheetView =0
    RecordSourceQualifier ="dbo"
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
        End
        Begin CommandButton
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
        Begin OptionGroup
            SpecialEffect =3
            Width =1701
            Height =1701
        End
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            OldBorderStyle =0
            Width =1701
            LabelX =-1701
            FontName ="Tahoma"
            AsianLineBreak =255
        End
        Begin ComboBox
            SpecialEffect =2
            Width =1701
            LabelX =-1701
            FontName ="Tahoma"
        End
        Begin FormHeader
            Height =806
            BackColor =8421376
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =93
                    TextFontCharSet =178
                    TextAlign =2
                    TextFontFamily =2
                    Left =2308
                    Top =26
                    Width =6705
                    Height =780
                    FontSize =23
                    FontWeight =700
                    ForeColor =16777215
                    Name ="Label26"
                    Caption ="جذب دستمزد به توليد"
                    FontName ="Shekoofeh"
                End
                Begin Label
                    OverlapFlags =215
                    TextFontCharSet =178
                    TextAlign =2
                    TextFontFamily =2
                    Left =2295
                    Width =6705
                    Height =780
                    FontSize =23
                    FontWeight =700
                    ForeColor =10040115
                    Name ="Label27"
                    Caption ="جذب دستمزد به توليد"
                    FontName ="Shekoofeh"
                End
            End
        End
        Begin Section
            Height =5412
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =6360
                    Top =135
                    Width =1755
                    Height =300
                    BackColor =16777164
                    Name ="ODATE"
                    ControlSource ="ODATE"
                    BeforeUpdate ="[Event Procedure]"
                    DefaultValue ="=farsidate(Date())"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =8310
                            Top =135
                            Width =1050
                            Height =240
                            Name ="Label1"
                            Caption ="تاريخ محاسبه:"
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
                    Left =6360
                    Top =1155
                    Width =1755
                    Height =300
                    TabIndex =3
                    BackColor =13434828
                    Name ="FRIDS"
                    ControlSource ="FRIDS"
                    Format ="#\" روز\""
                    StatusBarText ="جمعه"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =8310
                            Top =1155
                            Width =915
                            Height =240
                            Name ="Label7"
                            Caption ="تعداد جمعه:"
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
                    Left =6360
                    Top =1496
                    Width =1755
                    Height =300
                    TabIndex =4
                    BackColor =13434828
                    Name ="TATIL"
                    ControlSource ="TATIL"
                    Format ="#\" روز\""
                    StatusBarText ="تعطيلات"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =8310
                            Top =1500
                            Width =1095
                            Height =240
                            Name ="Label9"
                            Caption ="تعداد تعطيلات:"
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
                    Left =6360
                    Top =1836
                    Width =1755
                    Height =300
                    TabIndex =5
                    BackColor =13434828
                    Name ="MORA"
                    ControlSource ="MORA"
                    Format ="#\" دقيقه\""
                    StatusBarText ="مرخصي"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =8310
                            Top =1830
                            Width =1260
                            Height =240
                            Name ="Label11"
                            Caption ="مرخصي پرسنل :"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =6360
                    Top =2926
                    Width =1755
                    Height =300
                    TabIndex =6
                    BackColor =13434828
                    Name ="KOL_SALARY"
                    ControlSource ="KOL_SALARY"
                    Format ="#,##0.00\" ريال \";-#,##0.00\" ريال \""
                    StatusBarText ="كل حقوق پرداختي"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =8310
                            Top =2925
                            Width =2655
                            Height =240
                            Name ="Label13"
                            Caption ="كل حقوق پرداختي توليد  طبق ليست:"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =6360
                    Top =3266
                    Width =1755
                    Height =300
                    TabIndex =7
                    BackColor =13434828
                    Name ="KOSURAT"
                    ControlSource ="KOSURAT"
                    Format ="#,##0.00\" ريال \";-#,##0.00\" ريال \""
                    StatusBarText ="كسورات"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =8310
                            Top =3270
                            Width =2010
                            Height =240
                            Name ="Label15"
                            Caption ="جمع  دستمزد غير مستقيم:"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =6360
                    Top =3606
                    Width =1755
                    Height =300
                    TabIndex =8
                    BackColor =13434828
                    Name ="SALARYKH"
                    ControlSource ="SALARYKH"
                    Format ="#,##0.00\" ريال \";-#,##0.00\" ريال \""
                    StatusBarText ="خالص حقوق  قابل محاسبه"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =8310
                            Top =3600
                            Width =2040
                            Height =240
                            Name ="Label17"
                            Caption ="جمع دستمزد مستقيم توليد:"
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
                    Left =6360
                    Top =3946
                    Width =1755
                    Height =300
                    TabIndex =9
                    BackColor =13421619
                    Name ="JHOURVT"
                    ControlSource ="JHOURVT"
                    StatusBarText ="جمع ساعت كار واقعي توليد"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =8310
                            Top =3945
                            Width =3330
                            Height =240
                            Name ="Label19"
                            Caption ="جمع ساعت كار واقعي توليد طبق ليست حقوق:"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =6360
                    Top =4635
                    Width =1755
                    Height =300
                    TabIndex =10
                    BackColor =0
                    ForeColor =10092543
                    Name ="CUSTOH"
                    ControlSource ="CUSTOH"
                    Format ="#,##0.00\" ريال \";-#,##0.00\" ريال \""
                    StatusBarText ="نرخ يك ساعت كار"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =8310
                            Top =4639
                            Width =1320
                            Height =240
                            Name ="Label21"
                            Caption ="نرخ يك ساعت كار:"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =270
                    Top =480
                    Width =3465
                    Height =285
                    TabIndex =11
                    Name ="USER_NAME"
                    ControlSource ="USER_NAME"
                    DefaultValue ="Ucurrentuser()"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3916
                            Top =480
                            Width =615
                            Height =240
                            Name ="Label23"
                            Caption ="نام كاربر"
                        End
                    End
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =6360
                    Top =475
                    Width =1755
                    Height =285
                    TabIndex =1
                    BackColor =16777164
                    Name ="FMON"
                    ControlSource ="FMON"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT MON_ID, MON FROM MON ORDER BY MON_ID"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =8310
                            Top =480
                            Width =495
                            Height =240
                            Name ="Label3"
                            Caption ="از ماه:"
                        End
                    End
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =6360
                    Top =815
                    Width =1755
                    Height =285
                    TabIndex =2
                    BackColor =16777164
                    Name ="TMON"
                    ControlSource ="TMON"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT MON_ID, MON FROM MON ORDER BY MON_ID"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =8310
                            Top =810
                            Width =480
                            Height =240
                            Name ="Label5"
                            Caption ="تا ماه:"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    FELineBreak = NotDefault
                    DecimalPlaces =0
                    SpecialEffect =0
                    OverlapFlags =85
                    TextAlign =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =240
                    Top =105
                    Width =2781
                    Height =300
                    TabIndex =12
                    BackColor =32768
                    ForeColor =65535
                    Name ="N_S"
                    ControlSource ="N_S"
                    OnDblClick ="[Event Procedure]"
                    AsianLineBreak =0
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =85
                            Left =3030
                            Top =105
                            Width =1485
                            Height =300
                            BackColor =32768
                            ForeColor =65535
                            Name ="lsanad"
                            Caption ="ثبت درسند شماره:"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =240
                    Top =825
                    Width =4289
                    Height =523
                    TabIndex =13
                    Name ="Command28"
                    Caption ="دريافت اطلاعات حقوق از حسابداري مالي"
                    OnClick ="[Event Procedure]"
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =6360
                    Top =4999
                    Width =1755
                    Height =300
                    TabIndex =14
                    BackColor =0
                    ForeColor =10092543
                    Name ="Text29"
                    ControlSource ="=[CUSTOH]/60"
                    Format ="#,##0.00\" ريال \";-#,##0.00\" ريال \""
                    StatusBarText ="نرخ يك ساعت كار"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =8310
                            Top =5003
                            Width =1320
                            Height =240
                            Name ="Label30"
                            Caption ="نرخ هر دقيقه كار:"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =240
                    Top =1365
                    Width =4289
                    Height =523
                    TabIndex =15
                    Name ="Command31"
                    Caption ="صدور سند جذب دستمزد به توليد"
                    OnClick ="[Event Procedure]"
                End
                Begin OptionGroup
                    OverlapFlags =93
                    Left =270
                    Top =3610
                    Width =5336
                    Height =943
                    TabIndex =16
                    Name ="NORS"
                    ControlSource ="NORS"
                    DefaultValue ="-1"
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =215
                            Left =996
                            Top =3495
                            Width =3330
                            Height =240
                            BackColor =-2147483633
                            Name ="Label33"
                            Caption ="تفاوت ساعت كار  واقعي محاسبه شده و كاربرگ"
                        End
                        Begin OptionButton
                            OverlapFlags =87
                            Left =5336
                            Top =3884
                            Name ="Option35"
                            Begin
                                Begin Label
                                    OverlapFlags =87
                                    TextAlign =3
                                    Left =3900
                                    Top =3825
                                    Width =1350
                                    Height =240
                                    Name ="Label36"
                                    Caption ="در نرخ تعديل شود"
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =87
                            Left =5319
                            Top =4294
                            OptionValue =0
                            Name ="Option37"
                            Begin
                                Begin Label
                                    OverlapFlags =87
                                    TextAlign =3
                                    Left =1661
                                    Top =4212
                                    Width =3570
                                    Height =240
                                    Name ="Label38"
                                    Caption ="به حساب سربار منظور شود  مبلغ انتقالي به سربار:"
                                End
                            End
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =215
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =315
                    Top =3810
                    Width =1296
                    Height =300
                    TabIndex =17
                    BackColor =0
                    ForeColor =10092543
                    Name ="CUSTOHTDL"
                    ControlSource ="CUSTOHTDL"
                    Format ="#,##0.00\" ريال \";-#,##0.00\" ريال \""
                    StatusBarText ="نرخ يك ساعت كار تعديلي"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =1665
                            Top =3814
                            Width =1905
                            Height =240
                            Name ="Label40"
                            Caption ="نرخ تعديلي:"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =282
                    Top =2551
                    Width =4806
                    Height =510
                    TabIndex =18
                    Name ="KOSURATTEXT"
                    ControlSource ="KOSURATTEXT"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5131
                            Top =2551
                            Width =630
                            Height =240
                            Name ="Label43"
                            Caption ="شامل:"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =215
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =318
                    Top =4189
                    Width =1296
                    Height =300
                    TabIndex =19
                    BackColor =13434828
                    Name ="ENTEGHALS"
                    ControlSource ="ENTEGHALS"
                    Format ="#,##0.00\" ريال \";-#,##0.00\" ريال \""
                    StatusBarText ="مبلغ انتقالي به سربار"
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =6360
                    Top =4276
                    Width =1755
                    Height =300
                    TabIndex =20
                    BackColor =13421619
                    Name ="JHOURVTKAR"
                    ControlSource ="JHOURVTKAR"
                    StatusBarText ="جمع ساعت كار واقعي توليد"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =8310
                            Top =4275
                            Width =2835
                            Height =240
                            Name ="Label45"
                            Caption ="جمع ساعت كار واقعي توليدطبق كاربرگ:"
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
                    Left =6360
                    Top =2196
                    Width =1755
                    Height =300
                    TabIndex =21
                    BackColor =13434828
                    Name ="EZAFH"
                    ControlSource ="EZAFH"
                    Format ="#\" ساعت\""
                    StatusBarText ="مرخصي"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =8310
                            Top =2190
                            Width =1455
                            Height =240
                            Name ="Label47"
                            Caption ="ساعت اضافه كاري :"
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
                    Left =6360
                    Top =2556
                    Width =1755
                    Height =300
                    TabIndex =22
                    BackColor =13434828
                    Name ="DAYSTS"
                    ControlSource ="DAYSTS"
                    Format ="#\" روز\""
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =8310
                            Top =2550
                            Width =1740
                            Height =240
                            Name ="Label49"
                            Caption ="كل روزهاي كاركرد توليد :"
                        End
                    End
                End
            End
        End
        Begin FormFooter
            Height =0
            BackColor =-2147483633
            Name ="FormFooter"
        End
    End
End
CodeBehindForm
' See "AONE_SALARY.cls"
