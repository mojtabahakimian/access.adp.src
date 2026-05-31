Version =20
VersionRequired =20
Begin Form
    DefaultView =2
    AllowUpdating =4
    ViewsAllowed =2
    TabularCharSet =178
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =4950
    RowHeight =345
    ItemSuffix =24
    Left =1500
    Top =540
    Right =15120
    Bottom =7890
    DatasheetForeColor =33554432
    DatasheetGridlinesColor =12632256
    MaxRecords =10000000
    RecSrcDt = Begin
        0xdd62087cae6ae440
    End
    InputParameters ="@DT1 = FORMS![F_MENU_DATE]![DT1],@DT2 = FORMS![F_MENU_DATE]![DT2]"
    RecordSource ="Q_KHARED_CUSTOMER"
    Caption ="خريد فصلي"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Tahoma"
    UnknownProp = {258 ,66 ,12 ,4 ,130 } Begin
        0x400044005400310020003d00200046004f0052004d00530021005b0046005f00 ,
        0x4d0045004e0055005f0044004100540045005d0021005b004400540031005d00 ,
        0x2c00400044005400320020003d00200046004f0052004d00530021005b004600 ,
        0x5f004d0045004e0055005f0044004100540045005d0021005b00440054003200 ,
        0x5d00
    End
    AllowFormView =0
    RecordSourceQualifier ="dbo"
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
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
        Begin Section
            Height =5965
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1927
                    Top =113
                    Width =1710
                    Name ="hes"
                    ControlSource ="hes"
                    OnDblClick ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =120
                            Width =840
                            Height =240
                            Name ="Label1"
                            Caption ="كد مشتري"
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
                    Left =1927
                    Top =453
                    Width =1710
                    ColumnWidth =3195
                    TabIndex =1
                    Name ="NAME"
                    ControlSource ="NAME"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =450
                            Width =1380
                            Height =240
                            Name ="Label3"
                            Caption ="نام حساب مشتري"
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
                    Left =1927
                    Top =793
                    TabIndex =2
                    Name ="KHARED"
                    ControlSource ="KHARED"
                    OnDblClick ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =795
                            Width =735
                            Height =240
                            Name ="Label5"
                            Caption ="مبلغ خريد"
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
                    Left =1927
                    Top =1133
                    TabIndex =3
                    Name ="SumOfTAKHFIF"
                    ControlSource ="SumOfTAKHFIF"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =1133
                            Width =525
                            Height =240
                            Name ="Label7"
                            Caption ="تخفيف"
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
                    Left =1927
                    Top =1474
                    TabIndex =4
                    Name ="KHAREDKH"
                    ControlSource ="KHAREDKH"
                    OnDblClick ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =1470
                            Width =825
                            Height =240
                            Name ="Label9"
                            Caption ="خريد خالص"
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
                    Left =1927
                    Top =1814
                    TabIndex =5
                    Name ="SMBAA"
                    ControlSource ="SMBAA"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =1815
                            Width =1365
                            Height =240
                            Name ="Label11"
                            Caption ="ماليات ارزش افزوده"
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
                    Left =1927
                    Top =2154
                    TabIndex =6
                    Name ="GHABEL"
                    ControlSource ="GHABEL"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =2160
                            Width =900
                            Height =240
                            Name ="Label13"
                            Caption ="قابل پرداخت"
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
                    Left =1900
                    Top =2494
                    Width =1716
                    ColumnWidth =3285
                    TabIndex =7
                    Name ="ADDRESS"
                    ControlSource ="ADDRESS"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4072
                            Top =2494
                            Width =510
                            Height =240
                            Name ="Label14"
                            Caption ="آدرس"
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
                    Left =1900
                    Top =2834
                    Width =1716
                    TabIndex =8
                    Name ="CITY"
                    ControlSource ="CITY"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4102
                            Top =2834
                            Width =480
                            Height =240
                            Name ="Label15"
                            Caption ="شهر"
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
                    Left =1900
                    Top =3174
                    Width =1716
                    TabIndex =9
                    Name ="CODE_E"
                    ControlSource ="CODE_E"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3697
                            Top =3174
                            Width =885
                            Height =240
                            Name ="Label16"
                            Caption ="ساير"
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
                    Left =1900
                    Top =3515
                    Width =1716
                    TabIndex =10
                    Name ="ECODE"
                    ControlSource ="ECODE"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =3937
                            Top =3520
                            Width =720
                            Height =405
                            Name ="Label17"
                            Caption ="كد اقتصادي"
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
                    Left =1900
                    Top =3855
                    Width =1716
                    TabIndex =11
                    Name ="IYALAT"
                    ControlSource ="IYALAT"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =3922
                            Top =3855
                            Width =660
                            Height =240
                            Name ="Label18"
                            Caption ="استان"
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
                    Left =1900
                    Top =4195
                    Width =1716
                    TabIndex =12
                    Name ="MCODEM"
                    ControlSource ="MCODEM"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3787
                            Top =4195
                            Width =795
                            Height =240
                            Name ="Label19"
                            Caption ="كد ملي"
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
                    Left =1900
                    Top =4535
                    Width =1716
                    TabIndex =13
                    Name ="MOBILE"
                    ControlSource ="MOBILE"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3892
                            Top =4535
                            Width =690
                            Height =240
                            Name ="Label20"
                            Caption ="موبايل"
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
                    Left =1900
                    Top =4875
                    Width =1716
                    TabIndex =14
                    Name ="PCODE"
                    ControlSource ="PCODE"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3937
                            Top =4875
                            Width =645
                            Height =240
                            Name ="Label21"
                            Caption ="كد پستي"
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
                    Left =1900
                    Top =5215
                    Width =1716
                    ColumnWidth =3000
                    TabIndex =15
                    Name ="TEL"
                    ControlSource ="TEL"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4132
                            Top =5215
                            Width =450
                            Height =240
                            Name ="Label22"
                            Caption ="تلفن"
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
                    Left =1900
                    Top =5555
                    Width =1716
                    ColumnWidth =2625
                    TabIndex =16
                    Name ="TOZIH"
                    ControlSource ="TOZIH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3982
                            Top =5555
                            Width =600
                            Height =240
                            Name ="Label23"
                            Caption ="توضيحات"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "Q_KHARID_CUSTOMER.cls"
