Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    ControlBox = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    DefaultView =0
    AllowUpdating =4
    ScrollBars =0
    ViewsAllowed =1
    TabularCharSet =178
    TabularFamily =42
    BorderStyle =3
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =5952
    DatasheetFontHeight =10
    ItemSuffix =44
    Left =4650
    Top =2730
    Right =10605
    Bottom =5205
    DatasheetGridlinesColor =12632256
    MaxRecords =100000
    RecSrcDt = Begin
        0xbace0f5a1a4ce240
    End
    Caption ="گزارش"
    OnOpen ="[Event Procedure]"
    HelpFile ="dena.hlp"
    DatasheetFontName ="Arial (Arabic)"
    OnLoad ="[Event Procedure]"
    OnError ="[Event Procedure]"
    AllowDatasheetView =0
    Begin
        Begin Label
            BackStyle =0
            TextFontCharSet =178
            FontName ="Tahoma"
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
        Begin OptionGroup
            SpecialEffect =3
            Width =1701
            Height =1701
        End
        Begin TextBox
            SpecialEffect =2
            OldBorderStyle =0
            TextFontCharSet =178
            Width =1701
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
        Begin UnboundObjectFrame
            SpecialEffect =2
            OldBorderStyle =1
            Width =4536
            Height =2835
        End
        Begin Section
            Height =2494
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    TextFontCharSet =0
                    Left =2838
                    Top =2055
                    Width =870
                    Height =375
                    FontWeight =700
                    TabIndex =7
                    Name ="Command5"
                    Caption ="تاييد"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    Cancel = NotDefault
                    OverlapFlags =93
                    TextFontCharSet =0
                    Left =1970
                    Top =2055
                    Width =870
                    Height =377
                    FontWeight =700
                    TabIndex =8
                    Name ="Command6"
                    Caption ="انصراف"
                    OnClick ="[Event Procedure]"
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =113
                    Top =148
                    Width =234
                    Height =255
                    Name ="DTT"
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    AllowedText =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3082
                    Top =120
                    Width =1431
                    Height =300
                    TabIndex =1
                    Name ="DT1"
                    Format ="0000\\/00\\/00"
                    BeforeUpdate ="[Event Procedure]"
                    OnDblClick ="[Event Procedure]"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =4607
                            Top =122
                            Width =645
                            Height =285
                            Name ="Label21"
                            Caption ="از تاريخ:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    AllowedText =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =510
                    Top =127
                    Width =1416
                    Height =300
                    TabIndex =2
                    Name ="DT2"
                    Format ="0000\\/00\\/00"
                    BeforeUpdate ="[Event Procedure]"
                    OnDblClick ="[Event Procedure]"
                    DefaultValue ="=farsidate(Date())"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =2020
                            Top =129
                            Width =630
                            Height =285
                            Name ="Label23"
                            Caption ="تا تاريخ:"
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
                    Left =113
                    Top =523
                    Width =234
                    Height =255
                    TabIndex =9
                    Name ="MMOIN"
                End
                Begin OptionGroup
                    OverlapFlags =93
                    Left =285
                    Top =990
                    Width =5507
                    Height =943
                    TabIndex =10
                    Name ="VI"
                    DefaultValue ="1"
                    Begin
                        Begin OptionButton
                            OverlapFlags =87
                            ReadingOrder =1
                            Left =5475
                            Top =1228
                            OptionValue =1
                            Name ="Option32"
                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    TextFontCharSet =0
                                    TextAlign =3
                                    Left =2775
                                    Top =1200
                                    Width =2730
                                    Height =285
                                    Name ="Label33"
                                    Caption ="حسابهاي بدون مانده نما يش داده شود"
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =87
                            ReadingOrder =1
                            Left =5475
                            Top =1558
                            OptionValue =2
                            Name ="Option34"
                            Begin
                                Begin Label
                                    OverlapFlags =119
                                    TextFontCharSet =0
                                    TextAlign =3
                                    Left =2430
                                    Top =1530
                                    Width =3045
                                    Height =285
                                    Name ="Label35"
                                    Caption ="فقط حسابهاي داراي مانده نمايش داده شود"
                                End
                            End
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
                    IMESentenceMode =3
                    Left =345
                    Top =1140
                    Width =981
                    Height =285
                    TabIndex =5
                    Name ="HHHS"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT NUMBER FROM TOTA_HES ORDER BY NUMBER"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextFontCharSet =0
                            TextAlign =1
                            Left =1379
                            Top =1140
                            Width =885
                            Height =240
                            Name ="Label39"
                            Caption ="حساب كل:"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =215
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =345
                    Top =1545
                    Width =981
                    Height =285
                    TabIndex =6
                    Name ="HHHM"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextFontCharSet =0
                            TextAlign =1
                            Left =1380
                            Top =1545
                            Width =1005
                            Height =240
                            Name ="Label41"
                            Caption ="حساب معين:"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =345
                    Top =1950
                    Width =981
                    Height =285
                    TabIndex =11
                    Name ="HHHT"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextFontCharSet =0
                            TextAlign =1
                            Left =1380
                            Top =1950
                            Width =1140
                            Height =240
                            Name ="Label43"
                            Caption ="حساب تفصيلي"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    ReadingOrder =1
                    AllowedText =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =3080
                    Top =527
                    Width =1431
                    Height =300
                    TabIndex =3
                    Name ="SNDNUM1"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =4595
                            Top =527
                            Width =945
                            Height =285
                            Name ="Label17"
                            Caption ="از سند شماره:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    ReadingOrder =1
                    AllowedText =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =480
                    Top =517
                    Width =1431
                    Height =300
                    TabIndex =4
                    Name ="SNDNUM2"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =2010
                            Top =517
                            Width =960
                            Height =285
                            Name ="Label4"
                            Caption ="تا سند شماره:"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "FMENU_TARAZ_4.cls"
