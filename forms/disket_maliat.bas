Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
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
    Width =5295
    DatasheetFontHeight =10
    ItemSuffix =37
    Left =4980
    Top =1545
    Right =10275
    Bottom =7260
    HelpContextId =132
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x80244eb33dfae140
    End
    Caption ="ارسال اطلاعات"
    OnOpen ="[Event Procedure]"
    HelpFile ="AMIN98.HLP>MAIN"
    DatasheetFontName ="Arial (Arabic)"
    OnError ="[Event Procedure]"
    AllowDatasheetView =0
    Begin
        Begin Label
            BackStyle =0
            TextFontCharSet =178
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
        End
        Begin ComboBox
            SpecialEffect =2
            TextFontCharSet =178
            Width =1701
            LabelX =-1701
            FontName ="Tahoma"
        End
        Begin Section
            Height =5725
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin CommandButton
                    Cancel = NotDefault
                    OverlapFlags =85
                    AccessKey =199
                    TextFontCharSet =0
                    Left =1275
                    Top =5385
                    Width =1134
                    Height =340
                    FontWeight =700
                    TabIndex =3
                    Name ="Command2"
                    Caption ="&انصراف"
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"
                End
                Begin TextBox
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =1
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =315
                    Top =824
                    Width =3946
                    Height =345
                    TabIndex =1
                    Name ="PATH"
                    DefaultValue ="\"C:\\\""
                    FontName ="Tahoma"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextFontCharSet =0
                            Left =4384
                            Top =810
                            Width =660
                            Height =330
                            Name ="Label8"
                            Caption ="مسير:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin Rectangle
                    OverlapFlags =247
                    Left =220
                    Top =660
                    Width =4871
                    Height =645
                    Name ="Box11"
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =202
                    TextFontCharSet =0
                    Left =2430
                    Top =5385
                    Width =1134
                    Height =340
                    FontWeight =700
                    TabIndex =2
                    Name ="Command13"
                    Caption ="&تهيه ديسكت"
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"
                    ControlTipText ="اين گزينه را جهت انتقال اطلاعات به ديسك  تاييد نماييد"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    AllowedText =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =535
                    Top =138
                    Width =2796
                    Height =345
                    Name ="MMO"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT MON.MON_ID, MON.MON FROM MON ORDER BY MON.MON_ID;"
                    ColumnWidths ="0"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =3478
                            Top =135
                            Width =315
                            Height =285
                            Name ="Label4"
                            Caption =" ماه:"
                            FontName ="Tahoma"
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
                    Left =1860
                    Top =1439
                    Width =1411
                    Height =345
                    TabIndex =4
                    BackColor =255
                    Name ="tdf"
                    FontName ="Tahoma"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =3315
                            Top =1425
                            Width =1770
                            Height =330
                            Name ="Label22"
                            Caption ="تاريخ ثبت در دفتر روزنامه:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =285
                    Top =1923
                    Width =3816
                    Height =345
                    TabIndex =5
                    BackColor =255
                    Name ="nsa"
                    RowSourceType ="Value List"
                    RowSource ="1;\"چك شخصي\";2;\"پرداخت با كارت اعتباري\";3;\"انتقال بانكي\";4;\"سفته\";5;\"چك "
                        "تضمين شده\";6;\"واريز نقدي\";7;\"واريز خزانه\";1;\"ارائه ليست بدون پرداخت ماليات"
                        "\""
                    ColumnWidths ="0"
                    DefaultValue ="6"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =4140
                            Top =1920
                            Width =1140
                            Height =345
                            Name ="Label24"
                            Caption ="نحوه پرداخت:"
                            FontName ="Tahoma"
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
                    IMESentenceMode =3
                    Left =2490
                    Top =2388
                    Width =1161
                    Height =300
                    TabIndex =6
                    Name ="serial"
                    FontName ="Tahoma"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =3750
                            Top =2385
                            Width =1350
                            Height =330
                            Name ="Label26"
                            Caption ="شماره سريال چك:"
                            FontName ="Tahoma"
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
                    IMESentenceMode =3
                    Left =2490
                    Top =2838
                    Width =1776
                    Height =300
                    TabIndex =7
                    Name ="chdt"
                    FontName ="Tahoma"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =4365
                            Top =2835
                            Width =735
                            Height =330
                            Name ="Label28"
                            Caption ="تاريخ چك:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =840
                    Top =3228
                    Width =3816
                    Height =345
                    TabIndex =8
                    Name ="bnk"
                    RowSourceType ="Value List"
                    RowSource ="1;\"انصار\""
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =4695
                            Top =3225
                            Width =405
                            Height =240
                            Name ="Label30"
                            Caption ="بانك:"
                            FontName ="Tahoma"
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
                    IMESentenceMode =3
                    Left =855
                    Top =3633
                    Width =3366
                    Height =345
                    TabIndex =9
                    Name ="shobeh"
                    FontName ="Tahoma"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =4320
                            Top =3630
                            Width =780
                            Height =240
                            Name ="Label32"
                            Caption ="نام شعبه:"
                            FontName ="Tahoma"
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
                    IMESentenceMode =3
                    Left =855
                    Top =4056
                    Width =3186
                    Height =345
                    TabIndex =10
                    Name ="numh"
                    FontName ="Tahoma"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =4080
                            Top =4050
                            Width =1080
                            Height =315
                            Name ="Label34"
                            Caption ="شماره حساب:"
                            FontName ="Tahoma"
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
                    IMESentenceMode =3
                    Left =855
                    Top =4527
                    Width =3186
                    Height =345
                    TabIndex =11
                    BackColor =255
                    Name ="mabl"
                    FontName ="Tahoma"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =4080
                            Top =4521
                            Width =1080
                            Height =315
                            Name ="Label36"
                            Caption ="مبلغ پرداختي:"
                            FontName ="Tahoma"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "disket_maliat.cls"
