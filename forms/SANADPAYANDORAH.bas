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
    Width =5839
    DatasheetFontHeight =10
    ItemSuffix =15
    Left =3030
    Top =2925
    Right =8865
    Bottom =4950
    HelpContextId =131
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xbece4ccb101ce240
    End
    Caption ="صدور سند پايان دوره"
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
        Begin Section
            Height =2040
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =53
                    TextFontCharSet =0
                    Left =2827
                    Top =1339
                    Width =1644
                    Height =510
                    FontWeight =700
                    Name ="Command2"
                    Caption ="&صدور سند"
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"
                    UnicodeAccessKey =1589
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    BorderWidth =1
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =2
                    Left =453
                    Top =56
                    Width =4740
                    Height =855
                    FontWeight =700
                    Name ="Label9"
                    Caption ="لطفاً مطمئن شويد كه كاربر ديگري از سيستم استفاده نمي كند سپس روي گزينه صدور سند "
                        "كليك كنيد"
                    FontName ="Titr"
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =0
                    OverlapFlags =93
                    TextFontCharSet =2
                    TextFontFamily =10
                    BackStyle =0
                    Left =453
                    Top =963
                    Width =4743
                    Height =255
                    FontSize =6
                    TabIndex =1
                    ForeColor =8388608
                    Name ="Text0"
                    DefaultValue ="\"n\""
                    FontName ="Wingdings"
                End
                Begin Rectangle
                    Visible = NotDefault
                    SpecialEffect =2
                    OverlapFlags =247
                    Left =453
                    Top =993
                    Width =4743
                    Height =245
                    Name ="Box7"
                End
                Begin CommandButton
                    Cancel = NotDefault
                    OverlapFlags =85
                    AccessKey =199
                    TextFontCharSet =0
                    Left =1158
                    Top =1338
                    Width =1644
                    Height =510
                    FontWeight =700
                    TabIndex =2
                    Name ="Command14"
                    Caption ="&انصراف"
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"
                    ControlTipText ="اين گزينه را جهت دريافت  اطلاعات از ديسك  تاييد نماييد"
                End
            End
        End
    End
End
CodeBehindForm
' See "SANADPAYANDORAH.cls"
