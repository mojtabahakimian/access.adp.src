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
    Width =5272
    DatasheetFontHeight =10
    ItemSuffix =21
    Left =4995
    Top =2895
    Right =10260
    Bottom =4860
    HelpContextId =132
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x80244eb33dfae140
    End
    Caption ="ارسال اطلاعات"
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
            Height =1984
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin CommandButton
                    Cancel = NotDefault
                    OverlapFlags =85
                    AccessKey =199
                    TextFontCharSet =0
                    Left =1474
                    Top =1530
                    Width =1134
                    Height =340
                    FontWeight =700
                    TabIndex =2
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
                    KeyboardLanguage =11
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =315
                    Top =824
                    Width =3946
                    Height =345
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
                    Height =624
                    Name ="Box11"
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =202
                    TextFontCharSet =0
                    Left =2629
                    Top =1530
                    Width =1134
                    Height =340
                    FontWeight =700
                    TabIndex =1
                    Name ="Command13"
                    Caption ="&تهيه ديسكت"
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"
                    ControlTipText ="اين گزينه را جهت انتقال اطلاعات به ديسك  تاييد نماييد"
                End
            End
        End
    End
End
CodeBehindForm
' See "ERSAL_OFFLINE.cls"
