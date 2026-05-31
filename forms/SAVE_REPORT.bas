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
    Width =9315
    DatasheetFontHeight =10
    ItemSuffix =23
    Left =2970
    Top =2895
    Right =12285
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
                    Left =3174
                    Top =1417
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
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =210
                    Top =239
                    Width =7951
                    Height =345
                    Name ="TITEL"
                    FontName ="Tahoma"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =8287
                            Top =225
                            Width =1020
                            Height =240
                            Name ="Label8"
                            Caption ="عنوان گزارش:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontCharSet =0
                    Left =4329
                    Top =1417
                    Width =1134
                    Height =340
                    FontWeight =700
                    TabIndex =1
                    Name ="Command13"
                    Caption ="ذخيره"
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"
                End
                Begin CheckBox
                    OverlapFlags =85
                    ReadingOrder =1
                    Left =8759
                    Top =850
                    TabIndex =3
                    Name ="C1"
                    DefaultValue ="True"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5970
                            Top =825
                            Width =2670
                            Height =240
                            Name ="Label22"
                            Caption ="توسط ساير كاربران قابل استفاده باشد"
                            FontName ="Tahoma"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "SAVE_REPORT.cls"
