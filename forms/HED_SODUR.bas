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
    TabularFamily =42
    Orientation =1
    PictureAlignment =2
    PictureSizeMode =1
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =11281
    DatasheetFontHeight =10
    ItemSuffix =14
    Left =240
    Right =11280
    Bottom =5745
    HelpContextId =800802
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x5ae6c2259aa9e240
    End
    RecordSource ="HED_SODUR"
    Caption =" ليست روزانه چكها جهت چاپ چك"
    OnCurrent ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    HelpFile ="dena.hlp"
    DatasheetFontName ="Arial (Arabic)"
    OnError ="[Event Procedure]"
    AllowDatasheetView =0
    RecordSourceQualifier ="dbo"
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
        Begin CustomControl
            SpecialEffect =2
            Width =4536
            Height =2835
        End
        Begin Section
            CanGrow = NotDefault
            Height =6633
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =9184
                    Top =396
                    Width =1146
                    Height =315
                    ColumnOrder =2
                    Name ="SDATE"
                    ControlSource ="DATE"
                    AfterUpdate ="[Event Procedure]"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            TextAlign =1
                            ReadingOrder =2
                            NumeralShapes =2
                            Left =10369
                            Top =396
                            Width =765
                            Height =300
                            Name ="Label1"
                            Caption ="تاريخ:"
                        End
                    End
                End
                Begin Subform
                    OverlapFlags =85
                    Left =64
                    Top =963
                    Width =11160
                    Height =5055
                    TabIndex =1
                    Name ="CHK_SODUR"
                    SourceObject ="Form.CHK_SODUR"
                    LinkChildFields ="DATE"
                    LinkMasterFields ="DATE"
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    BackStyle =0
                    NumeralShapes =2
                    Left =8945
                    Top =6108
                    Width =1020
                    Height =315
                    ColumnOrder =1
                    TabIndex =2
                    BackColor =12632256
                    ForeColor =32768
                    Name ="Text6"
                    ControlSource ="=[CHK_SODUR].[Form]![CCODH]"
                    Format ="#\" برگ\""
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            TextAlign =1
                            ReadingOrder =2
                            NumeralShapes =2
                            Left =9995
                            Top =6105
                            Width =1230
                            Height =300
                            Name ="Label7"
                            Caption ="تعدادچك صادره:"
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    BackStyle =0
                    NumeralShapes =2
                    Left =4205
                    Top =6138
                    Width =1965
                    Height =315
                    ColumnOrder =0
                    TabIndex =3
                    BackColor =12632256
                    ForeColor =32768
                    Name ="Text8"
                    ControlSource ="=[CHK_SODUR].[Form]![SMABL]"
                    Format ="#,###\"ريال\""
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            TextAlign =1
                            ReadingOrder =2
                            NumeralShapes =2
                            Left =6200
                            Top =6135
                            Width =1260
                            Height =300
                            Name ="Label9"
                            Caption ="جمع مبلغ چكها:"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =141
                    TextFontCharSet =0
                    Left =675
                    Top =375
                    Width =930
                    Height =495
                    FontWeight =700
                    TabIndex =4
                    Name ="Command10"
                    Caption ="&چاپ چك"
                    OnClick ="[Event Procedure]"
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    BackStyle =0
                    NumeralShapes =2
                    Left =855
                    Top =6138
                    Width =1965
                    Height =315
                    ColumnOrder =3
                    TabIndex =5
                    BackColor =12632256
                    ForeColor =32768
                    Name ="Text11"
                    ControlSource ="=[CHK_SODUR].[Form]![MOGUDI]"
                    Format ="#,###\"ريال\""
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            TextAlign =1
                            ReadingOrder =2
                            NumeralShapes =2
                            Left =2850
                            Top =6135
                            Width =1260
                            Height =300
                            Name ="Label12"
                            Caption ="موجودي:"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "HED_SODUR.cls"
