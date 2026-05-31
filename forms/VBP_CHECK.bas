Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
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
    Width =15023
    DatasheetFontHeight =10
    ItemSuffix =64
    Left =810
    Top =285
    Right =15585
    Bottom =9000
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xf7065f484e9fe240
    End
    Caption ="به حساب گذاشتن سريالي چكها"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial (Arabic)"
    OnKeyPress ="[Event Procedure]"
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
        Begin Subform
            SpecialEffect =2
            Width =1701
            Height =1701
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
            Height =8984
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin OptionGroup
                    OverlapFlags =93
                    Left =112
                    Top =45
                    Width =14907
                    Height =591
                    Name ="Frame13"
                    DefaultValue ="1"
                End
                Begin Subform
                    OverlapFlags =85
                    Left =114
                    Top =737
                    Width =14820
                    Height =7200
                    TabIndex =5
                    Name ="VBP_CHECK_SUB"
                    SourceObject ="Form.VBP_CHECK_SUB"
                    LinkChildFields ="RADIF"
                    LinkMasterFields ="LIST_NO"
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1314
                    Top =8051
                    Width =2691
                    Height =315
                    TabIndex =8
                    BackColor =13434828
                    Name ="JJKOL"
                    ControlSource ="=IIf(IsError([VBP_CHECK_SUB].[Form]![SMAB]),0,[VBP_CHECK_SUB].[Form]![SMAB])"
                    Format ="#,##0.00\" ريال \";-#,##0.00\" ريال \""
                    Tag ="=IIf(IsError([PAY_GETD_mon].[Form]![SMAB]);0;[PAY_GETD_mon].[Form]![SMAB])"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =4104
                            Top =8051
                            Width =1260
                            Height =300
                            Name ="Label61"
                            Caption ="جمع كل :"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    AutoExpand = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =215
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =12765
                    Top =165
                    Width =1328
                    Height =345
                    TabIndex =3
                    Name ="N_SERI"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT N_SERI, N_S, N_KOL, N_KOL2 FROM PAY_GETD WHERE (N_S IS NULL) AND (N_KOL I"
                        "S NULL) AND (N_KOL2 IS NULL)"
                    ValidationRule ="Is Not Null"
                    BeforeUpdate ="[Event Procedure]"
                    AfterUpdate ="[Event Procedure]"
                    OnExit ="[Event Procedure]"
                    Format ="00"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextFontCharSet =0
                            TextAlign =2
                            Left =14198
                            Top =165
                            Width =675
                            Height =345
                            FontWeight =700
                            Name ="Label36"
                            Caption ="سريال :"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =215
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =9075
                    Top =165
                    Width =1118
                    Height =315
                    TabIndex =6
                    Name ="DTS"
                    ValidationRule ="Is Not Null"
                    DefaultValue ="=FARSIDATE(Date())"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =3
                            Left =10253
                            Top =165
                            Width =450
                            Height =240
                            Name ="Label53"
                            Caption ="تاريخ:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =215
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1455
                    Top =165
                    Width =803
                    Height =315
                    TabIndex =2
                    Name ="LIST_NO"
                    ValidationRule ="Is Not Null"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =3
                            Left =2332
                            Top =165
                            Width =1110
                            Height =240
                            Name ="Label57"
                            Caption ="شماره ليست :"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =215
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =3
                    Left =3559
                    Top =170
                    Width =4536
                    Height =315
                    TabIndex =1
                    Name ="HES"
                    RowSourceType ="Table/View/StoredProc"
                    ColumnWidths ="0;2835"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =3
                            Left =8121
                            Top =170
                            Width =780
                            Height =240
                            Name ="Label4"
                            Caption ="به حساب:"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =215
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =10830
                    Top =165
                    Width =1193
                    Height =345
                    TabIndex =4
                    Name ="RADIF"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT RADIF FROM PAY_GETD WHERE (N_S IS NULL) AND (N_KOL IS NULL) AND (N_KOL2 I"
                        "S NULL)"
                    ValidationRule ="Is Not Null"
                    Format ="00"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =3
                            Left =12053
                            Top =165
                            Width =630
                            Height =240
                            Name ="Label51"
                            Caption ="شماره :"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =283
                    Top =113
                    Width =1020
                    Height =397
                    TabIndex =7
                    Name ="Command58"
                    Caption ="چاپ"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =12750
                    Top =8060
                    Width =1200
                    Height =397
                    TabIndex =9
                    Name ="Command59"
                    Caption ="چاپ سه برگي"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =11589
                    Top =8060
                    Width =1140
                    Height =397
                    TabIndex =10
                    Name ="Command60"
                    Caption ="فايل تجارت"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =10440
                    Top =8060
                    Width =1140
                    Height =397
                    TabIndex =11
                    Name ="Command61"
                    Caption ="فايل مسكن"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =9285
                    Top =8060
                    Width =1140
                    Height =397
                    TabIndex =12
                    Name ="Command62"
                    Caption ="فايل ملي"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =8115
                    Top =8060
                    Width =1140
                    Height =397
                    TabIndex =13
                    Name ="Command63"
                    Caption ="فايل ملت"
                    OnClick ="[Event Procedure]"
                End
            End
        End
    End
End
CodeBehindForm
' See "VBP_CHECK.cls"
