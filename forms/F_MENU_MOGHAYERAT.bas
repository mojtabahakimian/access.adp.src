Version =20
VersionRequired =20
Begin Form
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
    ItemSuffix =38
    Left =2940
    Top =1575
    Right =8895
    Bottom =3090
    HelpContextId =500203
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xbace0f5a1a4ce240
    End
    Caption ="صورت مغايرت"
    OnOpen ="[Event Procedure]"
    HelpFile ="dena.hlp"
    DatasheetFontName ="Arial (Arabic)"
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
            Height =1537
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    TextFontCharSet =0
                    Left =3135
                    Top =997
                    Width =870
                    Height =315
                    FontWeight =700
                    TabIndex =5
                    Name ="Command5"
                    Caption ="تاييد"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    Cancel = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    Left =2267
                    Top =997
                    Width =870
                    Height =302
                    FontWeight =700
                    TabIndex =6
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
                    Left =5667
                    Width =234
                    Height =255
                    TabIndex =2
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
                    Left =3118
                    Top =525
                    Width =1431
                    Height =300
                    TabIndex =3
                    Name ="DT1"
                    Format ="0000\\/00\\/00"
                    BeforeUpdate ="[Event Procedure]"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =4643
                            Top =527
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
                    Left =892
                    Top =533
                    Width =1416
                    Height =300
                    TabIndex =4
                    Name ="DT2"
                    Format ="0000\\/00\\/00"
                    BeforeUpdate ="[Event Procedure]"
                    DefaultValue ="=FARSIDATE(Date())"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =2402
                            Top =535
                            Width =630
                            Height =285
                            Name ="Label23"
                            Caption ="تا تاريخ:"
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
                    AllowedText =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3109
                    Top =135
                    Width =1431
                    Height =300
                    Name ="XTAF"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT RTRIM(CAST(TDETA_HES.N_KOL AS nvarchar)) + '-' + RTRIM(CAST(TDETA_HES.NUM"
                        "BER AS nvarchar)) + '-' + RTRIM(CAST(TDETA_HES.TNUMBER AS nvarchar)) AS hes FROM"
                        " TOTA_HES INNER JOIN DETA_HES INNER JOIN TDETA_HES ON DETA_HES.NUMBER = TDETA_HE"
                        "S.NUMBER AND DETA_HES.N_KOL = TDETA_HES.N_KOL ON TOTA_HES.NUMBER = DETA_HES.N_KO"
                        "L"
                    AfterUpdate ="[Event Procedure]"
                    OnExit ="[Event Procedure]"
                    OnNotInList ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =4628
                            Top =142
                            Width =1020
                            Height =285
                            Name ="Label31"
                            Caption ="حساب :"
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
                    AllowedText =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =180
                    Top =150
                    Width =2841
                    Height =300
                    TabIndex =1
                    Name ="HHTAF"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT RTRIM(CAST(TDETA_HES.N_KOL AS nvarchar)) + '-' + RTRIM(CAST(TDETA_HES.NUM"
                        "BER AS nvarchar)) + '-' + RTRIM(CAST(TDETA_HES.TNUMBER AS nvarchar)) AS hes, TDE"
                        "TA_HES.NAME FROM TOTA_HES INNER JOIN DETA_HES INNER JOIN TDETA_HES ON DETA_HES.N"
                        "UMBER = TDETA_HES.NUMBER AND DETA_HES.N_KOL = TDETA_HES.N_KOL ON TOTA_HES.NUMBER"
                        " = DETA_HES.N_KOL"
                    ColumnWidths ="0"
                    AfterUpdate ="[Event Procedure]"
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5667
                    Top =375
                    Width =234
                    Height =255
                    TabIndex =7
                    Name ="HTTAF"
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5667
                    Top =750
                    Width =234
                    Height =255
                    TabIndex =8
                    Name ="KOL"
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5667
                    Top =1125
                    Width =234
                    Height =255
                    TabIndex =9
                    Name ="MOIN"
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4860
                    Top =997
                    Width =459
                    Height =255
                    TabIndex =10
                    Name ="TAF"
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4181
                    Top =1053
                    Width =459
                    Height =255
                    TabIndex =11
                    Name ="MON"
                End
            End
        End
    End
End
CodeBehindForm
' See "F_MENU_MOGHAYERAT.cls"
