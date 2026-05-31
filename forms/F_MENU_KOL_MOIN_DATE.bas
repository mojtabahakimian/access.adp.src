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
    ItemSuffix =30
    Left =-135
    Top =465
    Right =5820
    Bottom =2835
    HelpContextId =500203
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xbace0f5a1a4ce240
    End
    Caption ="دفتر معين"
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
            Height =2381
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    TextFontCharSet =0
                    Left =3058
                    Top =1920
                    Width =870
                    Height =345
                    FontWeight =700
                    TabIndex =8
                    Name ="Command5"
                    Caption ="تاييد"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    Cancel = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    Left =2190
                    Top =1920
                    Width =870
                    Height =347
                    FontWeight =700
                    TabIndex =9
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
                    Left =5669
                    Top =906
                    Width =234
                    Height =255
                    TabIndex =4
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
                    Left =3117
                    Top =1042
                    Width =1431
                    Height =300
                    TabIndex =5
                    Name ="DT1"
                    Format ="0000\\/00\\/00"
                    BeforeUpdate ="[Event Procedure]"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =4642
                            Top =1044
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
                    Left =889
                    Top =1050
                    Width =1416
                    Height =300
                    TabIndex =6
                    Name ="DT2"
                    Format ="0000\\/00\\/00"
                    BeforeUpdate ="[Event Procedure]"
                    DefaultValue ="=Farsidate(Date())"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =2399
                            Top =1052
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
                    Left =3117
                    Top =225
                    Width =1431
                    Height =300
                    Name ="HKOL"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TOTA_HES.NUMBER FROM TOTA_HES ORDER BY TOTA_HES.NUMBER;"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =4633
                            Top =225
                            Width =870
                            Height =285
                            Name ="Label17"
                            Caption ="حساب كل:"
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
                    Left =3117
                    Top =622
                    Width =1431
                    Height =300
                    TabIndex =2
                    Name ="HMOIN"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT DETA_HES.NUMBER, DETA_HES.N_KOL FROM DETA_HES WHERE (((DETA_HES.N_KOL)=[F"
                        "orms]![F_MENU_KOL_MOIN_DATE]![HKOL] Or (DETA_HES.N_KOL)=[Forms]![F_MENU_KOL_MOIN"
                        "_DATE]![HHKOL]) AND ((DETA_HES.CODE_E) Like IIf([Forms]![Baseknow]![PERSON]=2,\""
                        "0\",\"*\") Or (DETA_HES.CODE_E) Is Null));"
                    AfterUpdate ="[Event Procedure]"
                    OnExit ="[Event Procedure]"
                    OnNotInList ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =4642
                            Top =624
                            Width =975
                            Height =285
                            Name ="Label4"
                            Caption ="حساب معين:"
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
                    Left =173
                    Top =239
                    Width =2841
                    Height =300
                    TabIndex =1
                    Name ="HHKOL"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TOTA_HES.NUMBER, TOTA_HES.NAME FROM TOTA_HES ORDER BY TOTA_HES.NAME;"
                    ColumnWidths ="0"
                    AfterUpdate ="[Event Procedure]"
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    AllowedText =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =173
                    Top =636
                    Width =2841
                    Height =300
                    TabIndex =3
                    Name ="HHMOIN"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT DETA_HES.NUMBER, DETA_HES.NAME, DETA_HES.N_KOL FROM DETA_HES WHERE DETA_H"
                        "ES.N_KOL = [Forms] ! [F_MENU_KOL_MOIN_DATE] ! [HKOL]"
                    ColumnWidths ="0"
                    AfterUpdate ="[Event Procedure]"
                    OnExit ="[Event Procedure]"
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    AllowedText =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =117
                    Top =1462
                    Width =4431
                    Height =300
                    TabIndex =7
                    Name ="PY"
                    DefaultValue ="\" \""
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =4635
                            Top =1470
                            Width =1275
                            Height =285
                            Name ="Label29"
                            Caption ="پيغام پايان گزارش:"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "F_MENU_KOL_MOIN_DATE.cls"
