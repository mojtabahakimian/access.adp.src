Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    Modal = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    ControlBox = NotDefault
    AutoCenter = NotDefault
    AllowDeletions = NotDefault
    CloseButton = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    KeyPreview = NotDefault
    AllowEdits = NotDefault
    DefaultView =2
    AllowUpdating =4
    ScrollBars =1
    ViewsAllowed =2
    TabularCharSet =178
    TabularFamily =10
    BorderStyle =3
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =5722
    ItemSuffix =16
    Left =3870
    Top =2190
    Right =8490
    Bottom =5460
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x9056a219716fe240
    End
    RecordSource ="SELECT DETA_HES.NUMBER, DETA_HES.NAME, DETA_HES.N_KOL FROM DETA_HES WHERE (((DET"
        "A_HES.N_KOL)=[Forms]![PGET_HED]![PGET_LST_SUB].[Form]![FHES_K]) AND ((DETA_HES.C"
        "ODE_E) Like IIf([Forms]![Baseknow]![PERSON]=2,\"0\",\"*\") Or (DETA_HES.CODE_E) "
        "Is Null)) ORDER BY DETA_HES.NAME;"
    Caption ="."
    OnCurrent ="[Event Procedure]"
    DatasheetFontName ="Arabic Transparent"
    OnKeyPress ="[Event Procedure]"
    OnError ="[Event Procedure]"
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
        Begin Section
            Height =1417
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1927
                    Top =396
                    ColumnWidth =1050
                    Name ="NUMBER"
                    ControlSource ="NUMBER"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4464
                            Top =396
                            Width =1035
                            Height =240
                            Name ="Label3"
                            Caption ="شماره حساب"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1104
                    Top =793
                    Width =2490
                    ColumnWidth =3660
                    TabIndex =1
                    Name ="NAME"
                    ControlSource ="NAME"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4280
                            Top =793
                            Width =1185
                            Height =240
                            Name ="Label5"
                            Caption ="نام حساب كل"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "SERCHMOINF.cls"
