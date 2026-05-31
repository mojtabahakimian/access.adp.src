Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    ControlBox = NotDefault
    ShortcutMenu = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    CloseButton = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    KeyPreview = NotDefault
    AllowEdits = NotDefault
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
    GridX =24
    GridY =24
    Width =7594
    DatasheetFontHeight =10
    ItemSuffix =5
    Left =10560
    Top =4335
    Right =18150
    Bottom =6345
    TimerInterval =60000
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x65c4f194da82e540
    End
    RecordSource ="SELECT * FROM [DEFAULTDEP] WHERE USERID = 88"
    Caption ="تعريف واحد در حال كار"
    DatasheetFontName ="Arial (Arabic)"
    OnTimer ="[Event Procedure]"
    OnLoad ="[Event Procedure]"
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
        Begin TextBox
            SpecialEffect =2
            OldBorderStyle =0
            TextFontCharSet =178
            FontName ="Tahoma"
        End
        Begin ComboBox
            SpecialEffect =2
            TextFontCharSet =178
            FontName ="Tahoma"
        End
        Begin Section
            Height =2031
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =1774
                    Top =120
                    Width =4207
                    Height =315
                    Name ="TFSAZMAN"
                    ControlSource ="TFSAZMAN"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT DEPART.DEPATMAN, DEPART.DEPNAME FROM DEPART ORDER BY DEPART.DEPNAME;"
                    ColumnWidths ="0"
                    ValidationRule ="Not Is Null"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =6094
                            Top =120
                            Width =1500
                            Height =300
                            Name ="Label107"
                            Caption ="واحد سيستم جاري:"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =1774
                    Top =555
                    Width =4207
                    Height =315
                    TabIndex =1
                    Name ="SHIFT"
                    ControlSource ="SHIFT"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT SHIFT.SHIFT_ID, SHIFT.SHNAME FROM SHIFT;"
                    ColumnWidths ="0"
                    ValidationRule ="Not Is Null"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =6094
                            Top =555
                            Width =1275
                            Height =300
                            Name ="Label2"
                            Caption ="شيفت جاري:"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =199
                    TextFontCharSet =0
                    Left =3897
                    Top =1559
                    Width =1290
                    Height =345
                    FontWeight =700
                    TabIndex =3
                    Name ="Command147"
                    Caption ="&انصراف"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =208
                    TextFontCharSet =0
                    Left =2622
                    Top =1559
                    Width =1275
                    Height =345
                    FontWeight =700
                    TabIndex =2
                    Name ="Command148"
                    Caption ="&ذخيره و خروج"
                    OnClick ="[Event Procedure]"
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =247
                    IMESentenceMode =3
                    Left =1158
                    Top =425
                    TabIndex =4
                    Name ="USERID"
                    ControlSource ="USERID"
                    DefaultValue ="[Forms]![BASEKNOW]![USERCOD]"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =3348
                            Top =425
                            Width =690
                            Height =240
                            Name ="Label4"
                            Caption ="USERID:"
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
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =4299
                    Top =968
                    Width =1701
                    Height =345
                    TabIndex =5
                    Name ="menup"
                    RowSourceType ="Value List"
                    RowSource ="0;\"استاندارد\";1;\"تصويري\";2;\"پنلي\""
                    ColumnWidths ="0"
                    BeforeUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =6113
                            Top =1025
                            Width =1200
                            Height =240
                            Name ="Label15"
                            Caption ="منوي پيش فرض"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "DEFAULT.cls"
