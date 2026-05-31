Version =20
VersionRequired =20
Begin Form
    KeyPreview = NotDefault
    OrderByOn = NotDefault
    DefaultView =0
    AllowUpdating =4
    ViewsAllowed =2
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =4762
    RowHeight =300
    ItemSuffix =10
    Left =2640
    Top =1110
    Right =12975
    Bottom =5040
    DatasheetGridlinesColor =12632256
    OrderBy ="TOTA_HES.NUMBER"
    RecSrcDt = Begin
        0x9a6b21afb09ae240
    End
    RecordSource ="TOTA_HES"
    Caption ="سر فصل حسابهاي كل"
    OnCurrent ="[Event Procedure]"
    BeforeUpdate ="[Event Procedure]"
    AfterUpdate ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Tahoma"
    OnKeyPress ="[Event Procedure]"
    OnLoad ="[Event Procedure]"
    OnError ="[Event Procedure]"
    AllowFormView =0
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
        Begin Section
            Height =1834
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2268
                    Top =113
                    Width =1116
                    ColumnWidth =570
                    Name ="NUMBER"
                    ControlSource ="NUMBER"
                    OnDblClick ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3470
                            Top =113
                            Width =1035
                            Height =240
                            Name ="Label1"
                            Caption ="شماره حساب"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =144
                    Top =453
                    Width =3240
                    ColumnWidth =3300
                    TabIndex =1
                    Name ="NAME"
                    ControlSource ="NAME"
                    OnDblClick ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3725
                            Top =453
                            Width =780
                            Height =240
                            Name ="Label3"
                            Caption ="نام حساب"
                        End
                    End
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =933
                    Top =793
                    Width =2451
                    ColumnWidth =1245
                    TabIndex =2
                    Name ="NO_HES"
                    ControlSource ="NO_HES"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TCOD_HESKIND.CODE, TCOD_HESKIND.NAMES FROM TCOD_HESKIND GROUP BY TCOD_HES"
                        "KIND.CODE, TCOD_HESKIND.NAMES ORDER BY TCOD_HESKIND.NAMES;"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3695
                            Top =793
                            Width =810
                            Height =240
                            Name ="Label5"
                            Caption ="نوع حساب"
                        End
                    End
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =933
                    Top =1133
                    Width =2451
                    ColumnWidth =1110
                    TabIndex =3
                    Name ="M_D"
                    ControlSource ="M_D"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TCOD_HESVAZ.CODE, TCOD_HESVAZ.NAMES FROM TCOD_HESVAZ GROUP BY TCOD_HESVAZ"
                        ".CODE, TCOD_HESVAZ.NAMES ORDER BY TCOD_HESVAZ.NAMES;"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3920
                            Top =1133
                            Width =585
                            Height =240
                            Name ="Label7"
                            Caption ="وضعيت"
                        End
                    End
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =933
                    Top =1474
                    Width =2451
                    ColumnWidth =2910
                    TabIndex =4
                    Name ="GROUP"
                    ControlSource ="GROUP"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TCOD_HESGROUP.CODE, TCOD_HESGROUP.NAMES FROM TCOD_HESGROUP GROUP BY TCOD_"
                        "HESGROUP.CODE, TCOD_HESGROUP.NAMES ORDER BY TCOD_HESGROUP.NAMES;"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3620
                            Top =1474
                            Width =885
                            Height =240
                            Name ="Label9"
                            Caption ="گروه حساب"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "TOTA_HES_SHEET.cls"
