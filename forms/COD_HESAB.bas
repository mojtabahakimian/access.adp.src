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
    BorderStyle =3
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =9524
    DatasheetFontHeight =10
    ItemSuffix =19
    Left =5385
    Top =525
    Right =14910
    Bottom =6510
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xe070828899a9e240
    End
    RecordSource ="COD_HESAB"
    Caption ="تعريف حسابها"
    OnCurrent ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial (Arabic)"
    OnError ="[Event Procedure]"
    RecordSourceQualifier ="dbo"
    Begin
        Begin Label
            BackStyle =0
            TextFontCharSet =178
            FontName ="Tahoma"
        End
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
            Width =850
            Height =850
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
        Begin Section
            CanGrow = NotDefault
            Height =6000
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5900
                    Top =170
                    Width =2301
                    Height =315
                    ColumnWidth =1500
                    Name ="CODE"
                    ControlSource ="CODE"
                    Format ="Fixed"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextFontCharSet =0
                            Left =8259
                            Top =177
                            Width =1125
                            Height =300
                            Name ="Label1"
                            Caption ="شماره حساب:"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    DecimalPlaces =0
                    NumeralShapes =2
                    ColumnCount =2
                    Left =3502
                    Top =172
                    Width =1800
                    Height =315
                    TabIndex =1
                    Name ="BANK"
                    ControlSource ="BANK"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TCOD_BANKS.CODE, TCOD_BANKS.NAMES FROM TCOD_BANKS;"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextFontCharSet =0
                            Left =5347
                            Top =172
                            Width =495
                            Height =300
                            Name ="Label3"
                            Caption =" بانك:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =405
                    Top =165
                    Width =2400
                    Height =315
                    TabIndex =2
                    Name ="SHOBEH"
                    ControlSource ="SHOBEH"
                    Begin
                        Begin Label
                            OverlapFlags =95
                            TextFontCharSet =0
                            Left =2891
                            Top =165
                            Width =600
                            Height =285
                            Name ="Label5"
                            Caption ="شعبه"
                        End
                    End
                End
                Begin Subform
                    OverlapFlags =87
                    BorderWidth =3
                    Left =287
                    Top =1515
                    Width =9015
                    Height =3615
                    TabIndex =4
                    Name ="CHECK_BOOK"
                    SourceObject ="Form.CHECK_BOOK"
                    LinkChildFields ="COD_HES"
                    LinkMasterFields ="CODE"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextFontCharSet =0
                            Left =4790
                            Top =1215
                            Width =1425
                            Height =300
                            FontWeight =700
                            Name ="Label9"
                            Caption =" تعريف دسته چك"
                        End
                    End
                End
                Begin Rectangle
                    OverlapFlags =255
                    Left =124
                    Top =90
                    Width =9325
                    Height =874
                    Name ="Box10"
                End
                Begin TextBox
                    TabStop = NotDefault
                    DecimalPlaces =0
                    SpecialEffect =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    BackStyle =0
                    NumeralShapes =2
                    Left =793
                    Top =5159
                    Width =1984
                    Height =300
                    TabIndex =5
                    BackColor =12632256
                    ForeColor =32768
                    Name ="Text13"
                    ControlSource ="=[CHECK_BOOK].[Form]![CCODH]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =2848
                            Top =5174
                            Width =1410
                            Height =300
                            Name ="Label14"
                            Caption ="تعداد دسته چك:"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =247
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    DecimalPlaces =0
                    NumeralShapes =2
                    ColumnCount =2
                    Left =390
                    Top =570
                    Width =7800
                    Height =315
                    TabIndex =3
                    Name ="MOIN"
                    ControlSource ="MOIN"
                    RowSourceType ="Table/View/StoredProc"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextFontCharSet =0
                            Left =8235
                            Top =570
                            Width =1080
                            Height =300
                            Name ="Label18"
                            Caption =" معين معادل:"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "COD_HESAB.cls"
