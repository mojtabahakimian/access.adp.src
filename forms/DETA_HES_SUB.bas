Version =20
VersionRequired =20
Begin Form
    DefaultView =2
    AllowUpdating =4
    ViewsAllowed =2
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =10665
    RowHeight =285
    ItemSuffix =18
    Left =405
    Top =525
    Right =14715
    Bottom =2760
    DatasheetGridlinesColor =12632256
    AfterInsert ="[Event Procedure]"
    BeforeDelConfirm ="[Event Procedure]"
    RecSrcDt = Begin
        0x6d6f2d574e5ee240
    End
    RecordSource ="DETA_HES"
    Caption ="DETA_HES"
    OnCurrent ="[Event Procedure]"
    BeforeInsert ="[Event Procedure]"
    BeforeUpdate ="[Event Procedure]"
    AfterUpdate ="[Event Procedure]"
    DatasheetFontName ="Tahoma"
    OnLoad ="[Event Procedure]"
    OnError ="[Event Procedure]"
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
        Begin Section
            CanGrow = NotDefault
            Height =3195
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =93
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1893
                    Top =113
                    ColumnOrder =0
                    FontWeight =700
                    Name ="N_KOL"
                    ControlSource ="N_KOL"
                    FontName ="Titr"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextFontFamily =2
                            Left =4640
                            Top =113
                            Width =825
                            Height =240
                            FontWeight =700
                            Name ="Label1"
                            Caption ="حساب كل"
                            FontName ="Titr"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =93
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1893
                    Top =453
                    ColumnWidth =1710
                    ColumnOrder =1
                    FontWeight =700
                    TabIndex =1
                    Name ="NUMBER"
                    ControlSource ="NUMBER"
                    Format ="Fixed"
                    FontName ="Titr"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextFontFamily =2
                            Left =4775
                            Top =453
                            Width =690
                            Height =240
                            FontWeight =700
                            Name ="Label3"
                            Caption ="شماره"
                            FontName ="Titr"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1404
                    Top =793
                    Width =2190
                    ColumnWidth =5055
                    ColumnOrder =2
                    FontWeight =700
                    TabIndex =2
                    Name ="NAME"
                    ControlSource ="NAME"
                    FontName ="Titr"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextFontFamily =2
                            Left =4820
                            Top =793
                            Width =645
                            Height =240
                            FontWeight =700
                            Name ="Label5"
                            Caption ="نام حساب"
                            FontName ="Titr"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =144
                    Top =1133
                    Width =3450
                    ColumnWidth =6825
                    ColumnOrder =5
                    FontWeight =700
                    TabIndex =3
                    Name ="TOZIH"
                    ControlSource ="TOZIH"
                    FontName ="Titr"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextFontFamily =2
                            Left =4820
                            Top =1133
                            Width =645
                            Height =240
                            FontWeight =700
                            Name ="Label7"
                            Caption ="توضيحات"
                            FontName ="Titr"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =93
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1014
                    Top =1814
                    Width =2580
                    ColumnWidth =3060
                    ColumnOrder =4
                    TabIndex =4
                    Name ="ADDRESS"
                    ControlSource ="ADDRESS"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =4700
                            Top =1814
                            Width =765
                            Height =240
                            Name ="Label11"
                            Caption ="آدرس"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =93
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1884
                    Top =2154
                    Width =1710
                    ColumnOrder =3
                    TabIndex =5
                    Name ="TEL"
                    ControlSource ="TEL"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =5135
                            Top =2154
                            Width =330
                            Height =240
                            Name ="Label13"
                            Caption ="تلفن"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =93
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =566
                    Top =2551
                    TabIndex =6
                    Name ="BED_BES"
                    ControlSource ="BED_BES"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =2768
                            Top =2551
                            Width =1200
                            Height =240
                            Name ="Label16"
                            Caption ="اتوماتيك"
                        End
                    End
                End
                Begin Subform
                    OverlapFlags =247
                    Top =300
                    Width =10575
                    Height =2895
                    TabIndex =7
                    Name ="Child12"
                    SourceObject ="Form.TDETA_HES_SUB"
                    LinkChildFields ="N_KOL;NUMBER"
                    LinkMasterFields ="N_KOL;NUMBER"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextFontCharSet =0
                            Left =9240
                            Width =1425
                            Height =300
                            Name ="Label17"
                            Caption ="حسابهاي تفضيلي"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "DETA_HES_SUB.cls"
