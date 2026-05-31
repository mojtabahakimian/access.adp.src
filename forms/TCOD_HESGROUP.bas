Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    DividingLines = NotDefault
    OrderByOn = NotDefault
    DefaultView =2
    AllowUpdating =4
    ScrollBars =0
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =5062
    RowHeight =315
    ItemSuffix =4
    Left =3300
    Top =570
    Right =8220
    Bottom =4605
    HelpContextId =45
    DatasheetGridlinesColor =12632256
    OrderBy ="TCOD_HESGROUP.CODE"
    RecSrcDt = Begin
        0xa6a73f94ca6ae240
    End
    RecordSource ="TCOD_HESGROUP"
    Caption ="تعريف  گروه حسابها"
    OnOpen ="[Event Procedure]"
    HelpFile ="dena.hlp"
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
        Begin Section
            Height =850
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
                    Left =2718
                    Top =113
                    Width =1416
                    Height =300
                    ColumnWidth =765
                    Name ="CODE"
                    ControlSource ="CODE"
                    OnDblClick ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =4205
                            Top =113
                            Width =330
                            Height =285
                            Name ="Label1"
                            Caption ="كد"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =624
                    Top =453
                    Width =3510
                    Height =300
                    ColumnWidth =3915
                    TabIndex =1
                    Name ="NAMES"
                    ControlSource ="NAMES"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =4215
                            Top =450
                            Width =600
                            Height =285
                            Name ="Label3"
                            Caption ="نام گروه"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "TCOD_HESGROUP.cls"
