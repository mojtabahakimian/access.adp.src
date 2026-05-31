Version =20
VersionRequired =20
Begin Form
    KeyPreview = NotDefault
    DefaultView =2
    AllowUpdating =4
    ViewsAllowed =2
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =7294
    RowHeight =255
    ItemSuffix =3
    Left =600
    Top =720
    Right =8295
    Bottom =4470
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xfe119a0c0c8de240
    End
    RecordSource ="SHARH"
    Caption ="شرح آماده"
    OnCurrent ="[Event Procedure]"
    DatasheetFontName ="Tahoma"
    OnKeyPress ="[Event Procedure]"
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
            Height =473
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =201
                    Top =113
                    Width =4965
                    ColumnWidth =6195
                    ColumnOrder =1
                    Name ="SHARH"
                    ControlSource ="SHARH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6452
                            Top =113
                            Width =585
                            Height =240
                            Name ="Label1"
                            Caption ="شرح"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =247
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2437
                    ColumnWidth =600
                    ColumnOrder =0
                    TabIndex =1
                    Name ="ID"
                    ControlSource ="ID"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5539
                            Width =300
                            Height =240
                            Name ="Label2"
                            Caption ="شماره"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "SHARH_LIST.cls"
