Version =20
VersionRequired =20
Begin Form
    OrderByOn = NotDefault
    DefaultView =0
    AllowUpdating =4
    TabularCharSet =178
    TabularFamily =42
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =1701
    DatasheetFontHeight =10
    ItemSuffix =12
    Left =825
    Top =720
    Right =2820
    Bottom =1965
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x8b0df0d76dc2e240
    End
    DatasheetFontName ="Arial (Arabic)"
    OnLoad ="[Event Procedure]"
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
            Height =1261
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    Name ="HES_K"
                    ControlSource ="HES_K"
                End
                Begin TextBox
                    OverlapFlags =85
                    Top =340
                    TabIndex =1
                    Name ="HES_M"
                    ControlSource ="HES_M"
                End
                Begin TextBox
                    OverlapFlags =85
                    Top =680
                    TabIndex =2
                    Name ="HES_T"
                    ControlSource ="HES_T"
                End
                Begin TextBox
                    OverlapFlags =85
                    Top =1021
                    TabIndex =3
                    Name ="MOGH"
                    ControlSource ="MOGH"
                End
            End
        End
    End
End
CodeBehindForm
' See "HES_CONTROL_JARIAN.cls"
