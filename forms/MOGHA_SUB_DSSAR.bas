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
    Width =3973
    DatasheetFontHeight =10
    ItemSuffix =8
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x8309670813b1e240
    End
    RecordSource ="MOGHA_SUB_DSSAR"
    Caption ="MOGHA_SUB_DSSAR"
    DatasheetFontName ="Arial (Arabic)"
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
            Height =1493
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    Left =144
                    Top =113
                    Name ="HES_K"
                    ControlSource ="HES_K"
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =144
                    Top =453
                    TabIndex =1
                    Name ="HES_M"
                    ControlSource ="HES_M"
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =144
                    Top =793
                    TabIndex =2
                    Name ="HES_T"
                    ControlSource ="HES_T"
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =144
                    Top =1133
                    TabIndex =3
                    Name ="MOGH"
                    ControlSource ="MOGH"
                End
            End
        End
    End
End
