Version =20
VersionRequired =20
Begin Form
    DefaultView =0
    AllowUpdating =4
    TabularCharSet =178
    TabularFamily =42
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =5669
    DatasheetFontHeight =10
    ItemSuffix =4
    Left =600
    Top =720
    Right =8055
    Bottom =4470
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xd6d02e96d195e240
    End
    RecordSource ="AMALKARD_SUB"
    DatasheetFontName ="Arial (Arabic)"
    OnError ="[Event Procedure]"
    Begin
        Begin Label
            BackStyle =0
            TextFontCharSet =178
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
        Begin Section
            Height =1701
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =93
                    Left =1076
                    Top =623
                    Name ="MAND"
                    ControlSource ="MAND"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =3908
                            Top =623
                            Width =570
                            Height =240
                            Name ="Label0"
                            Caption ="MAND:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =1076
                    Top =963
                    TabIndex =1
                    Name ="HES_K"
                    ControlSource ="HES_K"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3638
                            Top =963
                            Width =840
                            Height =240
                            Name ="Label1"
                            Caption ="حساب كل:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =1076
                    Top =1303
                    TabIndex =2
                    Name ="HES_M"
                    ControlSource ="HES_M"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3473
                            Top =1303
                            Width =1005
                            Height =240
                            Name ="Label2"
                            Caption ="حساب معين:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =247
                    Left =736
                    Top =453
                    TabIndex =3
                    Name ="HES_T"
                    ControlSource ="HES_T"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =3043
                            Top =453
                            Width =1095
                            Height =240
                            Name ="Label3"
                            Caption ="معين تفضيلي:"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "AMALKARD_SUB.cls"
