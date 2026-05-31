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
    Width =3973
    DatasheetFontHeight =10
    ItemSuffix =8
    Left =600
    Top =720
    Right =6360
    Bottom =4275
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x8f9c105b6f7de240
    End
    RecordSource ="hes_haz_tolid"
    Caption ="hes_haz_tolid"
    DatasheetFontName ="Arial (Arabic)"
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
            Height =1493
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =144
                    Top =113
                    Name ="HES_K"
                    ControlSource ="HES_K"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2936
                            Top =113
                            Width =780
                            Height =240
                            Name ="Label1"
                            Caption ="حساب كل"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =144
                    Top =453
                    TabIndex =1
                    Name ="HES_M"
                    ControlSource ="HES_M"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2771
                            Top =453
                            Width =945
                            Height =240
                            Name ="Label3"
                            Caption ="حساب معين"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =144
                    Top =793
                    TabIndex =2
                    Name ="SumOfBED"
                    ControlSource ="SumOfBED"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2876
                            Top =793
                            Width =840
                            Height =240
                            Name ="Label5"
                            Caption ="SumOfBED"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =144
                    Top =1133
                    TabIndex =3
                    Name ="SumOfBES"
                    ControlSource ="SumOfBES"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2891
                            Top =1133
                            Width =825
                            Height =240
                            Name ="Label7"
                            Caption ="SumOfBES"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "hes_haz_tolid.cls"
