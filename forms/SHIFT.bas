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
    Width =3982
    RowHeight =300
    ItemSuffix =4
    Left =3285
    Top =570
    Right =8835
    Bottom =2760
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x1089df23ef9ae240
    End
    RecordSource ="SHIFT"
    Caption ="تعريف شيفت"
    BeforeUpdate ="[Event Procedure]"
    AfterUpdate ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Tahoma"
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
            Height =813
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =168
                    Top =113
                    Width =2586
                    Name ="SHIFT_ID"
                    ControlSource ="SHIFT_ID"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2975
                            Top =113
                            Width =750
                            Height =240
                            Name ="Label1"
                            Caption ="كد شيفت"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =159
                    Top =453
                    Width =2595
                    ColumnWidth =2685
                    TabIndex =1
                    Name ="SHNAME"
                    ControlSource ="SHNAME"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3185
                            Top =453
                            Width =540
                            Height =240
                            Name ="Label3"
                            Caption ="شيفت"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "SHIFT.cls"
