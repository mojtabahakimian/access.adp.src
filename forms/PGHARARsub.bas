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
    Width =3973
    RowHeight =270
    ItemSuffix =10
    Left =480
    Top =4485
    Right =8820
    Bottom =5745
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x9456dffa508ce240
    End
    RecordSource ="PGHARAR"
    Caption ="PGHARAR"
    DatasheetFontName ="Tahoma"
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
            Height =1834
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =144
                    Top =113
                    ColumnWidth =660
                    Name ="CODE"
                    ControlSource ="CODE"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2861
                            Top =113
                            Width =855
                            Height =240
                            Name ="Label1"
                            Caption ="كد پرسنلي"
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
                    ColumnWidth =1185
                    TabIndex =1
                    Name ="GHNUM"
                    ControlSource ="GHNUM"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2696
                            Top =453
                            Width =1020
                            Height =240
                            Name ="Label3"
                            Caption ="شماره قرارداد"
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
                    Top =793
                    ColumnWidth =1605
                    TabIndex =2
                    Name ="GHSTART"
                    ControlSource ="GHSTART"
                    BeforeUpdate ="[Event Procedure]"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2366
                            Top =793
                            Width =1350
                            Height =240
                            Name ="Label5"
                            Caption ="تاريخ شروع قرارداد"
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
                    Top =1133
                    ColumnWidth =1605
                    TabIndex =3
                    Name ="GHEND"
                    ControlSource ="GHEND"
                    BeforeUpdate ="[Event Procedure]"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2456
                            Top =1133
                            Width =1260
                            Height =240
                            Name ="Label7"
                            Caption ="تاريخ پايان قرار داد"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =144
                    Top =1474
                    ColumnWidth =1155
                    TabIndex =4
                    Name ="MORAKH"
                    ControlSource ="MORAKH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2681
                            Top =1474
                            Width =1035
                            Height =240
                            Name ="Label9"
                            Caption ="مانده مرخصي"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "PGHARARsub.cls"
