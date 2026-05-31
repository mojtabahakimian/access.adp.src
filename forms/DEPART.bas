Version =20
VersionRequired =20
Begin Form
    AutoCenter = NotDefault
    OrderByOn = NotDefault
    DefaultView =2
    AllowUpdating =4
    ViewsAllowed =2
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =4785
    RowHeight =315
    ItemSuffix =9
    Left =8760
    Top =2535
    Right =14865
    Bottom =4905
    DatasheetGridlinesColor =12632256
    OrderBy ="DEPART.DEPATMAN"
    RecSrcDt = Begin
        0xdfd3d67dee9ae240
    End
    RecordSource ="DEPART"
    Caption ="تعريف واحدهاي تابعه سازمان"
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
            Height =2040
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
                    Left =907
                    Width =2631
                    Name ="DEPATMAN"
                    ControlSource ="DEPATMAN"
                    BeforeUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3669
                            Width =810
                            Height =240
                            Name ="Label1"
                            Caption ="كد واحد"
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
                    Width =3420
                    ColumnWidth =6135
                    TabIndex =1
                    Name ="DEPNAME"
                    ControlSource ="DEPNAME"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4235
                            Top =453
                            Width =270
                            Height =240
                            Name ="Label3"
                            Caption ="نام"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Top =1080
                    Width =3420
                    ColumnWidth =9765
                    TabIndex =2
                    Name ="DEPART"
                    ControlSource ="DEPART"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4095
                            Top =1080
                            Width =450
                            Height =240
                            Name ="Label5"
                            Caption ="آدرس"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Top =1440
                    Width =3420
                    TabIndex =3
                    Name ="PCODE"
                    ControlSource ="PCODE"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =4095
                            Top =1440
                            Width =675
                            Height =405
                            Name ="Label6"
                            Caption ="کد پستي"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Top =1800
                    Width =3420
                    TabIndex =4
                    Name ="BBC"
                    ControlSource ="BBC"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =4095
                            Top =1800
                            Width =690
                            Height =240
                            Name ="Label8"
                            Caption ="کد شعبه"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "DEPART.cls"
