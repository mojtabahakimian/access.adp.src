Version =20
VersionRequired =20
Begin Form
    AutoCenter = NotDefault
    AllowDeletions = NotDefault
    AllowAdditions = NotDefault
    KeyPreview = NotDefault
    AllowEdits = NotDefault
    DefaultView =2
    AllowUpdating =4
    ViewsAllowed =2
    BorderStyle =3
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =11565
    RowHeight =375
    ItemSuffix =6
    Left =1785
    Top =2625
    Right =16860
    Bottom =11355
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xdfd3d67dee9ae240
    End
    RecordSource ="DEPART"
    Caption ="تعريف واحدهاي تابعه سازمان"
    OnCurrent ="[Event Procedure]"
    DatasheetFontName ="Tahoma"
    OnKeyPress ="[Event Procedure]"
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
            Height =1644
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =7736
                    Top =113
                    Width =2631
                    ColumnWidth =1305
                    Name ="DEPATMAN"
                    ControlSource ="DEPATMAN"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =10498
                            Top =113
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
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =6947
                    Top =453
                    Width =3420
                    ColumnWidth =4695
                    TabIndex =1
                    Name ="DEPNAME"
                    ControlSource ="DEPNAME"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =11038
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
                    Left =6938
                    Top =855
                    Width =3420
                    ColumnWidth =7425
                    TabIndex =2
                    Name ="DEPART"
                    ControlSource ="DEPART"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =11033
                            Top =855
                            Width =450
                            Height =240
                            Name ="Label5"
                            Caption ="آدرس"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "searchDEPART.cls"
