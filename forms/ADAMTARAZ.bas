Version =20
VersionRequired =20
Begin Form
    OrderByOn = NotDefault
    DefaultView =0
    AllowUpdating =4
    ViewsAllowed =2
    TabularCharSet =178
    TabularFamily =10
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =4536
    DatasheetFontHeight =20
    DatasheetFontWeight =700
    ItemSuffix =8
    Left =3015
    Top =450
    Right =8955
    Bottom =6045
    DatasheetForeColor =16711935
    DatasheetBackColor =13434828
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xa2d47217b4e4e540
    End
    RecordSource ="ADAMTARAZ"
    Caption ="اسنادي كه تراز نيستند"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="B Titr"
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
            Height =2394
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =709
                    Top =113
                    ColumnWidth =1695
                    Name ="N_S"
                    ControlSource ="N_S"
                    OnDblClick ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3396
                            Top =113
                            Width =885
                            Height =240
                            Name ="Label1"
                            Caption ="شماره سند"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =709
                    Top =453
                    ColumnWidth =2745
                    TabIndex =1
                    Name ="MAND"
                    ControlSource ="MAND"
                    Format ="#,###"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3771
                            Top =453
                            Width =510
                            Height =240
                            Name ="Label3"
                            Caption ="مبلغ مغايرت"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Top =1133
                    ColumnWidth =1965
                    TabIndex =2
                    Name ="DATE_S"
                    ControlSource ="DATE_S"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2697
                            Top =1133
                            Width =705
                            Height =240
                            Name ="Label4"
                            Caption ="تاريخ"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Top =1474
                    ColumnWidth =12420
                    TabIndex =3
                    Name ="SHARH_S"
                    ControlSource ="SHARH_S"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2577
                            Top =1474
                            Width =825
                            Height =240
                            Name ="Label5"
                            Caption ="شرح"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Top =1814
                    TabIndex =4
                    Name ="NO_S"
                    ControlSource ="NO_S"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2865
                            Top =1815
                            Width =660
                            Height =240
                            Name ="Label6"
                            Caption ="نوع سند"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Top =2154
                    ColumnWidth =2655
                    TabIndex =5
                    Name ="USER_NAME"
                    ControlSource ="USER_NAME"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2367
                            Top =2154
                            Width =1035
                            Height =240
                            Name ="Label7"
                            Caption ="کاربر"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "ADAMTARAZ.cls"
