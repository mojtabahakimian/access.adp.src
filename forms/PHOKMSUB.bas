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
    Width =4005
    ItemSuffix =22
    Left =480
    Top =4485
    Right =8820
    Bottom =5745
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xa9115c11518ce240
    End
    RecordSource ="PHOKM"
    Caption ="PHOKM"
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
            Height =3874
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =144
                    Top =113
                    ColumnWidth =510
                    ColumnOrder =0
                    Name ="HONUM"
                    ControlSource ="HONUM"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2831
                            Top =113
                            Width =885
                            Height =240
                            Name ="Label1"
                            Caption ="شماره "
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =144
                    Top =453
                    ColumnOrder =1
                    TabIndex =1
                    Name ="CODE"
                    ControlSource ="CODE"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3071
                            Top =453
                            Width =645
                            Height =240
                            Name ="Label3"
                            Caption ="كد معين"
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
                    ColumnWidth =1110
                    ColumnOrder =2
                    TabIndex =2
                    Name ="HDATE"
                    ControlSource ="HDATE"
                    BeforeUpdate ="[Event Procedure]"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2966
                            Top =793
                            Width =750
                            Height =240
                            Name ="Label5"
                            Caption ="تاريخ حكم"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =144
                    Top =1133
                    ColumnWidth =1035
                    ColumnOrder =4
                    TabIndex =3
                    Name ="SALARY_DAYLY"
                    ControlSource ="SALARY_DAYLY"
                    Format ="#,###"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2550
                            Top =1140
                            Width =1440
                            Height =240
                            Name ="Label7"
                            Caption ="حقوق روزانه اسمي"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =144
                    Top =1474
                    ColumnWidth =1335
                    ColumnOrder =3
                    TabIndex =4
                    Name ="SALARY_DAYLYB"
                    ControlSource ="SALARY_DAYLYB"
                    Format ="#,###"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2610
                            Top =1470
                            Width =1395
                            Height =285
                            Name ="Label9"
                            Caption ="حقوق روزانه رسمي"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =144
                    Top =1814
                    ColumnWidth =1665
                    TabIndex =5
                    Name ="HOME"
                    ControlSource ="HOME"
                    Format ="#,###"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2261
                            Top =1814
                            Width =1455
                            Height =240
                            Name ="Label11"
                            Caption ="حق خواربار و مسكن"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =144
                    Top =2154
                    ColumnWidth =1155
                    TabIndex =6
                    Name ="CHILDREN"
                    ControlSource ="CHILDREN"
                    Format ="#,###"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3041
                            Top =2154
                            Width =675
                            Height =240
                            Name ="Label13"
                            Caption ="حق اولاد"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =144
                    Top =2494
                    ColumnWidth =1365
                    TabIndex =7
                    Name ="CONDITIONS"
                    ControlSource ="CONDITIONS"
                    Format ="#,###"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2546
                            Top =2494
                            Width =1170
                            Height =240
                            Name ="Label15"
                            Caption ="شرايط محيط كار"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =144
                    Top =2834
                    ColumnWidth =1140
                    TabIndex =8
                    Name ="BON"
                    ControlSource ="BON"
                    Format ="#,###"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3476
                            Top =2834
                            Width =240
                            Height =240
                            Name ="Label17"
                            Caption ="بن"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =144
                    Top =3174
                    ColumnWidth =1155
                    TabIndex =9
                    Name ="JAZB"
                    ControlSource ="JAZB"
                    Format ="#,###"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3011
                            Top =3174
                            Width =705
                            Height =240
                            Name ="Label19"
                            Caption ="حق جذب"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =144
                    Top =3514
                    ColumnWidth =1005
                    TabIndex =10
                    Name ="SAYER"
                    ControlSource ="SAYER"
                    Format ="#,###"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3311
                            Top =3514
                            Width =405
                            Height =240
                            Name ="Label21"
                            Caption ="ساير"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "PHOKMSUB.cls"
