Version =20
VersionRequired =20
Begin Form
    AllowDeletions = NotDefault
    AllowAdditions = NotDefault
    OrderByOn = NotDefault
    DefaultView =2
    AllowUpdating =4
    ViewsAllowed =2
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =7237
    RowHeight =300
    ItemSuffix =33
    Left =75
    Top =855
    Right =14670
    Bottom =7200
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x4ced0b3351efe340
    End
    RecordSource ="MOGH_SUB"
    Caption ="MO_DTL"
    OnCurrent ="[Event Procedure]"
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
        Begin FormHeader
            Height =0
            BackColor =-2147483633
            Name ="FormHeader"
        End
        Begin Section
            Height =4555
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    ColumnHidden = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3408
                    Top =113
                    ColumnWidth =540
                    ColumnOrder =0
                    Name ="MONUM"
                    ControlSource ="MONUM"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6455
                            Top =113
                            Width =525
                            Height =240
                            Name ="Label1"
                            Caption ="شماره"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =93
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3408
                    Top =453
                    ColumnWidth =615
                    ColumnOrder =2
                    TabIndex =1
                    Name ="N_S"
                    ControlSource ="N_S"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6095
                            Top =453
                            Width =885
                            Height =240
                            Name ="Label3"
                            Caption =" سند"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    ColumnHidden = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3408
                    Top =793
                    ColumnWidth =765
                    ColumnOrder =3
                    TabIndex =2
                    Name ="HES_K"
                    ControlSource ="HES_K"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6200
                            Top =793
                            Width =780
                            Height =240
                            Name ="Label5"
                            Caption ="حساب كل"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    ColumnHidden = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3408
                    Top =1133
                    ColumnWidth =915
                    ColumnOrder =4
                    TabIndex =3
                    Name ="HES_M"
                    ControlSource ="HES_M"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6035
                            Top =1133
                            Width =945
                            Height =240
                            Name ="Label7"
                            Caption ="حساب معين"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    ColumnHidden = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3408
                    Top =1474
                    ColumnWidth =180
                    ColumnOrder =5
                    TabIndex =4
                    Name ="HES_T"
                    ControlSource ="HES_T"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5945
                            Top =1474
                            Width =1035
                            Height =240
                            Name ="Label9"
                            Caption ="معين تفضيلي"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =144
                    Top =1814
                    Width =4965
                    ColumnWidth =5925
                    ColumnOrder =6
                    TabIndex =5
                    Name ="SHARH"
                    ControlSource ="SHARH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6560
                            Top =1814
                            Width =420
                            Height =240
                            Name ="Label11"
                            Caption ="شرح"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3408
                    Top =2154
                    ColumnWidth =1560
                    ColumnOrder =7
                    TabIndex =6
                    Name ="BED"
                    ControlSource ="BED"
                    Format ="#,###"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6455
                            Top =2154
                            Width =525
                            Height =240
                            Name ="Label13"
                            Caption ="بدهكار"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3408
                    Top =2494
                    ColumnWidth =1515
                    ColumnOrder =8
                    TabIndex =7
                    Name ="BES"
                    ControlSource ="BES"
                    Format ="#,###"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6335
                            Top =2494
                            Width =645
                            Height =240
                            Name ="Label15"
                            Caption ="بستانكار"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    ReadingOrder =1
                    Left =4832
                    Top =2891
                    ColumnWidth =435
                    ColumnOrder =9
                    TabIndex =8
                    Name ="TICK"
                    ControlSource ="TICK"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6635
                            Top =2834
                            Width =345
                            Height =240
                            Name ="Label17"
                            Caption ="تيك"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3408
                    Top =3174
                    ColumnOrder =10
                    TabIndex =9
                    Name ="N_SERI"
                    ControlSource ="N_SERI"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6020
                            Top =3174
                            Width =960
                            Height =240
                            Name ="Label19"
                            Caption ="شماره سري"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3408
                    Top =3514
                    ColumnOrder =11
                    TabIndex =10
                    Name ="BANK"
                    ControlSource ="BANK"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6605
                            Top =3514
                            Width =375
                            Height =240
                            Name ="Label21"
                            Caption ="بانك"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3408
                    Top =3855
                    ColumnOrder =12
                    TabIndex =11
                    Name ="NUMBER"
                    ControlSource ="NUMBER"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6020
                            Top =3855
                            Width =960
                            Height =240
                            Name ="Label23"
                            Caption ="شماره فاكتور"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3408
                    Top =4195
                    ColumnWidth =645
                    ColumnOrder =13
                    TabIndex =12
                    Name ="TAG"
                    ControlSource ="TAG"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6350
                            Top =4195
                            Width =630
                            Height =240
                            Name ="Label25"
                            Caption ="برچسب"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =226
                    Top =453
                    ColumnWidth =1020
                    ColumnOrder =1
                    TabIndex =13
                    Name ="DATE_S"
                    ControlSource ="DATE_S"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =2818
                            Top =453
                            Width =810
                            Height =240
                            Name ="Label32"
                            Caption ="تاريخ سند"
                        End
                    End
                End
            End
        End
        Begin FormFooter
            Height =1280
            BackColor =-2147483633
            Name ="FormFooter"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Name ="SBED"
                    ControlSource ="=Sum([BED])"
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Top =340
                    TabIndex =1
                    Name ="SBES"
                    ControlSource ="=Sum([BES])"
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    IMEMode =1
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2154
                    TabIndex =2
                    Name ="STBED"
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2154
                    Top =340
                    TabIndex =3
                    Name ="STBES"
                End
            End
        End
    End
End
CodeBehindForm
' See "MO_DTL_SUB.cls"
