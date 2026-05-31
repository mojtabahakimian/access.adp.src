Version =20
VersionRequired =20
Begin Form
    AllowDeletions = NotDefault
    AllowAdditions = NotDefault
    KeyPreview = NotDefault
    AllowEdits = NotDefault
    DefaultView =2
    AllowUpdating =4
    ViewsAllowed =2
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =7237
    RowHeight =375
    ItemSuffix =30
    Left =780
    Top =1680
    Right =11760
    Bottom =6525
    DatasheetGridlinesColor =12632256
    MaxRecords =100
    ShortcutMenuBar ="SERVERFILTER"
    OnApplyFilter ="[Event Procedure]"
    RecSrcDt = Begin
        0xe6541e5528d1e240
    End
    RecordSource ="SEARCH_ON SANAD"
    Caption ="جستجو در شرح سند"
    DatasheetFontName ="Tahoma"
    OnKeyPress ="[Event Procedure]"
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
            Height =4409
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3405
                    Top =480
                    ColumnWidth =990
                    ColumnOrder =0
                    Name ="N_S"
                    ControlSource ="N_S"
                    OnDblClick ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6092
                            Top =480
                            Width =885
                            Height =240
                            Name ="Label1"
                            Caption =" سند"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =150
                    Top =1830
                    Width =4965
                    ColumnWidth =7965
                    ColumnOrder =5
                    TabIndex =3
                    Name ="SHARH"
                    ControlSource ="SHARH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6566
                            Top =1830
                            Width =420
                            Height =240
                            Name ="Label9"
                            Caption ="شرح"
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
                    Left =3414
                    Top =2170
                    ColumnWidth =1560
                    ColumnOrder =6
                    TabIndex =4
                    Name ="BED"
                    ControlSource ="BED"
                    Format ="#,##0.00;-#,##0.00"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6461
                            Top =2170
                            Width =525
                            Height =240
                            Name ="Label11"
                            Caption ="بدهكار"
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
                    Left =3414
                    Top =2510
                    ColumnWidth =1545
                    ColumnOrder =7
                    TabIndex =5
                    Name ="BES"
                    ControlSource ="BES"
                    Format ="#,##0.00;-#,##0.00"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6341
                            Top =2510
                            Width =645
                            Height =240
                            Name ="Label13"
                            Caption ="بستانكار"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3414
                    Top =2850
                    ColumnWidth =990
                    ColumnOrder =9
                    TabIndex =6
                    Name ="N_SERI"
                    ControlSource ="N_SERI"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6026
                            Top =2850
                            Width =960
                            Height =240
                            Name ="Label15"
                            Caption ="شماره سري"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3414
                    Top =3190
                    ColumnWidth =465
                    ColumnOrder =10
                    TabIndex =7
                    Name ="BANK"
                    ControlSource ="BANK"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6611
                            Top =3190
                            Width =375
                            Height =240
                            Name ="Label17"
                            Caption ="بانك"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3414
                    Top =3530
                    ColumnWidth =1125
                    ColumnOrder =11
                    TabIndex =8
                    Name ="NUMBER"
                    ControlSource ="NUMBER"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6026
                            Top =3530
                            Width =960
                            Height =240
                            Name ="Label19"
                            Caption ="شماره فاكتور"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3414
                    Top =3870
                    ColumnWidth =645
                    ColumnOrder =12
                    TabIndex =9
                    Name ="TAG"
                    ControlSource ="TAG"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6356
                            Top =3870
                            Width =630
                            Height =240
                            Name ="Label21"
                            Caption ="برچسب"
                        End
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1611
                    Top =4169
                    Width =3516
                    ColumnWidth =5250
                    ColumnOrder =8
                    TabIndex =10
                    Name ="hes"
                    ControlSource ="NAME"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5802
                            Top =4166
                            Width =1200
                            Height =240
                            Name ="Label22"
                            Caption ="شرح حساب"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2437
                    Top =113
                    ColumnWidth =1020
                    ColumnOrder =1
                    TabIndex =11
                    Name ="DATE_S"
                    ControlSource ="DATE_S"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5029
                            Top =113
                            Width =810
                            Height =240
                            Name ="Label27"
                            Caption ="تاريخ سند"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1143
                    Top =793
                    Width =3966
                    ColumnWidth =525
                    ColumnOrder =2
                    TabIndex =1
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
                            Caption =" كل"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1133
                    Top =1190
                    Width =3966
                    ColumnWidth =585
                    ColumnOrder =3
                    TabIndex =2
                    Name ="HES_M"
                    ControlSource ="HES_M"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6025
                            Top =1190
                            Width =945
                            Height =240
                            Name ="Label7"
                            Caption =" معين"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1143
                    Top =1493
                    Width =3966
                    ColumnWidth =825
                    ColumnOrder =4
                    TabIndex =12
                    Name ="TNUMBER"
                    ControlSource ="TNUMBER"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6030
                            Top =1500
                            Width =1140
                            Height =240
                            Name ="Label29"
                            Caption =" تفصيلي"
                        End
                    End
                End
            End
        End
        Begin FormFooter
            Height =1134
            BackColor =-2147483633
            Name ="FormFooter"
            Begin
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2267
                    Top =113
                    Name ="SBED"
                    ControlSource ="=Sum([BED])"
                    Format ="#,###"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5314
                            Top =113
                            Width =525
                            Height =240
                            Name ="Label24"
                            Caption ="بدهكار"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2267
                    Top =473
                    TabIndex =1
                    Name ="SBES"
                    ControlSource ="=Sum([BES])"
                    Format ="#,###"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5194
                            Top =473
                            Width =645
                            Height =240
                            Name ="Label26"
                            Caption ="بستانكار"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "DEED_SEARCH.cls"
