Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    AllowEdits = NotDefault
    DefaultView =2
    AllowUpdating =4
    ViewsAllowed =2
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =7237
    RowHeight =315
    ItemSuffix =24
    Left =405
    Top =930
    Right =11595
    Bottom =2370
    DatasheetGridlinesColor =12632256
    Tag ="@Forms___Baseknow___ADA =[Forms]![Baseknow]![ADA]"
    RecSrcDt = Begin
        0x1969f7c8b1f2e340
    End
    InputParameters ="@Forms___Baseknow___ADA =[Forms]![Baseknow]![ADA],@Forms___Baseknow___ADV =[Form"
        "s]![Baseknow]![ADV]"
    RecordSource ="dbo.ASNAD_WITHOUT_CHEK"
    Caption ="اسناد حسابداري دريافت چك كه چك آنها در دفتر چك موجود نيست"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Tahoma"
    UnknownProp = {258 ,66 ,12 ,4 ,198 } Begin
        0x400046006f0072006d0073005f005f005f0042006100730065006b006e006f00 ,
        0x77005f005f005f0041004400410020003d005b0046006f0072006d0073005d00 ,
        0x21005b0042006100730065006b006e006f0077005d0021005b00410044004100 ,
        0x5d002c00400046006f0072006d0073005f005f005f0042006100730065006b00 ,
        0x6e006f0077005f005f005f0041004400560020003d005b0046006f0072006d00 ,
        0x73005d0021005b0042006100730065006b006e006f0077005d0021005b004100 ,
        0x440056005d00
    End
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
            Height =4215
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3408
                    Top =113
                    ColumnOrder =1
                    Name ="N_SERI"
                    ControlSource ="N_SERI"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6020
                            Top =113
                            Width =960
                            Height =240
                            Name ="Label1"
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
                    Left =3408
                    Top =453
                    ColumnOrder =2
                    TabIndex =1
                    Name ="BANK"
                    ControlSource ="BANK"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6605
                            Top =453
                            Width =375
                            Height =240
                            Name ="Label3"
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
                    Left =3408
                    Top =793
                    ColumnOrder =3
                    TabIndex =2
                    Name ="BED"
                    ControlSource ="BED"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6455
                            Top =793
                            Width =525
                            Height =240
                            Name ="Label5"
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
                    Left =3408
                    Top =1133
                    ColumnOrder =0
                    TabIndex =3
                    Name ="N_S"
                    ControlSource ="N_S"
                    OnDblClick ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6095
                            Top =1133
                            Width =885
                            Height =240
                            Name ="Label7"
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
                    Left =3408
                    Top =1474
                    TabIndex =4
                    Name ="HES_K"
                    ControlSource ="HES_K"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6200
                            Top =1474
                            Width =780
                            Height =240
                            Name ="Label9"
                            Caption ="حساب كل"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3408
                    Top =1814
                    TabIndex =5
                    Name ="HES_M"
                    ControlSource ="HES_M"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6035
                            Top =1814
                            Width =945
                            Height =240
                            Name ="Label11"
                            Caption ="حساب معين"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3408
                    Top =2154
                    TabIndex =6
                    Name ="HES_T"
                    ControlSource ="HES_T"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5945
                            Top =2154
                            Width =1035
                            Height =240
                            Name ="Label13"
                            Caption ="معين تفضيلي"
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
                    Top =2494
                    Width =4965
                    ColumnWidth =5325
                    TabIndex =7
                    Name ="SHARH"
                    ControlSource ="SHARH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6560
                            Top =2494
                            Width =420
                            Height =240
                            Name ="Label15"
                            Caption ="شرح"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3408
                    Top =2834
                    TabIndex =8
                    Name ="NUMBER"
                    ControlSource ="NUMBER"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6020
                            Top =2834
                            Width =960
                            Height =240
                            Name ="Label17"
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
                    Left =3408
                    Top =3174
                    TabIndex =9
                    Name ="TAG"
                    ControlSource ="TAG"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6350
                            Top =3174
                            Width =630
                            Height =240
                            Name ="Label19"
                            Caption ="برچسب"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3408
                    Top =3514
                    TabIndex =10
                    Name ="PAY_GETD.N_SERI"
                    ControlSource ="pN_SERI"
                    EventProcPrefix ="PAY_GETD_N_SERI"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5975
                            Top =3514
                            Width =1005
                            Height =240
                            Name ="Label21"
                            Caption ="شماره سريال"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3408
                    Top =3855
                    TabIndex =11
                    Name ="PAY_GETD.BANK"
                    ControlSource ="pBANK"
                    EventProcPrefix ="PAY_GETD_BANK"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6605
                            Top =3855
                            Width =375
                            Height =240
                            Name ="Label23"
                            Caption ="بانك"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "FASNAD_WITHOUT_CHEK.cls"
