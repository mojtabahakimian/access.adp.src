Version =20
VersionRequired =20
Begin Form
    AllowDeletions = NotDefault
    AllowAdditions = NotDefault
    AllowEdits = NotDefault
    DefaultView =2
    AllowUpdating =4
    ViewsAllowed =2
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =5062
    RowHeight =270
    ItemSuffix =23
    Left =405
    Top =630
    Right =15105
    Bottom =8685
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xb8413cdc4ba9e240
    End
    ServerFilter ="DT >= 840101  AND DT <= 840619"
    RecordSource ="CHEK_BARGASHTI"
    Caption ="ليست چكهاي  برگشتي"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Tahoma"
    OnLoad ="[Event Procedure]"
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
        Begin FormHeader
            Height =0
            BackColor =-2147483633
            Name ="FormHeader"
        End
        Begin Section
            Height =4025
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1233
                    Top =113
                    Height =255
                    ColumnWidth =1095
                    ColumnOrder =0
                    Name ="N_SERI"
                    ControlSource ="N_SERI"
                    Format ="Fixed"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3800
                            Top =113
                            Width =1005
                            Height =240
                            Name ="Label1"
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
                    Left =1233
                    Top =793
                    Height =255
                    ColumnWidth =1125
                    ColumnOrder =1
                    TabIndex =1
                    Name ="DATE_S"
                    ControlSource ="DATE_S"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3755
                            Top =793
                            Width =1050
                            Height =240
                            Name ="Label5"
                            Caption ="تاريخ سررسيد"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1233
                    Top =1133
                    Height =255
                    ColumnWidth =1110
                    ColumnOrder =2
                    TabIndex =2
                    Name ="DATE"
                    ControlSource ="DATE"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3920
                            Top =1133
                            Width =885
                            Height =240
                            Name ="Label7"
                            Caption ="تاريخ دريافت"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1224
                    Top =1474
                    Width =1710
                    Height =255
                    ColumnOrder =5
                    TabIndex =3
                    Name ="SHOBEH"
                    ControlSource ="SHOBEH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4085
                            Top =1474
                            Width =720
                            Height =240
                            Name ="Label9"
                            Caption ="نام شعبه"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1233
                    Top =1814
                    Height =255
                    ColumnOrder =6
                    TabIndex =4
                    Name ="MABL"
                    ControlSource ="MABL"
                    Format ="#,###"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4430
                            Top =1814
                            Width =375
                            Height =240
                            Name ="Label11"
                            Caption ="مبلغ"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =354
                    Top =2154
                    Width =2580
                    Height =255
                    ColumnWidth =2325
                    ColumnOrder =7
                    TabIndex =5
                    Name ="NAME_TAH"
                    ControlSource ="NAME_TAH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3630
                            Top =2160
                            Width =1260
                            Height =240
                            Name ="Label13"
                            Caption ="نام  پرداخت كننده"
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
                    Left =1233
                    Top =2494
                    Height =255
                    ColumnOrder =8
                    TabIndex =6
                    Name ="N_HESAB"
                    ControlSource ="N_HESAB"
                    Format ="Fixed"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3770
                            Top =2494
                            Width =1035
                            Height =240
                            Name ="Label15"
                            Caption ="شماره حساب"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1233
                    Top =2834
                    Height =255
                    ColumnOrder =9
                    TabIndex =7
                    Name ="N_S"
                    ControlSource ="N_S"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3920
                            Top =2834
                            Width =885
                            Height =240
                            Name ="Label17"
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
                    Left =144
                    Top =3174
                    Width =2790
                    Height =255
                    ColumnOrder =4
                    TabIndex =8
                    Name ="NAMES"
                    ControlSource ="NAMES"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4190
                            Top =3174
                            Width =615
                            Height =240
                            Name ="Label19"
                            Caption ="نام بانك"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1245
                    Top =453
                    Height =255
                    ColumnOrder =10
                    TabIndex =9
                    Name ="RADIF"
                    ControlSource ="RADIF"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4177
                            Top =450
                            Width =630
                            Height =240
                            Name ="Label20"
                            Caption ="شماره "
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1247
                    Top =3571
                    Height =255
                    ColumnOrder =3
                    TabIndex =10
                    Name ="DT"
                    ControlSource ="DT"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3927
                            Top =3575
                            Width =960
                            Height =240
                            Name ="Label22"
                            Caption ="تاريخ برگشت"
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
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =737
                    Top =339
                    Name ="smabl"
                    ControlSource ="=Sum([MABL])"
                    Format ="#,###"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3078
                            Top =283
                            Width =375
                            Height =240
                            Name ="Label21"
                            Caption ="مبلغ"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "CHEK_BARGASHTI.cls"
