Version =20
VersionRequired =20
Begin Form
    NavigationButtons = NotDefault
    DefaultView =2
    AllowUpdating =4
    ViewsAllowed =2
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =3973
    RowHeight =330
    ItemSuffix =9
    Left =570
    Top =1770
    Right =6825
    Bottom =4110
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x4df2d4c10c9be240
    End
    RecordSource ="TAKHPERS"
    Caption ="تخفيفات مصوب"
    AfterUpdate ="[Event Procedure]"
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
        Begin Section
            Height =1827
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1485
                    Top =113
                    Width =1005
                    ColumnWidth =1275
                    Name ="TAKH_COD"
                    ControlSource ="TAKH_COD"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =2576
                            Top =113
                            Width =1140
                            Height =240
                            Name ="Label1"
                            Caption ="كد كالا و خدمات"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =144
                    Top =793
                    Width =2346
                    ColumnWidth =1035
                    TabIndex =2
                    Name ="TAFPER"
                    ControlSource ="TAFPER"
                    InputMask ="\\%###"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =2786
                            Top =793
                            Width =930
                            Height =240
                            Name ="Label5"
                            Caption ="درصد تخفيف"
                        End
                    End
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =144
                    Top =453
                    Width =2346
                    ColumnWidth =1950
                    TabIndex =1
                    Name ="CUST_CO"
                    ControlSource ="CUST_CO"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT CUSTKIND.CUST_COD, CUSTKIND.CUSTKNAME FROM CUSTKIND;"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =2606
                            Top =453
                            Width =1110
                            Height =240
                            Name ="Label3"
                            Caption ="كد نوع مشتري"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =144
                    Top =1153
                    Width =2346
                    TabIndex =3
                    Name ="PRICE_M"
                    ControlSource ="PRICE_M"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =2786
                            Top =1153
                            Width =930
                            Height =240
                            Name ="Label8"
                            Caption ="مبلغ"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "TAKHPERS_SUB.cls"
