Version =20
VersionRequired =20
Begin Form
    NavigationButtons = NotDefault
    OrderByOn = NotDefault
    DefaultView =2
    AllowUpdating =4
    ViewsAllowed =2
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =3973
    RowHeight =330
    ItemSuffix =7
    Left =705
    Top =4110
    Right =4635
    Bottom =6150
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xe178a0722ef7e340
    End
    RecordSource ="TR_TAKHPERS"
    Caption ="تخفيفات مصوب"
    OnCurrent ="[Event Procedure]"
    AfterUpdate ="[Event Procedure]"
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
            Height =2961
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =339
                    Top =1814
                    Name ="UP_DATE"
                    ControlSource ="UP_DATE"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =2848
                            Top =1814
                            Width =810
                            Height =240
                            Name ="Label48"
                            Caption ="UP_DATE:"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =283
                    Top =2268
                    TabIndex =1
                    Name ="UP_TIME"
                    ControlSource ="UP_TIME"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =2905
                            Top =2268
                            Width =780
                            Height =240
                            Name ="Label49"
                            Caption ="UP_TIME:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1341
                    Width =1005
                    TabIndex =2
                    Name ="TAKH_COD"
                    ControlSource ="TAKH_COD"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2432
                            Width =1140
                            Height =240
                            Name ="Label1"
                            Caption ="كد كالا و خدمات"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Top =680
                    Width =2346
                    TabIndex =3
                    Name ="TAFPER"
                    ControlSource ="TAFPER"
                    InputMask ="\\%###"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2642
                            Top =680
                            Width =930
                            Height =240
                            Name ="Label5"
                            Caption ="% تخفيف"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Top =340
                    Width =2346
                    TabIndex =4
                    Name ="CUST_CO"
                    ControlSource ="CUST_CO"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT CUSTKIND.CUST_COD, CUSTKIND.CUSTKNAME FROM CUSTKIND;"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2462
                            Top =340
                            Width =1110
                            Height =240
                            Name ="Label3"
                            Caption ="كد نوع مشتري"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =536
                    Top =1134
                    TabIndex =5
                    Name ="PRICE_M"
                    ControlSource ="PRICE_M"
                    Format ="#,##0.00;-#,##0.00"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3034
                            Top =1134
                            Width =795
                            Height =240
                            Name ="Label6"
                            Caption ="قيمت"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =247
                    TextAlign =3
                    ReadingOrder =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =422
                    Top =1927
                    TabIndex =6
                    Name ="PERS"
                    ControlSource ="PERS"
                    Begin
                        Begin Label
                            OverlapFlags =223
                            Left =3321
                            Top =1927
                            Width =555
                            Height =405
                            Name ="Label7"
                            Caption ="درصد +-"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =247
                    TextAlign =3
                    ReadingOrder =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =422
                    Top =2268
                    TabIndex =7
                    Name ="BLNS"
                    ControlSource ="BLNS"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =3336
                            Top =2272
                            Width =555
                            Height =405
                            Name ="Label8"
                            Caption ="بالانس+-"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    ReadingOrder =1
                    Left =1013
                    Top =2721
                    TabIndex =8
                    Name ="PUT"
                    ControlSource ="PUT"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1506
                            Top =2692
                            Width =1485
                            Height =240
                            Name ="Label11"
                            Caption ="قيمت گذاري گروهي"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "TR_TAKHPERS_SUB2.cls"
