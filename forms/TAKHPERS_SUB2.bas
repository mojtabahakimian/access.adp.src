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
    Width =4035
    RowHeight =330
    ItemSuffix =12
    Left =150
    Top =6225
    Right =7515
    Bottom =8865
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x4df2d4c10c9be240
    End
    RecordSource ="TAKHPERS"
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
            Height =3074
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
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
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =144
                    Top =793
                    Width =2346
                    ColumnWidth =735
                    TabIndex =2
                    Name ="TAFPER"
                    ControlSource ="TAFPER"
                    InputMask ="\\%###"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2786
                            Top =793
                            Width =930
                            Height =240
                            Name ="Label5"
                            Caption ="% تخفيف"
                        End
                    End
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =144
                    Top =453
                    Width =2346
                    ColumnWidth =1965
                    TabIndex =1
                    Name ="CUST_CO"
                    ControlSource ="CUST_CO"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT CUSTKIND.CUST_COD, CUSTKIND.CUSTKNAME FROM CUSTKIND;"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
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
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =680
                    Top =1247
                    ColumnWidth =1095
                    TabIndex =3
                    Name ="PRICE_M"
                    ControlSource ="PRICE_M"
                    Format ="#,##0.00;-#,##0.00"
                    OnKeyPress ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3178
                            Top =1247
                            Width =795
                            Height =240
                            Name ="Label6"
                            Caption ="قيمت"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =566
                    Top =2040
                    ColumnWidth =765
                    TabIndex =4
                    Name ="PERS"
                    ControlSource ="PERS"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =3465
                            Top =2040
                            Width =555
                            Height =405
                            Name ="Label7"
                            Caption ="درصد +-"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =566
                    Top =2381
                    ColumnWidth =810
                    TabIndex =5
                    Name ="BLNS"
                    ControlSource ="BLNS"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =3480
                            Top =2385
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
                    Left =1157
                    Top =2834
                    ColumnWidth =1590
                    TabIndex =6
                    Name ="PUT"
                    ControlSource ="PUT"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1650
                            Top =2805
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
' See "TAKHPERS_SUB2.cls"
