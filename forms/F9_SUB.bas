Version =20
VersionRequired =20
Begin Form
    OrderByOn = NotDefault
    DefaultView =2
    AllowUpdating =4
    ViewsAllowed =2
    TabularCharSet =178
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =6463
    RowHeight =405
    ItemSuffix =9
    Left =5295
    Top =3450
    Right =11235
    Bottom =4890
    DatasheetForeColor =33554432
    DatasheetGridlinesColor =12632256
    MaxRecords =1000000
    RecSrcDt = Begin
        0x4b780716cf2be440
    End
    RecordSource ="JAMFACTPRS"
    DatasheetFontName ="Tahoma"
    AllowFormView =0
    RecordSourceQualifier ="dbo"
    Begin
        Begin Label
            BackStyle =0
            TextFontCharSet =178
            FontName ="Tahoma"
        End
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            OldBorderStyle =0
            TextFontCharSet =178
            Width =1701
            LabelX =-1701
            FontName ="Tahoma"
            AsianLineBreak =255
        End
        Begin Subform
            SpecialEffect =2
            Width =1701
            Height =1701
        End
        Begin Section
            CanGrow = NotDefault
            Height =5102
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =4762
                    Top =907
                    ColumnWidth =630
                    Name ="NUMBER"
                    ControlSource ="NUMBER"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3061
                            Top =907
                            Width =750
                            Height =240
                            Name ="Label0"
                            Caption ="شماره"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =4762
                    Top =1247
                    ColumnWidth =1417
                    TabIndex =1
                    Name ="TAG"
                    ControlSource ="TAG"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3061
                            Top =1247
                            Width =435
                            Height =240
                            Name ="Label1"
                            Caption ="TAG:"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =4762
                    Top =1587
                    ColumnWidth =120
                    TabIndex =2
                    Name ="CUST_NO"
                    ControlSource ="CUST_NO"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3061
                            Top =1587
                            Width =840
                            Height =240
                            Name ="Label2"
                            Caption ="مشتري"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =4762
                    Top =1927
                    TabIndex =3
                    Name ="DATE_N"
                    ControlSource ="DATE_N"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3061
                            Top =1927
                            Width =885
                            Height =240
                            Name ="Label3"
                            Caption ="تاريخ فاكتور"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =4762
                    Top =2267
                    ColumnWidth =1815
                    TabIndex =4
                    Name ="mab"
                    ControlSource ="mab"
                    Format ="#,##0.00;-#,##0.00"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3061
                            Top =2267
                            Width =435
                            Height =240
                            Name ="Label4"
                            Caption ="مبلغ"
                        End
                    End
                End
                Begin Subform
                    OverlapFlags =87
                    Left =453
                    Top =3401
                    TabIndex =5
                    Name ="mandfactormr_visit_form"
                    SourceObject ="Form.mandfactormr_visit_form"
                    LinkChildFields ="NUMBER;TAG2"
                    LinkMasterFields ="NUMBER;TAG"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =699
                            Top =3161
                            Width =1455
                            Height =240
                            Name ="mandfactormr_visit Label"
                            Caption ="mandfactormr_visit"
                            EventProcPrefix ="mandfactormr_visit_Label"
                        End
                    End
                End
            End
        End
    End
End
