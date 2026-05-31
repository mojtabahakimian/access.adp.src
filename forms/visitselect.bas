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
    Width =10674
    RowHeight =450
    ItemSuffix =3
    Left =4170
    Top =120
    Right =12840
    Bottom =3870
    DatasheetGridlinesColor =12632256
    MaxRecords =1000000
    RecSrcDt = Begin
        0xa01518824d7ee440
    End
    RecordSource ="SELECT Visit_route.ROUTE_NAME, Visit_route.HES, Visit_route.IYALAT, Visit_route."
        "CITY, Visit_route.District, Visit_route.RACTIVE, CUST_HESAB.NAME FROM Visit_rout"
        "e INNER JOIN CUST_HESAB ON Visit_route.HES = CUST_HESAB.hes"
    Caption ="انتخاب مسير"
    OnCurrent ="[Event Procedure]"
    DatasheetFontName ="Tahoma"
    OnKeyPress ="[Event Procedure]"
    AllowFormView =0
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
        Begin ComboBox
            SpecialEffect =2
            TextFontCharSet =178
            Width =1701
            LabelX =-1701
            FontName ="Tahoma"
        End
        Begin Section
            Height =3288
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Width =8781
                    Height =317
                    ColumnWidth =10635
                    BackColor =13434828
                    ForeColor =8388608
                    Name ="ROUTE_NAME"
                    ControlSource ="ROUTE_NAME"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT Visit_route.ROUTE_NAME, Visit_route.ROUTE_NAME + N' - ' + CUST_HESAB.NAME"
                        " + N' - ' + CUST_HESAB.hes AS Expr1 FROM Visit_route INNER JOIN CUST_HESAB ON Vi"
                        "sit_route.HES = CUST_HESAB.hes WHERE (Visit_route.RACTIVE = 1)"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =8859
                            Width =1815
                            Height =330
                            ForeColor =8388608
                            Name ="Label26"
                            Caption ="مسير ويزيت"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =2267
                    Top =1870
                    TabIndex =1
                    Name ="IYALAT"
                    ControlSource ="IYALAT"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5009
                            Top =1870
                            Width =660
                            Height =240
                            Name ="Label0"
                            Caption ="استان"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =2380
                    Top =2154
                    TabIndex =2
                    Name ="CITY"
                    ControlSource ="CITY"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5295
                            Top =2160
                            Width =495
                            Height =240
                            Name ="Label1"
                            Caption ="استان"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1587
                    Top =2607
                    TabIndex =3
                    Name ="District"
                    ControlSource ="District"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4365
                            Top =2610
                            Width =675
                            Height =240
                            Name ="Label2"
                            Caption ="توضيحات"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "visitselect.cls"
