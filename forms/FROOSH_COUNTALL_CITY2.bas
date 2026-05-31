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
    Width =4940
    RowHeight =330
    ItemSuffix =18
    Left =345
    Top =3855
    Right =7050
    Bottom =5295
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xa7c0c8f675e9e340
    End
    RecordSource ="SELECT SUM(MEGHKF) AS MEGHF, SUM(MABK) AS MABKK, SUM(MEGHkK) AS MEGHk, CUST_NO, "
        "MPNAME FROM FROOSH_COUNTALL_CITY_N_FANI(900501,900602) FROOSH_COUNTALL_CITY_N_FA"
        "NI GROUP BY CUST_NO, MPNAME ORDER BY CUST_NO"
    Caption ="ليست فروش روزانه به تفكيك نوع كالا"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Tahoma"
    AllowFormView =0
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
        End
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            OldBorderStyle =0
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
            Height =2340
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
                    Left =1927
                    Top =113
                    ColumnOrder =3
                    Name ="MEGHk"
                    ControlSource ="MEGHk"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =113
                            Width =570
                            Height =240
                            Name ="Label1"
                            Caption ="متراژ"
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
                    Left =1863
                    Top =396
                    Width =2490
                    ColumnWidth =1260
                    ColumnOrder =2
                    TabIndex =1
                    Name ="MPNAME"
                    ControlSource ="MPNAME"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =56
                            Top =398
                            Width =540
                            Height =240
                            Name ="Label5"
                            Caption ="شهر"
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
                    Left =1927
                    Top =1133
                    ColumnWidth =1065
                    ColumnOrder =4
                    TabIndex =2
                    Name ="MEGHF"
                    ControlSource ="MEGHF"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =1133
                            Width =960
                            Height =240
                            Name ="Label7"
                            Caption ="كارتن"
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
                    Left =1856
                    Top =1587
                    Width =2490
                    ColumnOrder =0
                    TabIndex =3
                    Name ="CUST_NO"
                    ControlSource ="CUST_NO"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =56
                            Top =1592
                            Width =840
                            Height =240
                            Name ="Label11"
                            Caption ="كد مشتري"
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
                    Left =1984
                    Top =1927
                    TabIndex =5
                    Name ="MABKK"
                    ControlSource ="MABKK"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =170
                            Top =1927
                            Width =960
                            Height =240
                            Name ="Label17"
                            Caption ="مبلغ"
                        End
                    End
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =1863
                    Top =758
                    Width =2490
                    ColumnWidth =4050
                    ColumnOrder =1
                    TabIndex =4
                    Name ="NAME"
                    ControlSource ="CUST_NO"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT hes, NAME FROM CUST_HESAB"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =765
                            Width =870
                            Height =240
                            Name ="Label15"
                            Caption ="نام مشتري"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "FROOSH_COUNTALL_CITY2.cls"
