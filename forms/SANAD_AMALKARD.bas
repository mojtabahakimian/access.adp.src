Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    DefaultView =0
    AllowUpdating =4
    ScrollBars =0
    ViewsAllowed =1
    TabularCharSet =178
    TabularFamily =42
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =9127
    DatasheetFontHeight =10
    ItemSuffix =3
    Left =4995
    Top =3105
    Right =14115
    Bottom =4785
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x330fc2c8d095e240
    End
    Caption ="صدور سند عملكرد"
    DatasheetFontName ="Arial (Arabic)"
    OnError ="[Event Procedure]"
    AllowDatasheetView =0
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
        Begin TextBox
            SpecialEffect =2
            OldBorderStyle =0
            TextFontCharSet =178
            Width =1701
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
            Height =1701
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =42
                    Left =3174
                    Top =963
                    Width =1927
                    Height =567
                    FontWeight =700
                    Name ="Command0"
                    Caption ="صدور سند"
                    OnClick ="[Event Procedure]"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =225
                    Top =113
                    Width =6467
                    Height =330
                    FontWeight =700
                    TabIndex =1
                    Name ="AMAL"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT hes, NAME AS Expr1 FROM CUST_HESAB"
                    ColumnWidths ="1701"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =6810
                            Top =115
                            Width =2205
                            Height =330
                            FontWeight =700
                            Name ="Label2"
                            Caption ="حساب كل سود وزيان جاري"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "SANAD_AMALKARD.cls"
