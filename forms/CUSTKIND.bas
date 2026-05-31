Version =20
VersionRequired =20
Begin Form
    OrderByOn = NotDefault
    DefaultView =2
    AllowUpdating =4
    ViewsAllowed =2
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =4762
    RowHeight =300
    ItemSuffix =4
    Left =825
    Top =1110
    Right =5880
    Bottom =1980
    DatasheetGridlinesColor =12632256
    OrderBy ="CUSTKIND.CUST_COD"
    RecSrcDt = Begin
        0x0487b4fcd29ae240
    End
    RecordSource ="CUSTKIND"
    Caption ="تعريف نوع مشتري"
    BeforeUpdate ="[Event Procedure]"
    AfterUpdate ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Tahoma"
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
        Begin Section
            Height =813
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =933
                    Top =113
                    Name ="CUST_COD"
                    ControlSource ="CUST_COD"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3995
                            Top =113
                            Width =510
                            Height =240
                            Name ="Label1"
                            Caption ="كد نوع"
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
                    Top =453
                    Width =2490
                    ColumnWidth =4545
                    TabIndex =1
                    Name ="CUSTKNAME"
                    ControlSource ="CUSTKNAME"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3605
                            Top =453
                            Width =900
                            Height =240
                            Name ="Label3"
                            Caption ="نوع مشتري"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "CUSTKIND.cls"
