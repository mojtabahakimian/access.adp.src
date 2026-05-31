Version =20
VersionRequired =20
Begin Form
    DefaultView =2
    AllowUpdating =4
    ViewsAllowed =2
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =4762
    RowHeight =405
    ItemSuffix =7
    Left =600
    Top =1500
    Right =5640
    Bottom =2370
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xfc0a84489097e540
    End
    RecordSource ="PRICE_PAYNO"
    Caption ="تعريف نحوه پرداخت"
    BeforeUpdate ="[Event Procedure]"
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
            Height =1884
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
                    ColumnOrder =0
                    Name ="PPID"
                    ControlSource ="PPID"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3995
                            Top =113
                            Width =510
                            Height =240
                            Name ="Label1"
                            Caption ="كد "
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
                    ColumnOrder =1
                    TabIndex =1
                    Name ="PPAME"
                    ControlSource ="PPAME"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3605
                            Top =453
                            Width =900
                            Height =240
                            Name ="Label3"
                            Caption ="عنوان"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1417
                    Top =1303
                    ColumnWidth =2805
                    ColumnOrder =2
                    TabIndex =2
                    Name ="TR_DATE"
                    ControlSource ="TR_DATE"
                    DefaultValue ="Now()"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3952
                            Top =1303
                            Width =810
                            Height =240
                            Name ="Label4"
                            Caption ="تاريخ"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1417
                    Top =1644
                    ColumnOrder =4
                    TabIndex =3
                    Name ="USERNAME"
                    ControlSource ="USERNAME"
                    DefaultValue ="UCurrentUser()"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3817
                            Top =1644
                            Width =945
                            Height =240
                            Name ="Label5"
                            Caption ="نام کاربر"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =453
                    Top =850
                    ColumnOrder =3
                    TabIndex =4
                    Name ="MODAT"
                    ControlSource ="MODAT"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3180
                            Top =855
                            Width =810
                            Height =240
                            Name ="Label6"
                            Caption ="مدت به روز"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "PRICE_PAYNO_FORM.cls"
