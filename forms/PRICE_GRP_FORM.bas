Version =20
VersionRequired =20
Begin Form
    KeyPreview = NotDefault
    DefaultView =2
    AllowUpdating =4
    ViewsAllowed =2
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =4762
    RowHeight =405
    ItemSuffix =6
    Left =1080
    Top =1110
    Right =6150
    Bottom =1995
    DatasheetGridlinesColor =12632256
    Filter ="PGNAME like N'%کاشي%'"
    RecSrcDt = Begin
        0x32f08eaf7397e540
    End
    RecordSource ="PRICE_GRP"
    Caption ="تعريف گروه قيمت گذاري"
    OnCurrent ="[Event Procedure]"
    BeforeUpdate ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Tahoma"
    OnKeyPress ="[Event Procedure]"
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
                    Name ="PGID"
                    ControlSource ="PGID"
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
                    ColumnWidth =10350
                    TabIndex =1
                    Name ="PGNAME"
                    ControlSource ="PGNAME"
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
                    ColumnWidth =2085
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
            End
        End
    End
End
CodeBehindForm
' See "PRICE_GRP_FORM.cls"
