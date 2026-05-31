Version =20
VersionRequired =20
Begin Form
    KeyPreview = NotDefault
    OrderByOn = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =2
    AllowUpdating =4
    ViewsAllowed =2
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =7031
    RowHeight =330
    ItemSuffix =16
    Left =465
    Top =435
    Right =6435
    Bottom =3840
    DatasheetGridlinesColor =12632256
    OrderBy ="IDD"
    RecSrcDt = Begin
        0xac11994a7ec3e540
    End
    RecordSource ="SELECT SAL_NAME, PSAL_NAME, GRSAL, ENABL, IDD, HES, PORID, EMZA FROM SALA_DTL WH"
        "ERE (ENABL = 0) AND (GRSAL <> 11)"
    DatasheetFontName ="Tahoma"
    OnLoad ="[Event Procedure]"
    Moveable =0
    AllowFormView =0
    FetchDefaults =0
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
        End
        Begin BoundObjectFrame
            SpecialEffect =2
            Width =4536
            Height =2835
            LabelX =-1701
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
        Begin Section
            Height =6633
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =3629
                    Top =510
                    Height =255
                    ColumnWidth =2760
                    ColumnOrder =1
                    Name ="SAL_NAME1"
                    ControlSource ="=Trim(CODESAL(Fixp([SAL_NAME])))"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1927
                            Top =510
                            Width =1230
                            Height =240
                            Name ="Label0"
                            Caption ="نام كاربر"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =3629
                    Top =1190
                    ColumnOrder =2
                    TabIndex =1
                    Name ="GRSAL"
                    ControlSource ="GRSAL"
                    StatusBarText ="0بدون گروه 1 فرمانده 2 حسابدار 3 فروش 4 انبار5 كارگزيني10 صاحب"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1927
                            Top =1190
                            Width =930
                            Height =240
                            Name ="Label2"
                            Caption ="گروه"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =3629
                    Top =1871
                    ColumnOrder =0
                    TabIndex =2
                    Name ="IDD"
                    ControlSource ="IDD"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1932
                            Top =1875
                            Width =1635
                            Height =240
                            Name ="Label4"
                            Caption ="شماره شناسائي"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =3629
                    Top =870
                    TabIndex =3
                    Name ="SAL_NAME"
                    ControlSource ="SAL_NAME"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3175
                    Top =2381
                    ColumnWidth =2835
                    TabIndex =4
                    Name ="Text9"
                    ControlSource ="SAL_NAME"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6022
                            Top =2381
                            Width =555
                            Height =240
                            Name ="Label10"
                            Caption ="name"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2835
                    Top =3118
                    TabIndex =5
                    Name ="HES"
                    ControlSource ="HES"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5817
                            Top =3118
                            Width =420
                            Height =240
                            Name ="Label11"
                            Caption ="HES:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2835
                    Top =3458
                    ColumnWidth =435
                    TabIndex =6
                    Name ="PORID"
                    ControlSource ="PORID"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5622
                            Top =3458
                            Width =615
                            Height =240
                            Name ="Label12"
                            Caption ="PORID:"
                        End
                    End
                End
                Begin BoundObjectFrame
                    OverlapFlags =85
                    Top =3798
                    TabIndex =7
                    Name ="EMZA"
                    ControlSource ="EMZA"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5697
                            Top =3798
                            Width =540
                            Height =240
                            Name ="Label13"
                            Caption ="EMZA:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1702
                    TabIndex =8
                    Name ="Text14"
                    ControlSource ="=DECODEUN([SAL_NAME])"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Width =1230
                            Height =240
                            Name ="Label15"
                            Caption ="نام كاربر"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "USER_LOGIN_SUB_update.cls"
