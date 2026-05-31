Version =20
VersionRequired =20
Begin Form
    AllowDeletions = NotDefault
    AllowAdditions = NotDefault
    KeyPreview = NotDefault
    OrderByOn = NotDefault
    AllowEdits = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =2
    AllowUpdating =4
    ViewsAllowed =2
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =5669
    RowHeight =375
    ItemSuffix =9
    Left =11490
    Top =2490
    Right =17475
    Bottom =5880
    DatasheetGridlinesColor =12632256
    Filter ="((USER_LOGIN_SUB.GRSAL=1))"
    OrderBy ="IDD DESC"
    RecSrcDt = Begin
        0xb9fc87e4f00be540
    End
    RecordSource ="SELECT SAL_NAME, PSAL_NAME, GRSAL, ENABL, IDD FROM SALA_DTL WHERE (ENABL = 0) AN"
        "D (GRSAL <> 11)"
    OnCurrent ="[Event Procedure]"
    DatasheetFontName ="Tahoma"
    OnKeyPress ="[Event Procedure]"
    OnLoad ="[Event Procedure]"
    Moveable =0
    AllowFormView =0
    FetchDefaults =0
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
        Begin Section
            Height =2551
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
                    Left =2267
                    Top =510
                    ColumnWidth =2760
                    ColumnOrder =1
                    Name ="SAL_NAME1"
                    ControlSource ="=DECODEUN([SAL_NAME])"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =565
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
                    Left =2267
                    Top =1190
                    ColumnOrder =2
                    TabIndex =1
                    Name ="GRSAL"
                    ControlSource ="GRSAL"
                    StatusBarText ="0بدون گروه 1 فرمانده 2 حسابدار 3 فروش 4 انبار5 كارگزيني10 صاحب"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =565
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
                    Left =2267
                    Top =1871
                    ColumnOrder =0
                    TabIndex =2
                    Name ="IDD"
                    ControlSource ="IDD"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =570
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
                    Left =2267
                    Top =870
                    TabIndex =3
                    Name ="SAL_NAME"
                    ControlSource ="SAL_NAME"
                End
            End
        End
    End
End
CodeBehindForm
' See "USER_LOGIN_SUB.cls"
