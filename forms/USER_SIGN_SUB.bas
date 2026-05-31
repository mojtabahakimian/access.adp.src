Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    ControlBox = NotDefault
    ShortcutMenu = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    CloseButton = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    KeyPreview = NotDefault
    AllowEdits = NotDefault
    DefaultView =2
    AllowUpdating =4
    ScrollBars =0
    ViewsAllowed =2
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =0
    GridY =10
    Width =3403
    RowHeight =255
    ItemSuffix =9
    Left =2355
    Top =2190
    Right =4680
    Bottom =8145
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xe4a16579aad5e240
    End
    RecordSource ="SELECT SAL_NAME, PSAL_NAME, GRSAL, ENABL, IDD FROM SALA_DTL WHERE (ENABL = 0) AN"
        "D (IDD <> 1)"
    OnCurrent ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Tahoma"
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
            Height =1605
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
                    Left =1702
                    ColumnWidth =1755
                    ColumnOrder =1
                    Name ="SAL_NAME1"
                    ControlSource ="=DECODEUN([SAL_NAME])"
                    OnDblClick ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Width =1230
                            Height =240
                            Name ="Label0"
                            Caption ="نام كاربر"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1702
                    Top =680
                    ColumnWidth =0
                    ColumnOrder =2
                    TabIndex =1
                    Name ="GRSAL"
                    ControlSource ="GRSAL"
                    StatusBarText ="0بدون گروه 1 فرمانده 2 حسابدار 3 فروش 4 انبار5 كارگزيني10 صاحب"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Top =680
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
                    Left =1697
                    Top =963
                    ColumnWidth =15
                    ColumnOrder =0
                    TabIndex =2
                    Name ="IDD"
                    ControlSource ="IDD"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Top =967
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
                    Left =1702
                    Top =360
                    TabIndex =3
                    Name ="SAL_NAME"
                    ControlSource ="SAL_NAME"
                End
            End
        End
    End
End
CodeBehindForm
' See "USER_SIGN_SUB.cls"
