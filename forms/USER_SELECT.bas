Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    Modal = NotDefault
    ShortcutMenu = NotDefault
    AutoCenter = NotDefault
    AllowDeletions = NotDefault
    AllowAdditions = NotDefault
    KeyPreview = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =2
    AllowUpdating =4
    ViewsAllowed =2
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =5669
    RowHeight =330
    ItemSuffix =12
    Left =4380
    Top =2970
    Right =11340
    Bottom =9480
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xcbd45f71fc87e540
    End
    RecordSource ="SELECT temp_us.USERCO, temp_us.tic, temp_us.USERCOT, SALA_DTL.SAL_NAME, temp_us."
        "IDD FROM temp_us INNER JOIN SALA_DTL ON temp_us.USERCO = SALA_DTL.IDD WHERE (tem"
        "p_us.USERCOT = 1) ORDER BY temp_us.IDD, temp_us.tic DESC"
    Caption ="جهت انتخاب همه روي نام كاربر دابل كليك كنيدو بالعكس روي شماره شناسايي"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Tahoma"
    Moveable =0
    AllowFormView =0
    FetchDefaults =0
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
        End
        Begin CheckBox
            SpecialEffect =2
            LabelX =230
            LabelY =-30
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
            Height =2564
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2207
                    Top =1077
                    ColumnOrder =0
                    TabIndex =1
                    Name ="USERCO"
                    ControlSource ="USERCO"
                    OnDblClick ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =510
                            Top =1081
                            Width =1635
                            Height =240
                            Name ="Label4"
                            Caption ="شماره شناسائي*"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    ReadingOrder =1
                    Left =3028
                    Top =2324
                    ColumnWidth =915
                    TabIndex =2
                    Name ="tic"
                    ControlSource ="tic"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =2368
                            Top =2294
                            Width =690
                            Height =240
                            Name ="Label11"
                            Caption ="انتخاب"
                        End
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2267
                    Top =510
                    ColumnWidth =3705
                    ColumnOrder =1
                    Name ="SAL_NAME1"
                    ControlSource ="=DECODEUN([SAL_NAME])"
                    OnDblClick ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =565
                            Top =510
                            Width =1230
                            Height =240
                            Name ="Label0"
                            Caption ="نام كاربر*"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "USER_SELECT.cls"
