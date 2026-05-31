Version =20
VersionRequired =20
Begin Report
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    AllowUpdating =4
    ScrollBars =0
    ViewsAllowed =2
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =4940
    RowHeight =300
    DatasheetFontWeight =700
    ItemSuffix =8
    Left =450
    Top =3045
    DatasheetGridlinesColor =12632256
    MaxRecords =100000
    RecSrcDt = Begin
        0x1a48b413d026e440
    End
    RecordSource ="SELECT OPANBACCESS.*, TCOD_ANBAR.NAMES, OPANBACCESS.ANBCO AS Expr1 FROM OPANBACC"
        "ESS INNER JOIN TCOD_ANBAR ON OPANBACCESS.ANBCO = TCOD_ANBAR.CODE"
    Caption ="OPANBACCESS"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Tahoma"
    OnError ="[Event Procedure]"
    AllowFormView =0
    RecordSourceQualifier ="dbo"
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
            Width =1701
            LabelX =-1701
            FontName ="Tahoma"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =750
            Name ="ReportHeader"
            Begin
                Begin Label
                    TextFontCharSet =178
                    TextAlign =2
                    TextFontFamily =2
                    Left =1303
                    Width =2460
                    Height =480
                    FontSize =12
                    FontWeight =700
                    ForeColor =128
                    Name ="Label4"
                    Caption ="ليست انبارهاي قابل دسترس"
                    FontName ="Titr"
                End
                Begin TextBox
                    Visible = NotDefault
                    ColumnHidden = NotDefault
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =3061
                    ColumnWidth =1417
                    Name ="USERCO"
                    ControlSource ="USERCO"
                    Begin
                        Begin Label
                            Visible = NotDefault
                            Left =1247
                            Width =690
                            Height =240
                            Name ="Label1"
                            Caption ="كاربر"
                        End
                    End
                End
                Begin Label
                    Left =1410
                    Top =510
                    Width =615
                    Height =240
                    Name ="Label3"
                    Caption =" نام انبار"
                End
                Begin Label
                    Left =3636
                    Top =510
                    Width =615
                    Height =240
                    Name ="Label7"
                    Caption =" كد انبار"
                End
            End
        End
        Begin Section
            Height =240
            Name ="Detail"
            Begin
                Begin TextBox
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =3690
                    Width =576
                    ColumnWidth =3615
                    Name ="ANBCO"
                    ControlSource ="ANBCO"
                End
                Begin TextBox
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =58
                    Width =3576
                    TabIndex =1
                    Name ="Text5"
                    ControlSource ="NAMES"
                End
                Begin TextBox
                    RunningSum =2
                    TextAlign =2
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =4305
                    Width =576
                    TabIndex =2
                    Name ="Text6"
                    ControlSource ="=1"
                End
            End
        End
        Begin FormFooter
            KeepTogether = NotDefault
            Height =0
            Name ="ReportFooter"
        End
    End
End
CodeBehindForm
' See "OPANBACCESS_SUB_R.cls"
