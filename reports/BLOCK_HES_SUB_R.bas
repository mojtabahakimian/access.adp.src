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
    Left =5490
    Top =3015
    DatasheetGridlinesColor =12632256
    MaxRecords =100000
    RecSrcDt = Begin
        0xaf0b5353d226e440
    End
    RecordSource ="SELECT USERCO, HES FROM BLOCK_HES"
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
                    TextFontFamily =2
                    Left =780
                    Width =2865
                    Height =480
                    FontSize =12
                    FontWeight =700
                    ForeColor =128
                    Name ="Label4"
                    Caption ="ليست حسابهايي كه مسدود هستند"
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
                    Left =1755
                    Top =495
                    Width =750
                    Height =240
                    Name ="Label3"
                    Caption ="كد حساب"
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
                    Left =170
                    Width =4086
                    Name ="HES"
                    ControlSource ="HES"
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
                    TabIndex =1
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
' See "BLOCK_HES_SUB_R.cls"
