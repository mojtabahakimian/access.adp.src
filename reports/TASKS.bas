Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    FilterOn = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    AllowUpdating =4
    TabularCharSet =178
    TabularFamily =0
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =10605
    DatasheetFontHeight =10
    ItemSuffix =67
    Left =270
    Top =225
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x79b79ff9ee25e440
    End
    RecordSource ="SELECT     dbo.TASKS.IDNUM AS TASKS_IDNUM, dbo.TASKS.GR, dbo.TASKS.PERSONEL, dbo"
        ".TASKS.TASK, dbo.TASKS.PERIORITY, dbo.TASKS.STATUS,                        dbo.T"
        "ASKS.STDATE AS TASKS_STDATE, dbo.TASKS.STTIME AS TASKS_STTIME, dbo.TASKS.ENDATE,"
        " dbo.TASKS.ENTIME,                        dbo.TASKS.USERNAME AS TASKS_USERNAME, "
        "dbo.TASKS.COMP_COD, dbo.TASKS.SUMTIME AS TASKS_SUMTIME, dbo.EVENTS.IDNUM AS EVEN"
        "TS_IDNUM,                        dbo.EVENTS.IDD, dbo.EVENTS.EVENTS, dbo.EVENTS.S"
        "TDATE AS EVENTS_STDATE, dbo.EVENTS.STTIME AS EVENTS_STTIME,                     "
        "   dbo.EVENTS.USERNAME AS EVENTS_USERNAME, dbo.EVENTS.COMPANY, dbo.EVENTS.SUMTIM"
        "E AS EVENTS_SUMTIME, dbo.TASKS.pic FROM         dbo.TASKS INNER JOIN            "
        "           dbo.EVENTS ON dbo.TASKS.IDNUM = dbo.EVENTS.IDNUM"
    Caption ="گزارش خدمات ارائه شده"
    DatasheetFontName ="Arial"
    Begin
        Begin Label
            BackStyle =0
            TextAlign =1
            FontSize =12
            FontName ="Haettenschweiler"
        End
        Begin Rectangle
            BackStyle =0
            BorderWidth =1
            Width =850
            Height =850
        End
        Begin Line
            Width =1701
        End
        Begin Image
            OldBorderStyle =0
            PictureAlignment =2
            Width =1701
            Height =1701
        End
        Begin CheckBox
            LabelX =230
            LabelY =-30
        End
        Begin BoundObjectFrame
            Width =4536
            Height =2835
            LabelX =-1701
        End
        Begin TextBox
            FELineBreak = NotDefault
            OldBorderStyle =0
            BackStyle =0
            Width =1701
            LabelX =-1701
            FontSize =9
            FontName ="Arial"
            AsianLineBreak =255
        End
        Begin ListBox
            Width =1701
            Height =1417
            LabelX =-1701
            FontSize =9
            FontName ="Arial"
        End
        Begin ComboBox
            OldBorderStyle =0
            BackStyle =0
            Width =1701
            LabelX =-1701
            FontSize =9
            FontName ="Arial"
        End
        Begin Subform
            OldBorderStyle =0
            Width =1701
            Height =1701
        End
        Begin UnboundObjectFrame
            Width =4536
            Height =2835
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            ControlSource ="TASKS_IDNUM"
        End
        Begin BreakLevel
            ControlSource ="EVENTS_STDATE"
        End
        Begin BreakLevel
            ControlSource ="EVENTS_STTIME"
        End
        Begin PageHeader
            Height =1077
            BackColor =16777164
            Name ="PageHeaderSection"
            Begin
                Begin Label
                    TextFontCharSet =178
                    TextAlign =0
                    TextFontFamily =2
                    Left =56
                    Width =2220
                    Height =465
                    FontSize =8
                    Name ="Label42"
                    Caption ="گزارش خدمات ارائه شده"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextFontCharSet =178
                    TextAlign =3
                    TextFontFamily =2
                    Left =9141
                    Top =687
                    Width =570
                    Height =390
                    FontSize =8
                    Name ="TASKS_IDNUM_Label"
                    Caption ="شماره"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextFontCharSet =178
                    TextAlign =3
                    TextFontFamily =2
                    Left =8064
                    Top =687
                    Width =525
                    Height =390
                    FontSize =8
                    Name ="TASKS_STDATE_Label"
                    Caption ="تاريخ"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextFontCharSet =178
                    TextAlign =3
                    TextFontFamily =2
                    Left =5109
                    Top =687
                    Width =510
                    Height =390
                    FontSize =8
                    Name ="TASK_Label"
                    Caption ="شرح"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextFontCharSet =178
                    TextAlign =3
                    TextFontFamily =2
                    Left =120
                    Top =687
                    Width =930
                    Height =390
                    FontSize =8
                    Name ="ENDATE_Label"
                    Caption ="تاريخ انجام"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextFontCharSet =178
                    TextAlign =2
                    TextFontFamily =2
                    Left =9883
                    Top =687
                    Width =570
                    Height =390
                    FontSize =8
                    Name ="Label55"
                    Caption ="رديف"
                    FontName ="WeblogmaYekan"
                End
                Begin ComboBox
                    TextFontCharSet =178
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =6459
                    Top =56
                    Width =3519
                    Height =390
                    FontSize =8
                    Name ="COMP_COD1"
                    ControlSource ="COMP_COD"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT hes, NAME FROM CUST_HESAB"
                    ColumnWidths ="0"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextFontCharSet =178
                            TextFontFamily =2
                            Left =9994
                            Top =56
                            Width =585
                            Height =390
                            FontSize =8
                            Name ="COMP_COD_Label"
                            Caption ="شركت:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    TextFontCharSet =178
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =4535
                    Top =283
                    FontSize =8
                    TabIndex =1
                    Name ="COMP_COD"
                    ControlSource ="COMP_COD"
                    FontName ="WeblogmaYekan"
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            CanGrow = NotDefault
            Height =390
            BackColor =8454143
            Name ="GroupHeader0"
            Begin
                Begin TextBox
                    TextFontCharSet =178
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =9023
                    Width =741
                    Height =390
                    FontSize =8
                    Name ="TASKS_IDNUM"
                    ControlSource ="TASKS_IDNUM"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TextFontCharSet =178
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1257
                    Width =6564
                    Height =390
                    ColumnWidth =4605
                    FontSize =8
                    TabIndex =1
                    Name ="TASK"
                    ControlSource ="TASK"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    TextFontCharSet =178
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =7838
                    Width =1134
                    Height =390
                    FontSize =8
                    TabIndex =2
                    Name ="TASKS_STDATE"
                    ControlSource ="TASKS_STDATE"
                    FontName ="WeblogmaYekan"
                    InputMask ="####/##/##"
                End
                Begin TextBox
                    TextFontCharSet =178
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Width =1134
                    Height =390
                    FontSize =8
                    TabIndex =3
                    Name ="ENDATE"
                    ControlSource ="ENDATE"
                    FontName ="WeblogmaYekan"
                    InputMask ="####/##/##"
                End
                Begin TextBox
                    RunningSum =2
                    TextFontCharSet =178
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =9825
                    Width =741
                    Height =390
                    FontSize =8
                    TabIndex =4
                    Name ="Text56"
                    ControlSource ="=1"
                    FontName ="WeblogmaYekan"
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            CanGrow = NotDefault
            Height =390
            Name ="Detail"
            Begin
                Begin TextBox
                    TextFontCharSet =178
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =7830
                    Width =1155
                    Height =390
                    FontSize =8
                    Name ="EVENTS_STDATE"
                    ControlSource ="EVENTS_STDATE"
                    FontName ="WeblogmaYekan"
                    InputMask ="####/##/##"
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TextFontCharSet =178
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1305
                    Width =6495
                    Height =390
                    FontSize =8
                    TabIndex =1
                    Name ="EVENTS"
                    ControlSource ="EVENTS"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    TextFontCharSet =178
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Width =1260
                    Height =390
                    FontSize =8
                    TabIndex =2
                    Name ="EVENTS_STTIME"
                    ControlSource ="EVENTS_STTIME"
                    FontName ="WeblogmaYekan"
                    InputMask ="###:##"
                End
                Begin TextBox
                    RunningSum =1
                    TextFontCharSet =178
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =9075
                    Width =741
                    Height =390
                    FontSize =8
                    TabIndex =3
                    Name ="Text57"
                    ControlSource ="=1"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    RunningSum =2
                    TextFontCharSet =178
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1644
                    Width =1260
                    Height =390
                    FontSize =8
                    TabIndex =4
                    Name ="EVENTS_SUMTIMEs"
                    ControlSource ="EVENTS_SUMTIME"
                    FontName ="WeblogmaYekan"
                End
            End
        End
        Begin PageFooter
            Height =618
            Name ="PageFooterSection"
            Begin
                Begin TextBox
                    TextFontCharSet =178
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =6122
                    Top =56
                    Width =4386
                    Height =390
                    FontSize =8
                    Name ="Text44"
                    ControlSource ="=\"صفحه \" & [Page] & \" از \" & [Pages]"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    BorderWidth =1
                    Width =10603
                    Name ="Line53"
                End
                Begin Line
                    BorderWidth =1
                    Top =28
                    Width =10603
                    Name ="Line54"
                End
                Begin TextBox
                    RunningSum =2
                    TextFontCharSet =178
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Top =120
                    Width =1260
                    Height =465
                    FontSize =8
                    TabIndex =1
                    Name ="Text62"
                    ControlSource ="=[EVENTS_SUMTIMEs]"
                    FontName ="WeblogmaYekan"
                    InputMask ="###:##"
                End
                Begin Label
                    TextFontCharSet =178
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    NumeralShapes =2
                    Left =1380
                    Top =150
                    Width =1620
                    Height =390
                    FontSize =8
                    Name ="Label63"
                    Caption ="جمع زمان صرف شده:"
                    FontName ="WeblogmaYekan"
                End
            End
        End
    End
End
CodeBehindForm
' See "TASKS.cls"
