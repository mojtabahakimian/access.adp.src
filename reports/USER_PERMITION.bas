Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    DefaultView =0
    AllowUpdating =4
    TabularCharSet =178
    TabularFamily =42
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =10180
    DatasheetFontHeight =10
    ItemSuffix =99
    Left =270
    Top =600
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x8b76c7edcf26e440
    End
    RecordSource ="SELECT TOP 100 PERCENT TFORMS.CAPTION, SAL_CHEK.OBJECT, SAL_CHEK.RUN, SAL_CHEK.S"
        "EE, SAL_CHEK.INP, SAL_CHEK.UPD, SAL_CHEK.DEL, SAL_CHEK.USERCO, TFORMS.GRP FROM T"
        "FORMS INNER JOIN SAL_CHEK ON TFORMS.IDH = SAL_CHEK.OBJECT WHERE (SAL_CHEK.RUN = "
        "1) ORDER BY TFORMS.GRP, TFORMS.CAPTION"
    Caption ="دسترسي"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Arial (Arabic)"
    RecordSourceQualifier ="dbo"
    Begin
        Begin Label
            BackStyle =0
            TextFontCharSet =178
            TextAlign =1
            TextFontFamily =42
            FontSize =9
            FontWeight =700
            ForeColor =128
            FontName ="Arial (Arabic)"
        End
        Begin Rectangle
            BackStyle =0
            BorderWidth =1
            Width =850
            Height =850
        End
        Begin Line
            Width =1701
            BorderColor =128
        End
        Begin Image
            OldBorderStyle =0
            PictureAlignment =2
            Width =1701
            Height =1701
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
            LabelX =230
            LabelY =-30
        End
        Begin CheckBox
            LabelX =230
            LabelY =-30
        End
        Begin OptionGroup
            BackStyle =1
            Width =1701
            Height =1701
        End
        Begin BoundObjectFrame
            Width =4536
            Height =2835
            LabelX =-1701
        End
        Begin TextBox
            OldBorderStyle =0
            TextFontCharSet =178
            TextFontFamily =42
            BackStyle =0
            Width =1701
            LabelX =-1701
            FontSize =9
            FontName ="Arial (Arabic)"
        End
        Begin ListBox
            TextFontCharSet =178
            TextFontFamily =42
            OldBorderStyle =0
            Width =1701
            Height =1417
            LabelX =-1701
            FontSize =9
            FontName ="Arial (Arabic)"
        End
        Begin ComboBox
            OldBorderStyle =0
            TextFontCharSet =178
            TextFontFamily =42
            BackStyle =0
            Width =1701
            LabelX =-1701
            FontSize =9
            FontName ="Arial (Arabic)"
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
        Begin FormHeader
            KeepTogether = NotDefault
            Height =0
            Name ="ReportHeader"
        End
        Begin PageHeader
            Height =1259
            Name ="PageHeader"
            Begin
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =9583
                    Top =867
                    Width =555
                    Height =315
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="N_S Label"
                    Caption ="رديف"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="N_S_Label"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =4980
                    Top =870
                    Width =4576
                    Height =315
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="SHARH Label"
                    Caption ="شرح"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="SHARH_Label"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =4020
                    Top =873
                    Width =921
                    Height =315
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="BED Label"
                    Caption ="اجراء"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="BED_Label"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =3034
                    Top =870
                    Width =920
                    Height =315
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="BES Label"
                    Caption ="مشاهده"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="BES_Label"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =2040
                    Top =870
                    Width =920
                    Height =315
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="Expr1 Label"
                    Caption ="ورود"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="Expr1_Label"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =1035
                    Top =873
                    Width =920
                    Height =315
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="MAND Label"
                    Caption ="به هنگام"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="MAND_Label"
                End
                Begin Line
                    BorderWidth =2
                    Left =15
                    Top =855
                    Width =10165
                    Name ="Line26"
                End
                Begin Line
                    BorderWidth =2
                    Top =1245
                    Width =10165
                    Name ="Line27"
                End
                Begin Line
                    Left =9585
                    Top =876
                    Width =0
                    Height =374
                    Name ="Line29"
                End
                Begin Line
                    Left =4950
                    Top =876
                    Width =0
                    Height =374
                    Name ="Line32"
                End
                Begin Line
                    Left =4009
                    Top =885
                    Width =0
                    Height =374
                    Name ="Line33"
                End
                Begin Line
                    Left =2984
                    Top =870
                    Width =0
                    Height =374
                    Name ="Line34"
                End
                Begin Line
                    Left =1995
                    Top =873
                    Width =0
                    Height =374
                    Name ="Line35"
                End
                Begin Line
                    BorderWidth =2
                    Top =873
                    Width =0
                    Height =374
                    Name ="Line36"
                End
                Begin Line
                    BorderWidth =2
                    Left =10155
                    Top =873
                    Width =0
                    Height =374
                    Name ="Line37"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =45
                    Top =870
                    Width =945
                    Height =330
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="Label78"
                    Caption ="حذف"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    Left =990
                    Top =873
                    Width =0
                    Height =374
                    Name ="Line79"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =2610
                    Top =390
                    Width =5100
                    Height =465
                    FontSize =12
                    FontWeight =400
                    ForeColor =0
                    Name ="Label94"
                    Caption ="شرح"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin TextBox
                    Visible = NotDefault
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =340
                    Top =686
                    Name ="USERCO"
                    ControlSource ="USERCO"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2430
                    Width =5433
                    Height =375
                    FontSize =10
                    TabIndex =1
                    BorderColor =32768
                    Name ="Text90"
                    ControlSource ="=[Forms]![Baseknow]![WIDTH_D]"
                    FontName ="WeblogmaYekan"
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =419
            Name ="Detail"
            Begin
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4980
                    Width =4575
                    Height =375
                    ColumnWidth =4650
                    FontSize =10
                    Name ="CAPTION"
                    ControlSource ="CAPTION"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    RunningSum =1
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =9589
                    Width =561
                    Height =375
                    FontSize =10
                    TabIndex =1
                    Name ="N_S"
                    ControlSource ="=1"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    Left =9585
                    Width =0
                    Height =419
                    Name ="Line39"
                End
                Begin Line
                    Left =4950
                    Width =0
                    Height =419
                    Name ="Line42"
                End
                Begin Line
                    Left =4009
                    Width =0
                    Height =419
                    Name ="Line43"
                End
                Begin Line
                    Left =2984
                    Width =0
                    Height =419
                    Name ="Line44"
                End
                Begin Line
                    Left =1995
                    Width =0
                    Height =419
                    Name ="Line45"
                End
                Begin Line
                    BorderWidth =2
                    Width =0
                    Height =419
                    Name ="Line46"
                End
                Begin Line
                    BorderWidth =2
                    Left =10155
                    Width =0
                    Height =419
                    Name ="Line47"
                End
                Begin Line
                    LineSlant = NotDefault
                    Top =419
                    Width =10178
                    BorderColor =4210752
                    Name ="Line48"
                End
                Begin Line
                    Left =990
                    Width =0
                    Height =419
                    Name ="Line81"
                End
                Begin CheckBox
                    Left =4365
                    Top =135
                    TabIndex =2
                    Name ="RUN"
                    ControlSource ="RUN"
                End
                Begin CheckBox
                    Left =3390
                    Top =135
                    TabIndex =3
                    Name ="SEE"
                    ControlSource ="SEE"
                End
                Begin CheckBox
                    Left =2370
                    Top =135
                    TabIndex =4
                    Name ="Check91"
                    ControlSource ="SEE"
                End
                Begin CheckBox
                    Left =1410
                    Top =120
                    TabIndex =5
                    Name ="Check92"
                    ControlSource ="SEE"
                End
                Begin CheckBox
                    Left =340
                    Top =135
                    TabIndex =6
                    Name ="Check93"
                    ControlSource ="SEE"
                End
            End
        End
        Begin PageFooter
            Height =0
            Name ="PageFooter"
        End
        Begin FormFooter
            KeepTogether = NotDefault
            CanGrow = NotDefault
            Height =1077
            Name ="ReportFooter"
            Begin
                Begin Subform
                    Left =120
                    Top =120
                    Width =4935
                    Height =675
                    Name ="OPANBACCESS_SUB_R"
                    SourceObject ="Report.OPANBACCESS_SUB_R"
                    LinkChildFields ="USERCO"
                    LinkMasterFields ="USERCO"
                End
                Begin Subform
                    Left =5160
                    Top =90
                    Width =4935
                    Height =675
                    TabIndex =1
                    Name ="Child98"
                    SourceObject ="Report.BLOCK_HES_SUB_R"
                    LinkChildFields ="USERCO"
                    LinkMasterFields ="USERCO"
                End
            End
        End
    End
End
CodeBehindForm
' See "USER_PERMITION.cls"
