Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    OrderByOn = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    AllowUpdating =4
    TabularCharSet =178
    TabularFamily =11
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =5272
    DatasheetFontHeight =10
    ItemSuffix =58
    Left =5595
    Top =2670
    DatasheetGridlinesColor =12632256
    MaxRecords =100000
    RecSrcDt = Begin
        0xcf265fc932d5e440
    End
    RecordSource ="SELECT NUMBER, NAM, CODE, NAMES, daraee, grp, GR, DARAEECH, daraeeE, BEDEHECH, D"
        "ARAEECH + BEDEHECH * - 1 AS change FROM SURATMANABE(5, 950105, 999999, 960101) S"
        "URATMANABE WHERE (GR = 1) AND (DARAEECH >= 0) OR (GR = 2) AND (BEDEHECH < 0)"
    Caption ="ترازنامه"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial"
    Begin
        Begin Label
            FontItalic = NotDefault
            BackStyle =0
            TextAlign =1
            TextFontFamily =18
            FontSize =11
            FontWeight =700
            ForeColor =8388608
            FontName ="Times New Roman"
        End
        Begin Rectangle
            BackStyle =0
            BorderWidth =1
            Width =850
            Height =850
            BorderColor =8388608
        End
        Begin Line
            Width =1701
            BorderColor =8388608
        End
        Begin Image
            OldBorderStyle =0
            PictureAlignment =2
            Width =1701
            Height =1701
        End
        Begin CommandButton
            Width =1701
            Height =283
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Tahoma"
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
            FontName ="Arial"
            AsianLineBreak =255
        End
        Begin ListBox
            OldBorderStyle =0
            Width =1701
            Height =1417
            LabelX =-1701
            FontName ="Arial"
        End
        Begin ComboBox
            OldBorderStyle =0
            BackStyle =0
            Width =1701
            LabelX =-1701
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
            ControlSource ="CODE"
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            GroupFooter = NotDefault
            ControlSource ="NAMES"
        End
        Begin BreakLevel
            ControlSource ="NUMBER"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Visible = NotDefault
            Height =0
            Name ="ReportHeader"
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =405
            BreakLevel =1
            Name ="GroupHeader2"
            Begin
                Begin TextBox
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =3
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =3165
                    Width =1995
                    Height =375
                    ForeColor =10040115
                    Name ="NAMES11"
                    ControlSource ="NAMES"
                    FontName ="WeblogmaYekan"
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =567
            Name ="Detail"
            Begin
                Begin TextBox
                    TextFontCharSet =178
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =4176
                    Top =120
                    Width =630
                    Height =375
                    ForeColor =10040115
                    Name ="NUMBER"
                    ControlSource ="NUMBER"
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2085
                    Top =120
                    Width =2040
                    Height =375
                    TabIndex =1
                    ForeColor =10040115
                    Name ="NAME22"
                    ControlSource ="NAM"
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Top =120
                    Width =2043
                    Height =375
                    TabIndex =2
                    ForeColor =10040115
                    Name ="change"
                    ControlSource ="change"
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    TextFontCharSet =178
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =4595
                    Top =120
                    Width =396
                    Height =375
                    TabIndex =3
                    ForeColor =10040115
                    Name ="GRP"
                    ControlSource ="GRP"
                    FontName ="WeblogmaYekan"
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =623
            BreakLevel =1
            Name ="GroupFooter3"
            Begin
                Begin TextBox
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =4
                    Top =120
                    Width =1863
                    Height =375
                    ForeColor =10040115
                    Name ="Sum Of daraee"
                    ControlSource ="=Sum([change])"
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                    EventProcPrefix ="Sum_Of_daraee"
                End
                Begin CommandButton
                    TextFontCharSet =178
                    TextFontFamily =2
                    Width =2268
                    Height =83
                    TabIndex =1
                    ForeColor =128
                    Name ="Command33"
                    Caption ="Command33"
                    FontName ="WeblogmaYekan"
                End
            End
        End
        Begin FormFooter
            KeepTogether = NotDefault
            Visible = NotDefault
            Height =396
            Name ="ReportFooter"
            Begin
                Begin TextBox
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Width =1818
                    Height =375
                    ForeColor =10040115
                    Name ="SDARAEE"
                    ControlSource ="=Sum([change])"
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                End
            End
        End
    End
End
CodeBehindForm
' See "SURATMANABEVA22.cls"
