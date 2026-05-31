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
    Width =10150
    DatasheetFontHeight =10
    ItemSuffix =53
    DatasheetGridlinesColor =12632256
    MaxRecords =100000
    RecSrcDt = Begin
        0x35072fe45ee5e240
    End
    RecordSource ="taraznamah"
    Caption ="ترازنامه"
    DatasheetFontName ="Arial"
    RecordSourceQualifier ="dbo"
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
            GroupHeader = NotDefault
            GroupFooter = NotDefault
            KeepTogether =1
            GroupOn =9
            ControlSource ="grp"
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
        Begin PageHeader
            Height =675
            Name ="PageHeaderSection"
            Begin
                Begin Label
                    FontItalic = NotDefault
                    BackStyle =1
                    TextFontCharSet =178
                    TextAlign =2
                    TextFontFamily =2
                    Left =4862
                    Width =1140
                    Height =675
                    FontSize =20
                    Name ="Label24"
                    Caption ="ترازنامه"
                    FontName ="Zar"
                End
            End
        End
        Begin BreakHeader
            Height =470
            Name ="GroupHeader0"
            Begin
                Begin TextBox
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =3
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =7654
                    Width =2337
                    Height =420
                    FontSize =12
                    FontWeight =700
                    BackColor =12632256
                    Name ="CODE1"
                    ControlSource ="=IIf([grp]=1,'دارائيها',Null)"
                    FontName ="Zar"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =3
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =3174
                    Width =2337
                    Height =420
                    FontSize =12
                    FontWeight =700
                    TabIndex =1
                    BackColor =12632256
                    Name ="Text30"
                    ControlSource ="=IIf([grp]<>1,'بدهيها و سرمايه',Null)"
                    FontName ="Zar"
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =405
            BreakLevel =2
            Name ="GroupHeader2"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =3
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =6973
                    Width =1995
                    Height =405
                    FontSize =11
                    FontWeight =700
                    ForeColor =8388608
                    Name ="NAMES11"
                    ControlSource ="=IIf([grp]<>1,Null,[NAMES])"
                    FontName ="Zar"
                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =3
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =2551
                    Width =1995
                    Height =405
                    FontSize =11
                    FontWeight =700
                    TabIndex =1
                    ForeColor =8388608
                    Name ="Text39"
                    ControlSource ="=IIf([grp]<>1,[NAMES],Null)"
                    FontName ="Zar"
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =300
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
                    Left =7369
                    Width =630
                    Height =300
                    FontWeight =700
                    Name ="NUMBER22"
                    ControlSource ="=IIf([grp]=1,[NUMBER],Null)"
                    Format ="#,###"
                    FontName ="Zar"
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
                    Left =5278
                    Width =2040
                    Height =300
                    FontWeight =700
                    TabIndex =1
                    Name ="NAME22"
                    ControlSource ="=IIf([grp]=1,[NAM],Null)"
                    Format ="#,###"
                    FontName ="Zar"
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
                    Left =3798
                    Width =1488
                    Height =300
                    FontWeight =700
                    TabIndex =2
                    Name ="daraee22"
                    ControlSource ="=IIf([grp]=1,[daraee],Null)"
                    Format ="#,###"
                    FontName ="Zar"
                End
                Begin TextBox
                    TextFontCharSet =178
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =3571
                    Width =630
                    Height =300
                    FontWeight =700
                    TabIndex =3
                    Name ="Text43"
                    ControlSource ="=IIf([grp]<>1,[NUMBER],Null)"
                    Format ="#,###"
                    FontName ="Zar"
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
                    Left =1480
                    Width =2040
                    Height =300
                    FontWeight =700
                    TabIndex =4
                    Name ="Text44"
                    ControlSource ="=IIf([grp]<>1,[NAM],Null)"
                    Format ="#,###"
                    FontName ="Zar"
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
                    Width =1488
                    Height =300
                    FontWeight =700
                    TabIndex =5
                    Name ="Text45"
                    ControlSource ="=IIf([grp]<>1,[bedehe],Null)"
                    Format ="#,###"
                    FontName ="Zar"
                End
                Begin TextBox
                    Visible = NotDefault
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    FontWeight =700
                    TabIndex =6
                    Name ="CODE"
                    ControlSource ="CODE"
                    Format ="#,###"
                End
                Begin TextBox
                    Visible = NotDefault
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    FontWeight =700
                    TabIndex =7
                    Name ="NAMES"
                    ControlSource ="NAMES"
                    Format ="#,###"
                End
                Begin TextBox
                    Visible = NotDefault
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    FontWeight =700
                    TabIndex =8
                    Name ="NUMBER"
                    ControlSource ="NUMBER"
                    Format ="#,###"
                End
                Begin TextBox
                    Visible = NotDefault
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    FontWeight =700
                    TabIndex =9
                    Name ="NAM"
                    ControlSource ="NAM"
                    Format ="#,###"
                End
                Begin TextBox
                    Visible = NotDefault
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    FontWeight =700
                    TabIndex =10
                    Name ="daraee"
                    ControlSource ="daraee"
                    Format ="#,###"
                End
                Begin TextBox
                    Visible = NotDefault
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    FontWeight =700
                    TabIndex =11
                    Name ="bedehe"
                    ControlSource ="bedehe"
                    Format ="#,###"
                End
                Begin TextBox
                    Visible = NotDefault
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Height =225
                    FontWeight =700
                    TabIndex =12
                    Name ="grp"
                    ControlSource ="grp"
                    Format ="#,###"
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =480
            BreakLevel =2
            OnFormat ="[Event Procedure]"
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
                    Left =3405
                    Top =120
                    Width =1863
                    Height =360
                    FontSize =10
                    FontWeight =700
                    Name ="Sum Of daraee"
                    ControlSource ="=IIf([grp]=1,Sum([daraee]),Null)"
                    Format ="#,###"
                    FontName ="Zar"
                    EventProcPrefix ="Sum_Of_daraee"
                End
                Begin CommandButton
                    TextFontCharSet =178
                    TextFontFamily =2
                    Left =3401
                    Top =56
                    Width =1924
                    Height =68
                    TabIndex =1
                    Name ="Command33"
                    Caption ="Command33"
                    FontName ="Zar"
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
                    Width =1698
                    Height =360
                    FontSize =10
                    FontWeight =700
                    TabIndex =2
                    Name ="Text36"
                    ControlSource ="=IIf([grp]<>1,Sum([bedehe]),Null)"
                    Format ="#,###"
                    FontName ="Zar"
                End
                Begin CommandButton
                    TextFontCharSet =178
                    TextFontFamily =2
                    Top =56
                    Width =1819
                    Height =68
                    TabIndex =3
                    Name ="Command37"
                    Caption ="Command33"
                    FontName ="Zar"
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =645
            Name ="GroupFooter1"
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
                    Left =7080
                    Top =135
                    Width =3003
                    Height =510
                    FontSize =14
                    FontWeight =700
                    Name ="Sum Of daraee1"
                    ControlSource ="=Sum([daraee])"
                    Format ="#,###"
                    FontName ="Zar"
                    EventProcPrefix ="Sum_Of_daraee1"
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
                    Left =2324
                    Top =113
                    Width =3062
                    Height =510
                    FontSize =14
                    FontWeight =700
                    TabIndex =1
                    Name ="Sum Of bedehe1"
                    ControlSource ="=Sum([bedehe])"
                    Format ="#,###"
                    FontName ="Zar"
                    EventProcPrefix ="Sum_Of_bedehe1"
                End
                Begin CommandButton
                    TextFontCharSet =178
                    TextFontFamily =2
                    Left =7029
                    Width =3098
                    Height =113
                    TabIndex =2
                    Name ="Command34"
                    Caption ="Command34"
                    FontName ="Zar"
                End
                Begin CommandButton
                    TextFontCharSet =178
                    TextFontFamily =2
                    Left =2324
                    Width =3098
                    Height =113
                    TabIndex =3
                    Name ="Command38"
                    Caption ="Command34"
                    FontName ="Zar"
                End
            End
        End
        Begin PageFooter
            Height =0
            Name ="PageFooterSection"
        End
    End
End
CodeBehindForm
' See "TARAZNAMAH.cls"
