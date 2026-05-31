Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
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
    Width =8997
    DatasheetFontHeight =10
    ItemSuffix =26
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xff50081d9cace240
    End
    RecordSource ="LIST_SALARY5"
    Caption ="LIST_SALARY5"
    DatasheetFontName =""
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
        Begin FormHeader
            KeepTogether = NotDefault
            Height =909
            Name ="ReportHeader"
            Begin
                Begin Label
                    BackStyle =1
                    OverlapFlags =4
                    TextAlign =0
                    Left =57
                    Top =57
                    Width =2865
                    Height =510
                    FontSize =20
                    Name ="Label22"
                    Caption ="LIST_SALARY5"
                End
            End
        End
        Begin PageHeader
            Height =0
            Name ="PageHeaderSection"
        End
        Begin Section
            KeepTogether = NotDefault
            Height =4434
            Name ="Detail"
            Begin
                Begin TextBox
                    OldBorderStyle =1
                    IMESentenceMode =3
                    Left =1977
                    Top =57
                    Width =1185
                    Height =270
                    Name ="CODE"
                    ControlSource ="CODE"
                    Begin
                        Begin Label
                            Left =57
                            Top =57
                            Width =1860
                            Height =330
                            Name ="CODE_Label"
                            Caption ="CODE"
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =1
                    IMESentenceMode =3
                    Left =1977
                    Top =456
                    Width =4560
                    Height =270
                    TabIndex =1
                    Name ="PRF"
                    ControlSource ="PRF"
                    Begin
                        Begin Label
                            Left =57
                            Top =456
                            Width =1860
                            Height =330
                            Name ="PRF_Label"
                            Caption ="PRF"
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =1
                    IMESentenceMode =3
                    Left =1977
                    Top =855
                    Width =2310
                    Height =270
                    TabIndex =2
                    Name ="PFAMILY"
                    ControlSource ="PFAMILY"
                    Begin
                        Begin Label
                            Left =57
                            Top =855
                            Width =1860
                            Height =330
                            Name ="PFAMILY_Label"
                            Caption ="PFAMILY"
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =1
                    IMESentenceMode =3
                    Left =1977
                    Top =1254
                    Width =2310
                    Height =270
                    TabIndex =3
                    Name ="KHNOWNUM"
                    ControlSource ="KHNOWNUM"
                    Begin
                        Begin Label
                            Left =57
                            Top =1254
                            Width =1860
                            Height =330
                            Name ="KHNOWNUM_Label"
                            Caption ="KHNOWNUM"
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =1
                    IMESentenceMode =3
                    Left =1977
                    Top =1653
                    Width =2310
                    Height =270
                    TabIndex =4
                    Name ="FATHER"
                    ControlSource ="FATHER"
                    Begin
                        Begin Label
                            Left =57
                            Top =1653
                            Width =1860
                            Height =330
                            Name ="FATHER_Label"
                            Caption ="FATHER"
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =1
                    IMESentenceMode =3
                    Left =1977
                    Top =2052
                    Width =1185
                    Height =270
                    TabIndex =5
                    Name ="MM"
                    ControlSource ="MM"
                    Begin
                        Begin Label
                            Left =57
                            Top =2052
                            Width =1860
                            Height =330
                            Name ="MM_Label"
                            Caption ="MM"
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =1
                    IMESentenceMode =3
                    Left =1977
                    Top =2451
                    Width =1035
                    Height =270
                    TabIndex =6
                    Name ="WDATE"
                    ControlSource ="WDATE"
                    Begin
                        Begin Label
                            Left =57
                            Top =2451
                            Width =1860
                            Height =330
                            Name ="WDATE_Label"
                            Caption ="WDATE"
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =1
                    IMESentenceMode =3
                    Left =1977
                    Top =2850
                    Width =1185
                    Height =270
                    TabIndex =7
                    Name ="DAYS"
                    ControlSource ="DAYS"
                    Begin
                        Begin Label
                            Left =57
                            Top =2850
                            Width =1860
                            Height =330
                            Name ="DAYS_Label"
                            Caption ="DAYS"
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =1
                    IMESentenceMode =3
                    Left =1977
                    Top =3249
                    Width =2310
                    Height =270
                    TabIndex =8
                    Name ="BEDH"
                    ControlSource ="BEDH"
                    Begin
                        Begin Label
                            Left =57
                            Top =3249
                            Width =1860
                            Height =330
                            Name ="BEDH_Label"
                            Caption ="BEDH"
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =1
                    IMESentenceMode =3
                    Left =1977
                    Top =3648
                    Width =2310
                    Height =270
                    TabIndex =9
                    Name ="VMAN"
                    ControlSource ="VMAN"
                    Begin
                        Begin Label
                            Left =57
                            Top =3648
                            Width =1860
                            Height =330
                            Name ="VMAN_Label"
                            Caption ="VMAN"
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =1
                    IMESentenceMode =3
                    Left =1977
                    Top =4047
                    Width =1185
                    Height =270
                    TabIndex =10
                    Name ="HES_T"
                    ControlSource ="HES_T"
                    Begin
                        Begin Label
                            Left =57
                            Top =4047
                            Width =1860
                            Height =330
                            Name ="HES_T_Label"
                            Caption ="HES_T"
                        End
                    End
                End
                Begin Line
                    Left =57
                    Top =57
                    Width =6480
                    Name ="Line25"
                End
            End
        End
        Begin PageFooter
            Height =528
            Name ="PageFooterSection"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    OverlapFlags =4
                    TextAlign =1
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =57
                    Top =228
                    Width =4384
                    Height =300
                    FontSize =9
                    FontWeight =700
                    ForeColor =8388608
                    Name ="Text23"
                    ControlSource ="=Now()"
                    Format ="Long Date"
                    FontName ="Times New Roman"
                End
                Begin TextBox
                    FontItalic = NotDefault
                    OverlapFlags =4
                    TextAlign =3
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =4556
                    Top =228
                    Width =4384
                    Height =300
                    FontSize =9
                    FontWeight =700
                    TabIndex =1
                    ForeColor =8388608
                    Name ="Text24"
                    ControlSource ="=\"Page \" & [Page] & \" of \" & [Pages]"
                    FontName ="Times New Roman"
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
