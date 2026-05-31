Version =20
VersionRequired =20
Begin Form
    AllowDeletions = NotDefault
    AllowAdditions = NotDefault
    OrderByOn = NotDefault
    AllowEdits = NotDefault
    DefaultView =2
    AllowUpdating =4
    ViewsAllowed =2
    TabularCharSet =178
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =6773
    RowHeight =480
    ItemSuffix =15
    Left =1200
    Top =4770
    Right =6030
    Bottom =6930
    DatasheetForeColor =33554432
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x6a596e8af2c5e340
    End
    RecordSource ="EVENTS"
    Caption ="EVENTS"
    DatasheetFontName ="Tahoma"
    OnError ="[Event Procedure]"
    AllowFormView =0
    RecordSourceQualifier ="dbo"
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
        End
        Begin BoundObjectFrame
            SpecialEffect =2
            Width =4536
            Height =2835
            LabelX =-1701
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
        Begin Section
            Height =5046
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =93
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =4139
                    Top =113
                    Name ="IDNUM"
                    ControlSource ="IDNUM"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =2325
                            Top =113
                            Width =570
                            Height =240
                            Name ="Label1"
                            Caption ="شماره"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =4139
                    Top =453
                    ColumnWidth =660
                    TabIndex =1
                    Name ="IDD"
                    ControlSource ="IDD"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2332
                            Top =450
                            Width =405
                            Height =240
                            Name ="Label3"
                            Caption ="رديف"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =4139
                    Top =793
                    Width =2490
                    ColumnWidth =9120
                    TabIndex =2
                    Name ="EVENTS"
                    ControlSource ="EVENTS"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2325
                            Top =793
                            Width =630
                            Height =240
                            Name ="Label5"
                            Caption ="روي داد"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =4139
                    Top =1814
                    Width =2490
                    TabIndex =3
                    Name ="USERNAME"
                    ControlSource ="USERNAME"
                    DefaultValue ="UCurrentUser()"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2325
                            Top =1814
                            Width =885
                            Height =240
                            Name ="Label11"
                            Caption ="كاربر"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =4359
                    Top =1077
                    Width =471
                    ColumnWidth =945
                    TabIndex =4
                    Name ="STDATE"
                    ControlSource ="STDATE"
                    DefaultValue ="=FARSIDATE(Date())"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2552
                            Top =1083
                            Width =840
                            Height =240
                            Name ="Label13"
                            Caption ="تاريخ  ارجاع"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =4359
                    Top =1417
                    Width =471
                    ColumnWidth =885
                    TabIndex =5
                    Name ="STTIME"
                    ControlSource ="STTIME"
                    DefaultValue ="=Hour(Now())*100+Minute(Now())"
                    InputMask ="##:##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2552
                            Top =1413
                            Width =780
                            Height =240
                            Name ="Label15"
                            Caption ="زمان ارجاع"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =215
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2212
                    ColumnWidth =870
                    TabIndex =6
                    Name ="SUMTIME"
                    ControlSource ="SUMTIME"
                    InputMask ="##:##"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =4804
                            Width =810
                            Height =240
                            Name ="Label12"
                            Caption ="ساعت كار"
                        End
                    End
                End
                Begin BoundObjectFrame
                    OverlapFlags =85
                    Top =2211
                    TabIndex =7
                    Name ="pic"
                    ControlSource ="pic"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5910
                            Top =2205
                            Width =525
                            Height =240
                            Name ="Label14"
                            Caption ="ضمائم"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "EVENTS_COMPANY.cls"
