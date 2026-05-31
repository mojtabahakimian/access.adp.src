Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    DefaultView =0
    AllowUpdating =4
    ScrollBars =0
    ViewsAllowed =1
    TabularCharSet =178
    TabularFamily =0
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    Cycle =1
    GridY =10
    Width =9720
    DatasheetFontHeight =10
    ItemSuffix =24
    Left =10635
    Top =14925
    Right =20520
    Bottom =17625
    DatasheetGridlinesColor =12632256
    MaxRecords =1000000
    RecSrcDt = Begin
        0xd7b32174b11ee440
    End
    RecordSource ="OTHER_DTL"
    Caption ="ساير اطلاعات"
    BeforeInsert ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Arial"
    OnError ="[Event Procedure]"
    AllowDatasheetView =0
    RecordSourceQualifier ="dbo"
    Begin
        Begin Label
            BackStyle =0
            TextFontCharSet =178
            FontName ="Tahoma"
        End
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            OldBorderStyle =0
            TextFontCharSet =178
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
            Height =2437
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =93
                    TextAlign =3
                    IMESentenceMode =3
                    Left =120
                    Top =1080
                    Width =51
                    Name ="NUMBER"
                    ControlSource ="NUMBER"
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =93
                    TextAlign =3
                    IMESentenceMode =3
                    Left =120
                    Top =1420
                    Width =51
                    TabIndex =1
                    Name ="TAG"
                    ControlSource ="TAG"
                    DefaultValue ="2"
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6690
                    Top =114
                    Width =1215
                    Height =345
                    TabIndex =2
                    BackColor =13434828
                    Name ="REQUEST_NO"
                    ControlSource ="REQUEST_NO"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =7998
                            Top =113
                            Width =1575
                            Height =240
                            Name ="Label5"
                            Caption ="شماره درخواست كالا:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =3401
                    Top =162
                    Width =1980
                    Height =345
                    TabIndex =3
                    BackColor =13434828
                    Name ="BARNAMEH"
                    ControlSource ="BARNAMEH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5506
                            Top =159
                            Width =1050
                            Height =240
                            Name ="Label7"
                            Caption ="شماره بارنامه:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6572
                    Top =690
                    Width =2025
                    Height =345
                    TabIndex =5
                    Name ="DRIVER_MOB"
                    ControlSource ="DRIVER_MOB"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =8655
                            Top =745
                            Width =975
                            Height =240
                            Name ="Label11"
                            Caption ="موبايل راننده:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =3056
                    Top =1260
                    Height =345
                    TabIndex =9
                    Name ="CAM_KHALY"
                    ControlSource ="CAM_KHALY"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4870
                            Top =1260
                            Width =1245
                            Height =240
                            Name ="Label15"
                            Caption ="وزن ماشين خالي"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =247
                    TextAlign =3
                    IMESentenceMode =3
                    Left =113
                    Top =1260
                    Height =345
                    TabIndex =10
                    Name ="CAM_POOR"
                    ControlSource ="CAM_POOR"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1979
                            Top =1260
                            Width =1005
                            Height =240
                            Name ="Label17"
                            Caption ="وزن ماشين پر"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =113
                    Top =1793
                    Width =8835
                    Height =345
                    TabIndex =11
                    Name ="TOZIH"
                    ControlSource ="TOZIH"
                    DefaultValue ="\"اقلام فوق صحيح و سالم تحويل اينجانب گرديد\""
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =9039
                            Top =1793
                            Width =675
                            Height =240
                            Name ="Label21"
                            Caption ="توضيحات"
                        End
                    End
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =6240
                    Top =1245
                    Width =2601
                    Height =315
                    TabIndex =8
                    Name ="MAGHSAD"
                    ControlSource ="MAGHSAD"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT MPCODE, MPNAME FROM TCOD_MAP WHERE (MPP = 100) ORDER BY MPNAME"
                    ColumnWidths ="0"
                    OnDblClick ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =8895
                            Top =1290
                            Width =825
                            Height =240
                            Name ="Label19"
                            Caption ="*مقصد بار:"
                        End
                    End
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =113
                    Top =159
                    Width =2370
                    Height =345
                    TabIndex =4
                    Name ="DRIVER"
                    ControlSource ="DRIVER"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TOP 100 PERCENT DRIVER FROM dbo.OTHER_DTL GROUP BY DRIVER ORDER BY DRIVER"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2610
                            Top =158
                            Width =750
                            Height =240
                            Name ="Label9"
                            Caption ="نام راننده:"
                        End
                    End
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =3075
                    Top =707
                    Width =2265
                    Height =300
                    TabIndex =6
                    Name ="CAMIUN_NUM"
                    ControlSource ="CAMIUN_NUM"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT     TOP 100 PERCENT DRIVER_MOB\015\012FROM         dbo.OTHER_DTL\015\012G"
                        "ROUP BY DRIVER_MOB\015\012ORDER BY DRIVER_MOB"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5403
                            Top =717
                            Width =1110
                            Height =285
                            Name ="Label13"
                            Caption ="شماره ماشين:"
                        End
                    End
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =113
                    Top =729
                    Width =1965
                    Height =300
                    TabIndex =7
                    Name ="CAMIUN"
                    ControlSource ="CAMIUN"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TOP 100 PERCENT CAMIUN FROM OTHER_DTL GROUP BY CAMIUN ORDER BY CAMIUN"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2133
                            Top =739
                            Width =885
                            Height =240
                            Name ="Label23"
                            Caption ="نوع ماشين:"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "OTHER_DTL_SUB1.cls"
