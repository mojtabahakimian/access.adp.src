Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    DefaultView =2
    AllowUpdating =4
    ScrollBars =0
    ViewsAllowed =2
    BorderStyle =3
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =6286
    RowHeight =495
    ItemSuffix =19
    Left =1425
    Top =360
    Right =12825
    Bottom =3120
    DatasheetGridlinesColor =12632256
    MaxRecords =1000000
    RecSrcDt = Begin
        0x49e5625bf4f6e540
    End
    RecordSource ="IVO_EXTENDED"
    Caption ="پارامترها"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
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
        Begin UnboundObjectFrame
            SpecialEffect =2
            OldBorderStyle =1
            Width =4536
            Height =2835
        End
        Begin Section
            Height =4818
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1927
                    Top =113
                    Width =1710
                    ColumnWidth =1417
                    Name ="id"
                    ControlSource ="id"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =113
                            Width =195
                            Height =240
                            Name ="Label1"
                            Caption ="id"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1927
                    Top =453
                    ColumnWidth =930
                    TabIndex =1
                    Name ="FLD1"
                    ControlSource ="FLD1"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =450
                            Width =465
                            Height =240
                            Name ="Label3"
                            Caption ="چربي"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1927
                    Top =793
                    ColumnWidth =930
                    TabIndex =2
                    Name ="FLD2"
                    ControlSource ="FLD2"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =795
                            Width =855
                            Height =240
                            Name ="Label5"
                            Caption ="ماده خشک"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1927
                    Top =1133
                    ColumnWidth =930
                    TabIndex =3
                    Name ="FLD3"
                    ControlSource ="FLD3"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =1140
                            Width =510
                            Height =240
                            Name ="Label7"
                            Caption ="رطوبت"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1927
                    Top =1474
                    ColumnWidth =930
                    TabIndex =4
                    Name ="FLD4"
                    ControlSource ="FLD4"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =1474
                            Width =435
                            Height =240
                            Name ="Label9"
                            Caption ="PH"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1927
                    Top =1814
                    ColumnWidth =930
                    TabIndex =5
                    Name ="FLD5"
                    ControlSource ="FLD5"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =1814
                            Width =435
                            Height =240
                            Name ="Label11"
                            Caption ="نمک"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1927
                    Top =2161
                    ColumnWidth =930
                    TabIndex =6
                    Name ="FLD6"
                    ControlSource ="FLD6"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4841
                            Top =2154
                            Width =630
                            Height =240
                            Name ="Label12"
                            Caption ="دانسيته"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1927
                    Top =2501
                    ColumnWidth =930
                    TabIndex =7
                    Name ="FLD7"
                    ControlSource ="FLD7"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4841
                            Top =2499
                            Width =555
                            Height =240
                            Name ="Label13"
                            Caption ="پروتئين"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1927
                    Top =2842
                    ColumnWidth =930
                    TabIndex =8
                    Name ="FLD8"
                    ControlSource ="FLD8"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4834
                            Top =2842
                            Width =495
                            Height =240
                            Name ="Label14"
                            Caption ="انجماد"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1927
                    Top =3238
                    ColumnWidth =930
                    TabIndex =9
                    Name ="FLD9"
                    ControlSource ="FLD9"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4834
                            Top =3238
                            Width =495
                            Height =240
                            Name ="Label15"
                            Caption ="اسيد"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1927
                    Top =3522
                    ColumnWidth =930
                    TabIndex =10
                    Name ="FLD10"
                    ControlSource ="FLD10"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =4744
                            Top =3522
                            Width =585
                            Height =240
                            Name ="Label16"
                            Caption ="الکل"
                        End
                    End
                End
                Begin TextBox
                    FELineBreak = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3123
                    Top =4484
                    Width =786
                    TabIndex =11
                    BackColor =13434828
                    ForeColor =128
                    Name ="PARAMS"
                    ControlSource ="=\"انتقال به فاکتور\""
                    OnClick ="[Event Procedure]"
                    ConditionalFormat = Begin
                        0x0100000088000000010000000000000002000000000000001300000001000000 ,
                        0x80000000ccffcc00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x2200270646062a06420627064406200028064706200041062706a9062a064806 ,
                        0x3106220000000000
                    End
                    AsianLineBreak =0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =1422
                            Top =4484
                            Width =585
                            Height =240
                            Name ="Label18"
                            Caption ="."
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "ZF_IVO_EXTENDED.cls"
