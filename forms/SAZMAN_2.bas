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
    AllowDesignChanges = NotDefault
    DefaultView =0
    AllowUpdating =4
    ScrollBars =0
    ViewsAllowed =1
    TabularCharSet =178
    TabularFamily =0
    BorderStyle =3
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =5612
    DatasheetFontHeight =10
    ItemSuffix =8
    Left =4830
    Top =2475
    Right =10440
    Bottom =5685
    DatasheetGridlinesColor =12632256
    MaxRecords =100000
    RecSrcDt = Begin
        0x1fc8b6839b7fe340
    End
    RecordSource ="SAZMAN"
    Caption ="ادامه مشخصات حقوق"
    DatasheetFontName ="Arial"
    AllowDatasheetView =0
    RecordSourceQualifier ="dbo"
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
        End
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
            Width =850
            Height =850
        End
        Begin CommandButton
            Width =1701
            Height =283
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
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
        Begin Section
            Height =3231
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1305
                    Top =694
                    Height =300
                    Name ="lECOL1"
                    ControlSource ="lECOL1"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =3114
                            Top =694
                            Width =645
                            Height =240
                            Name ="Label0"
                            Caption ="ساير 1"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1305
                    Top =1035
                    Height =300
                    TabIndex =1
                    Name ="lECOL2"
                    ControlSource ="lECOL2"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =3114
                            Top =1035
                            Width =645
                            Height =240
                            Name ="Label1"
                            Caption ="ساير 2"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1305
                    Top =1375
                    Height =300
                    TabIndex =2
                    Name ="lECOL3"
                    ControlSource ="lECOL3"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =3114
                            Top =1375
                            Width =645
                            Height =240
                            Name ="Label2"
                            Caption ="ساير3"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1305
                    Top =2092
                    Height =300
                    TabIndex =4
                    Name ="lKCOL1"
                    ControlSource ="lKCOL1"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =3112
                            Top =2085
                            Width =1245
                            Height =300
                            Name ="Label3"
                            Caption ="ساير كسورات 1"
                        End
                    End
                End
                Begin CommandButton
                    Cancel = NotDefault
                    OverlapFlags =85
                    AccessKey =199
                    Left =2805
                    Top =2835
                    Width =1290
                    Height =345
                    FontWeight =700
                    TabIndex =5
                    Name ="Command147"
                    Caption ="&انصراف"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =208
                    Left =1530
                    Top =2835
                    Width =1275
                    Height =345
                    FontWeight =700
                    TabIndex =6
                    Name ="Command148"
                    Caption ="&ذخيره و خروج"
                    OnClick ="[Event Procedure]"
                End
                Begin Rectangle
                    OverlapFlags =255
                    Left =510
                    Top =300
                    Width =4478
                    Height =2188
                    Name ="Box4"
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =247
                    TextAlign =2
                    Left =2775
                    Top =120
                    Width =1935
                    Height =330
                    BackColor =-2147483633
                    Name ="Label5"
                    Caption ="برچسب فيلدها در حقوق"
                End
                Begin TextBox
                    OverlapFlags =247
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1305
                    Top =1740
                    Height =300
                    TabIndex =3
                    Name ="lECOL4"
                    ControlSource ="lECOL4"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =3114
                            Top =1740
                            Width =645
                            Height =240
                            Name ="Label7"
                            Caption ="ساير4"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "SAZMAN_2.cls"
