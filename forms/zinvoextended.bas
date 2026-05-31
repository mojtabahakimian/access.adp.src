Version =20
VersionRequired =20
Begin Form
    DefaultView =2
    AllowUpdating =4
    ViewsAllowed =2
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =4940
    RowHeight =375
    ItemSuffix =36
    DatasheetGridlinesColor =12632256
    MaxRecords =1000000
    RecSrcDt = Begin
        0x551344ad2bfee540
    End
    RecordSource ="zchrbi"
    Caption ="پارامترها"
    DatasheetFontName ="Tahoma"
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
        Begin Section
            Height =6803
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1927
                    Top =113
                    Width =1710
                    ColumnWidth =1140
                    ColumnOrder =0
                    Name ="DATE_N"
                    ControlSource ="DATE_N"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =113
                            Width =660
                            Height =240
                            Name ="Label1"
                            Caption ="تاريخ"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1927
                    Top =3855
                    ColumnWidth =930
                    ColumnOrder =6
                    TabIndex =1
                    Name ="MEGHK"
                    ControlSource ="MEGHK"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =3855
                            Width =585
                            Height =240
                            Name ="Label23"
                            Caption ="مقدار کل"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1927
                    Top =4195
                    Width =2475
                    ColumnWidth =1035
                    ColumnOrder =1
                    TabIndex =2
                    Name ="CUST_NO"
                    ControlSource ="CUST_NO"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =4195
                            Width =780
                            Height =240
                            Name ="Label25"
                            Caption ="کد مشتري"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1927
                    Top =4535
                    Width =2475
                    ColumnWidth =3675
                    ColumnOrder =2
                    TabIndex =3
                    Name ="NAME"
                    ControlSource ="NAME"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =4535
                            Width =495
                            Height =240
                            Name ="Label27"
                            Caption ="نام"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1927
                    Top =4875
                    ColumnWidth =1005
                    ColumnOrder =3
                    TabIndex =4
                    Name ="NUMBER"
                    ControlSource ="NUMBER"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =4875
                            Width =690
                            Height =240
                            Name ="Label29"
                            Caption ="شماره فاکتور"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1927
                    Top =5215
                    ColumnOrder =4
                    TabIndex =5
                    Name ="TAG"
                    ControlSource ="TAG"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =5215
                            Width =375
                            Height =240
                            Name ="Label31"
                            Caption ="TAG"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1927
                    Top =5555
                    ColumnWidth =930
                    ColumnOrder =5
                    TabIndex =6
                    Name ="mm"
                    ControlSource ="mm"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =5555
                            Width =315
                            Height =240
                            Name ="Label33"
                            Caption ="ماه"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =2041
                    Top =399
                    ColumnWidth =930
                    TabIndex =7
                    Name ="FLD1"
                    ControlSource ="FLD1"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =234
                            Top =396
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
                    Left =2041
                    Top =739
                    ColumnWidth =930
                    TabIndex =8
                    Name ="FLD2"
                    ControlSource ="FLD2"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =234
                            Top =741
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
                    Left =2041
                    Top =1079
                    ColumnWidth =930
                    TabIndex =9
                    Name ="FLD3"
                    ControlSource ="FLD3"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =234
                            Top =1086
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
                    Left =2041
                    Top =1420
                    ColumnWidth =930
                    TabIndex =10
                    Name ="FLD4"
                    ControlSource ="FLD4"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =227
                            Top =1420
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
                    Left =2041
                    Top =1760
                    ColumnWidth =930
                    TabIndex =11
                    Name ="FLD5"
                    ControlSource ="FLD5"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =227
                            Top =1760
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
                    Left =170
                    Top =2383
                    ColumnWidth =930
                    TabIndex =12
                    Name ="FLD6"
                    ControlSource ="FLD6"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3084
                            Top =2376
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
                    Left =170
                    Top =2723
                    ColumnWidth =930
                    TabIndex =13
                    Name ="FLD7"
                    ControlSource ="FLD7"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3084
                            Top =2721
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
                    Left =170
                    Top =3064
                    ColumnWidth =930
                    TabIndex =14
                    Name ="FLD8"
                    ControlSource ="FLD8"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3077
                            Top =3064
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
                    Left =170
                    Top =3460
                    ColumnWidth =930
                    TabIndex =15
                    Name ="FLD9"
                    ControlSource ="FLD9"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3077
                            Top =3460
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
                    Left =1077
                    Top =2040
                    ColumnWidth =930
                    TabIndex =16
                    Name ="FLD10"
                    ControlSource ="FLD10"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =3894
                            Top =2040
                            Width =585
                            Height =240
                            Name ="Label16"
                            Caption ="الکل"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1133
                    Top =5952
                    TabIndex =17
                    Name ="DEPATMAN"
                    ControlSource ="DEPATMAN"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3590
                            Top =5952
                            Width =945
                            Height =240
                            Name ="Label34"
                            Caption ="واحد"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1020
                    Top =6406
                    TabIndex =18
                    Name ="DEPNAME"
                    ControlSource ="DEPNAME"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3582
                            Top =6406
                            Width =840
                            Height =240
                            Name ="Label35"
                            Caption ="دپارتمان"
                        End
                    End
                End
            End
        End
    End
End
