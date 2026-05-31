Version =20
VersionRequired =20
Begin Form
    AllowDeletions = NotDefault
    AllowAdditions = NotDefault
    KeyPreview = NotDefault
    AllowEdits = NotDefault
    DefaultView =2
    AllowUpdating =4
    ViewsAllowed =2
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =5442
    RowHeight =330
    ItemSuffix =40
    Left =7485
    Top =690
    Right =12870
    Bottom =2280
    DatasheetGridlinesColor =12632256
    MaxRecords =1000000
    RecSrcDt = Begin
        0xc45aa3e84f2ae540
    End
    RecordSource ="MOIN1"
    Caption ="دفتر معين"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Tahoma"
    OnKeyPress ="[Event Procedure]"
    OnLoad ="[Event Procedure]"
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
        Begin Section
            Height =7107
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2826
                    Top =113
                    ColumnWidth =795
                    ColumnOrder =0
                    Name ="N_S"
                    ControlSource ="N_S"
                    OnDblClick ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1019
                            Top =120
                            Width =885
                            Height =240
                            Name ="Label1"
                            Caption ="شماره سند"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2826
                    Top =453
                    ColumnWidth =1125
                    ColumnOrder =1
                    TabIndex =1
                    Name ="DATE_S"
                    ControlSource ="DATE_S"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1019
                            Top =450
                            Width =750
                            Height =240
                            Name ="Label3"
                            Caption ="تاريخ سند"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2826
                    Top =793
                    ColumnWidth =885
                    ColumnOrder =2
                    TabIndex =2
                    Name ="HES_K"
                    ControlSource ="HES_K"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1019
                            Top =795
                            Width =780
                            Height =240
                            Name ="Label5"
                            Caption ="حساب كل"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2826
                    Top =1133
                    ColumnWidth =345
                    ColumnOrder =3
                    TabIndex =3
                    Name ="HES_M"
                    ControlSource ="HES_M"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1012
                            Top =1133
                            Width =1380
                            Height =240
                            Name ="Label7"
                            Caption ="حساب معين"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2826
                    Top =1474
                    Width =2475
                    ColumnWidth =7695
                    ColumnOrder =5
                    TabIndex =4
                    Name ="SHARH"
                    ControlSource ="SHARH"
                    OnDblClick ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1012
                            Top =1474
                            Width =585
                            Height =240
                            Name ="Label9"
                            Caption ="شرح"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2826
                    Top =1814
                    ColumnWidth =1770
                    ColumnOrder =6
                    TabIndex =5
                    Name ="BED"
                    ControlSource ="BED"
                    Format ="#,##0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1012
                            Top =1814
                            Width =360
                            Height =240
                            Name ="Label11"
                            Caption ="بدهكار"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2826
                    Top =2154
                    ColumnWidth =1680
                    ColumnOrder =7
                    TabIndex =6
                    Name ="BES"
                    ControlSource ="BES"
                    Format ="#,##0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1012
                            Top =2154
                            Width =345
                            Height =240
                            Name ="Label13"
                            Caption ="بستانكار"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2826
                    Top =2494
                    ColumnWidth =1635
                    ColumnOrder =9
                    TabIndex =7
                    Name ="MAND11"
                    ControlSource ="=Abs([MAND])"
                    Format ="#,##0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1012
                            Top =2494
                            Width =510
                            Height =240
                            Name ="Label15"
                            Caption ="مانده"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2826
                    Top =4195
                    ColumnWidth =660
                    ColumnOrder =15
                    TabIndex =8
                    Name ="NO_S"
                    ControlSource ="NO_S"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =1019
                            Top =4200
                            Width =660
                            Height =240
                            Name ="Label25"
                            Caption ="نوع سند"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2826
                    Top =4535
                    ColumnWidth =1110
                    ColumnOrder =12
                    TabIndex =9
                    Name ="N_SERI"
                    ControlSource ="N_SERI"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1012
                            Top =4535
                            Width =615
                            Height =240
                            Name ="Label27"
                            Caption ="شماره سريال"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2826
                    Top =4875
                    ColumnWidth =0
                    ColumnOrder =13
                    TabIndex =10
                    Name ="BANK"
                    ControlSource ="BANK"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1012
                            Top =4875
                            Width =465
                            Height =240
                            Name ="Label29"
                            Caption ="بانك"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2826
                    Top =5215
                    ColumnWidth =540
                    ColumnOrder =4
                    TabIndex =11
                    Name ="NUMBER"
                    ControlSource ="NUMBER"
                    OnDblClick ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1012
                            Top =5215
                            Width =690
                            Height =240
                            Name ="Label31"
                            Caption ="حواله"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2826
                    Top =5555
                    ColumnWidth =0
                    ColumnOrder =14
                    TabIndex =12
                    Name ="TAG"
                    ControlSource ="TAG"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1012
                            Top =5555
                            Width =375
                            Height =240
                            Name ="Label33"
                            Caption ="نوع فاكتور"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2883
                    Top =3685
                    ColumnWidth =450
                    ColumnOrder =8
                    TabIndex =13
                    Name ="Text34"
                    ControlSource ="=IIf([MAND]>0,\"بد\",IIf([mand]=0,\"--\",\"بس\"))"
                    OnDblClick ="[Event Procedure]"
                    Begin
                        Begin Label
                            Vertical = NotDefault
                            OverlapFlags =87
                            Left =1069
                            Top =3685
                            Width =510
                            Height =510
                            Name ="Label35"
                            Caption ="تش"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2721
                    Top =2891
                    TabIndex =14
                    Name ="HES_T"
                    ControlSource ="HES_T"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =907
                            Top =2891
                            Width =1380
                            Height =240
                            Name ="Label37"
                            Caption ="تفصيلي"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2614
                    Top =3231
                    Width =2475
                    ColumnWidth =240
                    TabIndex =15
                    Name ="TAFZILN"
                    ControlSource ="TAFZILN"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =793
                            Top =3237
                            Width =1380
                            Height =240
                            Name ="Label39"
                            Caption ="نام حساب تفصيلي"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2728
                    Top =6014
                    ColumnWidth =1620
                    TabIndex =16
                    Name ="ARZD"
                    ControlSource ="ARZD"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =907
                            Top =6009
                            Width =510
                            Height =240
                            Name ="Label41"
                            Caption ="نرخ ارز"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2398
                    Top =6746
                    TabIndex =17
                    Name ="HES"
                    ControlSource ="HES"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =793
                            Top =6867
                            Width =540
                            Height =240
                            Name ="Label43"
                            Caption ="حساب"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "R_DAFTAR_MOIN_LIST_ARZI_NT.cls"
