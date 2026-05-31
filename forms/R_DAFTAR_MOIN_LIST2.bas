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
    Width =5865
    RowHeight =330
    ItemSuffix =41
    Right =15120
    Bottom =3780
    DatasheetGridlinesColor =12632256
    MaxRecords =1000000
    RecSrcDt = Begin
        0x84ed13081b76e640
    End
    Caption ="دفتر معين"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Tahoma"
    OnKeyPress ="[Event Procedure]"
    OnError ="[Event Procedure]"
    AllowFormView =0
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
            Height =6992
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
                    Left =2146
                    ColumnWidth =990
                    ColumnOrder =0
                    Name ="N_S"
                    ControlSource ="N_S"
                    OnDblClick ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =339
                            Top =7
                            Width =885
                            Height =240
                            Name ="Label1"
                            Caption ="شماره سند"
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
                    Left =2146
                    Top =340
                    ColumnWidth =1020
                    ColumnOrder =1
                    TabIndex =1
                    Name ="DATE_S"
                    ControlSource ="DATE_S"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =339
                            Top =337
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
                    Left =2146
                    Top =680
                    ColumnOrder =2
                    TabIndex =2
                    Name ="HES_K"
                    ControlSource ="HES_K"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =339
                            Top =682
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
                    Left =2146
                    Top =1020
                    ColumnOrder =3
                    TabIndex =3
                    Name ="HES_M"
                    ControlSource ="HES_M"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =332
                            Top =1020
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
                    Left =2146
                    Top =1361
                    Width =2475
                    ColumnWidth =8595
                    ColumnOrder =4
                    TabIndex =4
                    Name ="SHARH"
                    ControlSource ="SHARH"
                    OnDblClick ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =332
                            Top =1361
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
                    Left =2146
                    Top =1701
                    ColumnWidth =1530
                    ColumnOrder =5
                    TabIndex =5
                    Name ="BED"
                    ControlSource ="BED"
                    Format ="#,##0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =332
                            Top =1701
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
                    Left =2146
                    Top =2041
                    ColumnWidth =1485
                    ColumnOrder =6
                    TabIndex =6
                    Name ="BES"
                    ControlSource ="BES"
                    Format ="#,##0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =332
                            Top =2041
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
                    Left =2146
                    Top =2381
                    ColumnWidth =1605
                    ColumnOrder =8
                    TabIndex =7
                    Name ="MAND11"
                    ControlSource ="MAND"
                    Format ="#,##0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =332
                            Top =2381
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
                    Left =2146
                    Top =4082
                    ColumnOrder =9
                    TabIndex =8
                    Name ="NO_S"
                    ControlSource ="NO_S"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =339
                            Top =4087
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
                    Left =2146
                    Top =4422
                    ColumnOrder =10
                    TabIndex =9
                    Name ="N_SERI"
                    ControlSource ="N_SERI"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =332
                            Top =4422
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
                    Left =2146
                    Top =4762
                    ColumnOrder =11
                    TabIndex =10
                    Name ="BANK"
                    ControlSource ="BANK"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =332
                            Top =4762
                            Width =465
                            Height =240
                            Name ="Label29"
                            Caption ="بانك"
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
                    Left =2146
                    Top =5102
                    ColumnOrder =12
                    TabIndex =11
                    Name ="NUMBER"
                    ControlSource ="NUMBER"
                    OnDblClick ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =332
                            Top =5102
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
                    Left =2146
                    Top =5442
                    ColumnOrder =13
                    TabIndex =12
                    Name ="TAG1"
                    ControlSource ="TAG"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =332
                            Top =5442
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
                    Left =2203
                    Top =3572
                    ColumnWidth =435
                    ColumnOrder =7
                    TabIndex =13
                    Name ="Text34"
                    ControlSource ="TASH"
                    Begin
                        Begin Label
                            Vertical = NotDefault
                            OverlapFlags =87
                            Left =389
                            Top =3572
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
                    Left =2494
                    Top =3288
                    TabIndex =14
                    Name ="HES_T"
                    ControlSource ="HES_T"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =680
                            Top =3288
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
                    Left =1821
                    Top =6746
                    Width =2475
                    TabIndex =15
                    Name ="TAFZILN"
                    ControlSource ="TAFZILN"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Top =6752
                            Width =1380
                            Height =240
                            Name ="Label39"
                            Caption ="نام حساب تفصيلي2"
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
                    Left =2146
                    Top =5802
                    TabIndex =16
                    Name ="ARZD"
                    ControlSource ="ARZD"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =325
                            Top =5797
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
                    IMESentenceMode =3
                    Left =283
                    Top =6236
                    TabIndex =17
                    Name ="HES_T2"
                    ControlSource ="HES_T2"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2995
                            Top =6232
                            Width =765
                            Height =240
                            Name ="Label42"
                            Caption ="تفصيلي 2"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =215
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    TabIndex =18
                    Name ="hes"
                    ControlSource ="hes"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =3008
                            Top =6
                            Width =525
                            Height =405
                            Name ="Label44"
                            Caption ="حساب:"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "R_DAFTAR_MOIN_LIST2.cls"
