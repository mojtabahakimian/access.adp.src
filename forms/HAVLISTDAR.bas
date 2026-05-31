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
    TabularCharSet =178
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =11145
    RowHeight =345
    ItemSuffix =18
    Left =1125
    Top =990
    Right =9090
    Bottom =4995
    DatasheetForeColor =33554432
    DatasheetGridlinesColor =12632256
    MaxRecords =100000
    RecSrcDt = Begin
        0x24f53c1317ebe340
    End
    RecordSource ="SELECT NUMBER, TAG, DATE_N, TAH, FNUMCO, SADER, MEGHk, MEGHKS, MAND FROM MEGH_DA"
        "R WHERE (TAG = 23) ORDER BY NUMBER DESC"
    Caption ="ليست درخواست خريد"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Tahoma"
    OnKeyPress ="[Event Procedure]"
    OnError ="[Event Procedure]"
    AllowFormView =0
    Begin
        Begin Label
            BackStyle =0
            TextFontCharSet =178
            FontName ="Tahoma"
        End
        Begin CheckBox
            SpecialEffect =2
            LabelX =230
            LabelY =-30
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
        Begin Subform
            SpecialEffect =2
            Width =1701
            Height =1701
        End
        Begin Section
            CanGrow = NotDefault
            Height =9016
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =7202
                    Top =113
                    ColumnWidth =1440
                    ColumnOrder =0
                    Name ="NUMBER"
                    ControlSource ="NUMBER"
                    OnDblClick ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5505
                            Top =120
                            Width =1335
                            Height =240
                            Name ="Label0"
                            Caption ="شماره درخواست*"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =7202
                    Top =793
                    ColumnWidth =1170
                    ColumnOrder =1
                    TabIndex =1
                    Name ="DATE_N"
                    ControlSource ="DATE_N"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5501
                            Top =793
                            Width =885
                            Height =240
                            Name ="Label2"
                            Caption ="تاريخ "
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =7202
                    Top =1133
                    ColumnWidth =5310
                    ColumnOrder =4
                    TabIndex =2
                    Name ="TAH"
                    ControlSource ="TAH"
                    OnDblClick ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5505
                            Top =1140
                            Width =1335
                            Height =240
                            Name ="Label3"
                            Caption ="در خواست كننده *"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =7202
                    Top =2175
                    ColumnWidth =1275
                    TabIndex =3
                    Name ="FNUMCO"
                    ControlSource ="FNUMCO"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5499
                            Top =2173
                            Width =1020
                            Height =240
                            Name ="Label7"
                            Caption ="شماره داخلي"
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
                    Left =6940
                    Top =2948
                    TabIndex =4
                    Name ="SADER"
                    ControlSource ="SADER"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =9708
                            Top =2952
                            Width =765
                            Height =240
                            Name ="Label12"
                            Caption ="نوع فروش"
                        End
                    End
                End
                Begin Subform
                    OverlapFlags =85
                    Top =3571
                    Width =6825
                    Height =5445
                    TabIndex =5
                    Name ="HAVLISTDAR_SUB"
                    SourceObject ="Form.HAVLISTDAR_SUB"
                    LinkChildFields ="NUMBER"
                    LinkMasterFields ="NUMBER"
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =7483
                    Top =3685
                    ColumnWidth =1515
                    TabIndex =6
                    Name ="MEGHk"
                    ControlSource ="MEGHk"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =9525
                            Top =3690
                            Width =1590
                            Height =285
                            Name ="Label15"
                            Caption ="مقدار كل درخواست"
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
                    Left =7485
                    Top =4200
                    ColumnWidth =1785
                    TabIndex =7
                    Name ="MEGHKS"
                    ControlSource ="MEGHKS"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =10155
                            Top =4200
                            Width =975
                            Height =600
                            Name ="Label16"
                            Caption ="مقدار كل خريداري شده"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =7370
                    Top =4932
                    ColumnWidth =1245
                    TabIndex =8
                    Name ="MAND"
                    ControlSource ="MAND"
                    ConditionalFormat = Begin
                        0x0100000066000000010000000000000004000000000000000200000001000000 ,
                        0x00000000ffcc9900000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x300000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =10200
                            Top =4935
                            Width =945
                            Height =405
                            Name ="Label17"
                            Caption ="مانده درخواست"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "HAVLISTDAR.cls"
