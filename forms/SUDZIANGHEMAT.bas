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
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    DatasheetBorderLineStyle =6
    GridY =10
    Width =4940
    RowHeight =345
    ItemSuffix =17
    DatasheetGridlinesColor =12632256
    MaxRecords =100000
    Tag ="@FROOSH=FORMS![BASEKNOW]![FROSH],@GHEMAT = FORMS![BASEKNOW]![GHEYMAT]"
    RecSrcDt = Begin
        0x2b1d85790cdbe240
    End
    InputParameters ="@FROOSH=FORMS![BASEKNOW]![FROSH],@GHEMAT = FORMS![BASEKNOW]![GHEYMAT],@DT1= FORM"
        "S![F_MENU_DATE]![DT1],@DT2= FORMS![F_MENU_DATE]![DT2]"
    RecordSource ="SUDZIANGHEMAT"
    Caption ="مقايسه قيمت تمام شده با فروش"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Tahoma"
    OnError ="[Event Procedure]"
    UnknownProp = {258 ,66 ,12 ,4 ,266 } Begin
        0x4000460052004f004f00530048003d0046004f0052004d00530021005b004200 ,
        0x4100530045004b004e004f0057005d0021005b00460052004f00530048005d00 ,
        0x2c0040004700480045004d004100540020003d00200046004f0052004d005300 ,
        0x21005b0042004100530045004b004e004f0057005d0021005b00470048004500 ,
        0x59004d00410054005d002c0040004400540031003d00200046004f0052004d00 ,
        0x530021005b0046005f004d0045004e0055005f0044004100540045005d002100 ,
        0x5b004400540031005d002c0040004400540032003d00200046004f0052004d00 ,
        0x530021005b0046005f004d0045004e0055005f0044004100540045005d002100 ,
        0x5b004400540032005d00
    End
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
            Height =3548
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1927
                    Top =113
                    Width =2490
                    ColumnWidth =3360
                    ColumnOrder =2
                    Name ="NAME"
                    ControlSource ="NAME"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =113
                            Width =1185
                            Height =240
                            Name ="Label1"
                            Caption ="كالا"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1927
                    Top =453
                    ColumnWidth =1350
                    ColumnOrder =7
                    TabIndex =1
                    Name ="FROOSH"
                    ControlSource ="FROOSH"
                    Format ="#,###"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =226
                            Top =453
                            Width =690
                            Height =240
                            Name ="Label10"
                            Caption ="فروش"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1927
                    Top =793
                    ColumnWidth =1290
                    ColumnOrder =6
                    TabIndex =2
                    Name ="GHEMAT"
                    ControlSource ="GHEMAT"
                    Format ="#,###"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =793
                            Width =690
                            Height =240
                            Name ="Label5"
                            Caption ="قيمت تمام شده"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1927
                    Top =1133
                    ColumnWidth =1260
                    ColumnOrder =8
                    TabIndex =3
                    Name ="SUD"
                    ControlSource ="SUD"
                    Format ="#,###"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =1133
                            Width =375
                            Height =240
                            Name ="Label7"
                            Caption ="سود"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1927
                    Top =1474
                    ColumnWidth =435
                    ColumnOrder =9
                    TabIndex =4
                    Name ="ZIAN"
                    ControlSource ="ZIAN"
                    Format ="#,###"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =1474
                            Width =435
                            Height =240
                            Name ="Label9"
                            Caption ="زيان"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =963
                    Top =2948
                    ColumnWidth =615
                    ColumnOrder =0
                    TabIndex =5
                    Name ="HES_M"
                    ControlSource ="HES_M"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3360
                            Top =2948
                            Width =1005
                            Height =240
                            Name ="Label11"
                            Caption ="كد كالا"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =453
                    Top =1870
                    ColumnWidth =735
                    ColumnOrder =3
                    TabIndex =6
                    Name ="DAST"
                    ControlSource ="DAST"
                    Format ="#,###"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3330
                            Top =1875
                            Width =630
                            Height =240
                            Name ="Label12"
                            Caption ="دستمزد"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =453
                    Top =2211
                    ColumnWidth =570
                    ColumnOrder =4
                    TabIndex =7
                    Name ="SARBAR"
                    ControlSource ="SARBAR"
                    Format ="#,###"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3120
                            Top =2211
                            Width =735
                            Height =240
                            Name ="Label13"
                            Caption ="سربار"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =453
                    Top =2551
                    ColumnWidth =1260
                    ColumnOrder =5
                    TabIndex =8
                    Name ="MAVAD"
                    ControlSource ="MAVAD"
                    Format ="#,###"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3195
                            Top =2551
                            Width =660
                            Height =240
                            Name ="Label14"
                            Caption ="مواد"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =963
                    Top =3308
                    ColumnWidth =1590
                    ColumnOrder =1
                    TabIndex =9
                    Name ="megh"
                    ControlSource ="megh"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3360
                            Top =3308
                            Width =1005
                            Height =240
                            Name ="Label16"
                            Caption ="مقدار"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "SUDZIANGHEMAT.cls"
