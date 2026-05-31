Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    AllowEdits = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =2
    AllowUpdating =4
    ViewsAllowed =2
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =5669
    RowHeight =315
    ItemSuffix =12
    Left =405
    Top =1560
    Right =14175
    Bottom =9510
    DatasheetGridlinesColor =12632256
    Tag ="@F_MENU_DATE_DT1=FORMS![F_MENU_DATE]![DT1],@F_MENU_DATE_DT2=FORMS![F_MENU_DATE]!"
        "[DT2]"
    Filter ="DEPATMAN = 1"
    RecSrcDt = Begin
        0xe05537deeeafe240
    End
    InputParameters ="@F_MENU_DATE_DT1=FORMS![F_MENU_DATE]![DT1],@F_MENU_DATE_DT2=FORMS![F_MENU_DATE]!"
        "[DT2]"
    RecordSource ="FROOSH_COUNT"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Tahoma"
    OnError ="[Event Procedure]"
    UnknownProp = {258 ,66 ,12 ,4 ,170 } Begin
        0x400046005f004d0045004e0055005f0044004100540045005f00440054003100 ,
        0x3d0046004f0052004d00530021005b0046005f004d0045004e0055005f004400 ,
        0x4100540045005d0021005b004400540031005d002c00400046005f004d004500 ,
        0x4e0055005f0044004100540045005f004400540032003d0046004f0052004d00 ,
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
        Begin FormHeader
            Height =623
            BackColor =-2147483633
            Name ="FormHeader"
        End
        Begin Section
            Height =2847
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
                    Left =3401
                    Top =907
                    ColumnOrder =3
                    Name ="TEDAD"
                    ControlSource ="TEDAD"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1695
                            Top =900
                            Width =855
                            Height =240
                            Name ="Label0"
                            Caption ="تعداد فاكتور"
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
                    Left =3401
                    Top =1247
                    ColumnOrder =0
                    TabIndex =1
                    Name ="CODE"
                    ControlSource ="CODE"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1695
                            Top =1245
                            Width =855
                            Height =240
                            Name ="Label1"
                            Caption ="كد كالا"
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
                    Left =3401
                    Top =1587
                    ColumnWidth =6705
                    ColumnOrder =1
                    TabIndex =2
                    Name ="NAME"
                    ControlSource ="NAME"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1700
                            Top =1587
                            Width =555
                            Height =240
                            Name ="Label2"
                            Caption ="نام كالا"
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
                    Left =3401
                    Top =1927
                    ColumnWidth =1410
                    ColumnOrder =2
                    TabIndex =3
                    Name ="SMEGHk"
                    ControlSource ="SMEGHk"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1700
                            Top =1927
                            Width =1020
                            Height =240
                            Name ="Label3"
                            Caption ="مقدار كل"
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
                    Left =3401
                    Top =2267
                    ColumnWidth =645
                    TabIndex =4
                    Name ="DEPATMAN"
                    ControlSource ="DEPATMAN"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1700
                            Top =2267
                            Width =600
                            Height =240
                            Name ="Label4"
                            Caption ="دپاتمان"
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
                    Left =3401
                    Top =2607
                    TabIndex =5
                    Name ="DEPNAME"
                    ControlSource ="DEPNAME"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1700
                            Top =2607
                            Width =330
                            Height =240
                            Name ="Label5"
                            Caption ="نام"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =3408
                    Top =566
                    Width =891
                    ColumnWidth =2025
                    TabIndex =6
                    Name ="SMABL"
                    ControlSource ="SMABL"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =1927
                            Top =566
                            Width =615
                            Height =240
                            Name ="Label10"
                            Caption ="مبلغ"
                        End
                    End
                End
            End
        End
        Begin FormFooter
            Height =1272
            BackColor =-2147483633
            Name ="FormFooter"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1479
                    Top =226
                    Name ="STEDAD"
                    ControlSource ="=Sum([TEDAD])"
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1474
                    Top =672
                    TabIndex =1
                    Name ="SSMEGHk"
                    ControlSource ="=Sum([SMEGHk])"
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1474
                    Top =1032
                    TabIndex =2
                    Name ="smab"
                    ControlSource ="=Sum([SMABL])"
                End
            End
        End
    End
End
CodeBehindForm
' See "FROOSH_COUNT_SUB.cls"
