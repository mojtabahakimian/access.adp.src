Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    DefaultView =2
    AllowUpdating =4
    ViewsAllowed =2
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =4940
    RowHeight =300
    ItemSuffix =4
    Left =570
    Top =3750
    Right =3030
    Bottom =5340
    DatasheetGridlinesColor =12632256
    MaxRecords =100000
    RecSrcDt = Begin
        0x5fb397177bf4e240
    End
    RecordSource ="DAFT_ASN"
    Caption ="DAFT_ASN"
    OnDelete ="[Event Procedure]"
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
        Begin Section
            Height =813
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1927
                    Top =113
                    ColumnWidth =1095
                    Name ="FIRSTNUM"
                    ControlSource ="FIRSTNUM"
                    BeforeUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =120
                            Width =990
                            Height =240
                            Name ="Label1"
                            Caption ="شماره شروع"
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
                    Left =1927
                    Top =453
                    ColumnWidth =960
                    TabIndex =1
                    Name ="BOOKNUM"
                    ControlSource ="BOOKNUM"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =450
                            Width =855
                            Height =240
                            Name ="Label3"
                            Caption ="شماره دفتر"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "FDAFT_ASN.cls"
