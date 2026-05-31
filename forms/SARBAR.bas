Version =20
VersionRequired =20
Begin Form
    AutoResize = NotDefault
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    AllowEdits = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    AllowUpdating =4
    ScrollBars =0
    ViewsAllowed =1
    TabularCharSet =178
    TabularFamily =99
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =5669
    DatasheetFontHeight =10
    ItemSuffix =2
    Left =270
    Top =600
    Right =7725
    Bottom =3045
    DatasheetGridlinesColor =12632256
    MaxRecords =100000
    RecSrcDt = Begin
        0x7dfaa72cd9c0e240
    End
    RecordSource ="SELECT PRGHEAD.GRP, SUM(HEAD_MANF.IMBIBE_MANF) AS DAST, SUM(HEAD_MANF.IMBIBE_SAR"
        ") AS SARB FROM PRGHEAD INNER JOIN HEAD_MANF ON PRGHEAD.PCODE = HEAD_MANF.CODE GR"
        "OUP BY PRGHEAD.GRP"
    DatasheetFontName =""
    AllowDatasheetView =0
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
            Height =396
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2834
                    Top =56
                    Height =285
                    Name ="DAST"
                    ControlSource ="DAST"
                    Format ="#,###"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4592
                            Top =56
                            Width =1065
                            Height =240
                            Name ="Label0"
                            Caption ="جذب دستمزد:"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =56
                    Top =56
                    Height =285
                    TabIndex =1
                    Name ="SARB"
                    ControlSource ="SARB"
                    Format ="#,###"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1870
                            Top =56
                            Width =900
                            Height =240
                            Name ="Label1"
                            Caption ="جذب سربار:"
                        End
                    End
                End
            End
        End
    End
End
