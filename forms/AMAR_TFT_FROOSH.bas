Version =20
VersionRequired =20
Begin Form
    AllowDeletions = NotDefault
    AllowAdditions = NotDefault
    DefaultView =2
    AllowUpdating =4
    ViewsAllowed =2
    TabularCharSet =178
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =4940
    RowHeight =315
    ItemSuffix =16
    Left =945
    Top =1815
    Right =6180
    Bottom =4665
    DatasheetForeColor =33554432
    DatasheetGridlinesColor =12632256
    MaxRecords =100000
    RecSrcDt = Begin
        0x50637cf58c29e340
    End
    RecordSource ="SELECT DEED_DTL.HES_K, SUM(DEED_DTL.BED) AS mabl, DEED_DTL.TAG, DEPART.DEPATMAN,"
        " DEPART.DEPNAME FROM DEED_DTL INNER JOIN HEAD_LST ON DEED_DTL.NUMBER = HEAD_LST."
        "NUMBER AND DEED_DTL.TAG = HEAD_LST.TAG INNER JOIN DEPART ON HEAD_LST.DEPATMAN = "
        "DEPART.DEPATMAN WHERE (HEAD_LST.DATE_N BETWEEN 1 AND 3) GROUP BY DEED_DTL.HES_K,"
        " DEED_DTL.TAG, DEPART.DEPATMAN, DEPART.DEPNAME HAVING (DEED_DTL.TAG = 2) AND (DE"
        "ED_DTL.HES_K = 514)"
    Caption ="آمار تخفيفات واحدها"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Tahoma"
    AllowFormView =0
    RecordSourceQualifier ="dbo"
    Begin
        Begin Label
            BackStyle =0
            TextFontCharSet =178
            FontName ="Tahoma"
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
        Begin Section
            Height =2854
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1927
                    Top =453
                    ColumnWidth =705
                    Name ="DEPATMAN"
                    ControlSource ="DEPATMAN"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =453
                            Width =810
                            Height =240
                            Name ="Label3"
                            Caption ="كد واحد"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1927
                    Top =793
                    Width =2490
                    ColumnWidth =3255
                    TabIndex =1
                    Name ="DEPNAME"
                    ControlSource ="DEPNAME"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =795
                            Width =630
                            Height =240
                            Name ="Label5"
                            Caption ="نام واحد"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1927
                    Top =1133
                    TabIndex =2
                    Name ="MTAKH"
                    ControlSource ="mabl"
                    Format ="#,###"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =1140
                            Width =870
                            Height =240
                            Name ="Label7"
                            Caption ="مبلغ تخفيف"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "AMAR_TFT_FROOSH.cls"
