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
    TabularCharSet =178
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =5787
    RowHeight =270
    ItemSuffix =23
    Left =300
    Top =1395
    Right =14775
    Bottom =8700
    DatasheetForeColor =33554432
    DatasheetGridlinesColor =12632256
    MaxRecords =100000
    OrderBy ="PFAMILY"
    RecSrcDt = Begin
        0x7754e70c3772e340
    End
    RecordSource ="SELECT PERSONEL.CODE, PERSONEL.BIMEH_NUM, PERSONEL.PNAME, PERSONEL.PFAMILY, PERS"
        "ONEL.JOB, SALARY_EYDY.ROOZ, SALARY_EYDY.DARAMAD, SALARY_EYDY.MALIAT, SALARY_EYDY"
        ".EYDY, SALARY_EYDY.SANAVAT, SALARY_EYDY.MOAFIYAT, SALARY_EYDY.DARMASH, SALARY_EY"
        "DY.MALIKOL, SALARY_EYDY.EPM, SALARY_EYDY.ROOZ + SALARY_EYDY.EPM AS ROZHA, SALARY"
        "_EYDY.EYDYMAL, SALARY_EYDY.DARAMAD + SALARY_EYDY.EYDY AS darkol FROM SALARY_EYDY"
        " INNER JOIN PERSONEL ON SALARY_EYDY.CODE = PERSONEL.CODE ORDER BY PERSONEL.CODE"
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
            Height =7041
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =3798
                    Top =907
                    ColumnWidth =660
                    ColumnOrder =0
                    Name ="CODE"
                    ControlSource ="CODE"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2097
                            Top =907
                            Width =555
                            Height =240
                            Name ="Label0"
                            Caption ="كد"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =3798
                    Top =1247
                    ColumnWidth =990
                    ColumnOrder =16
                    TabIndex =1
                    Name ="BIMEH_NUM"
                    ControlSource ="BIMEH_NUM"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2097
                            Top =1247
                            Width =930
                            Height =240
                            Name ="Label1"
                            Caption ="شماره بيمه"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =3798
                    Top =1587
                    ColumnWidth =1080
                    ColumnOrder =1
                    TabIndex =2
                    Name ="PNAME"
                    ControlSource ="PNAME"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2097
                            Top =1587
                            Width =330
                            Height =240
                            Name ="Label2"
                            Caption ="نام"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =3798
                    Top =1927
                    ColumnWidth =1515
                    ColumnOrder =2
                    TabIndex =3
                    Name ="PFAMILY"
                    ControlSource ="PFAMILY"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2097
                            Top =1927
                            Width =1005
                            Height =240
                            Name ="Label3"
                            Caption ="نام خانوادگي"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =3798
                    Top =2607
                    ColumnWidth =930
                    ColumnOrder =3
                    TabIndex =4
                    Name ="JOB"
                    ControlSource ="JOB"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2097
                            Top =2607
                            Width =510
                            Height =240
                            Name ="Label5"
                            Caption ="شغل"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =3798
                    Top =2948
                    ColumnWidth =1485
                    ColumnOrder =4
                    TabIndex =5
                    Name ="ROOZ"
                    ControlSource ="ROOZ"
                    Format ="#,###"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2100
                            Top =2955
                            Width =1380
                            Height =240
                            Name ="Label6"
                            Caption ="روز كاركرد در ليست"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =3798
                    Top =3288
                    ColumnOrder =7
                    TabIndex =6
                    Name ="DARAMAD"
                    ControlSource ="DARAMAD"
                    Format ="#,###"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2100
                            Top =3285
                            Width =1305
                            Height =240
                            Name ="Label7"
                            Caption ="جمع حقوق و مزايا"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =3798
                    Top =3628
                    ColumnWidth =1215
                    ColumnOrder =13
                    TabIndex =7
                    Name ="MALIAT"
                    ControlSource ="MALIAT"
                    Format ="#,###"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2100
                            Top =3630
                            Width =1110
                            Height =240
                            Name ="Label8"
                            Caption ="ماليات پرداختي"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2607
                    Top =4478
                    ColumnOrder =8
                    TabIndex =8
                    Name ="EYDY"
                    ControlSource ="EYDY"
                    Format ="#,###"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =5160
                            Top =4485
                            Width =555
                            Height =405
                            Name ="Label11"
                            Caption ="عيدي پاداش"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2607
                    Top =4818
                    ColumnOrder =15
                    TabIndex =9
                    Name ="SANAVAT"
                    ControlSource ="SANAVAT"
                    Format ="#,###"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =4845
                            Top =4815
                            Width =855
                            Height =240
                            Name ="Label12"
                            Caption ="حق سنوات"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2607
                    Top =5159
                    ColumnOrder =10
                    TabIndex =10
                    Name ="MOAFIYAT"
                    ControlSource ="MOAFIYAT"
                    Format ="#,###"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4754
                            Top =5159
                            Width =915
                            Height =240
                            Name ="Label13"
                            Caption ="معافيت"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2607
                    Top =5499
                    ColumnOrder =11
                    TabIndex =11
                    Name ="DARMASH"
                    ControlSource ="DARMASH"
                    Format ="#,###"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =4800
                            Top =5505
                            Width =930
                            Height =405
                            Name ="Label14"
                            Caption ="درآمد مشمول"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2607
                    Top =5839
                    ColumnOrder =12
                    TabIndex =12
                    Name ="MALIKOL"
                    ControlSource ="MALIKOL"
                    Format ="#,###"
                    Begin
                        Begin Label
                            OverlapFlags =223
                            Left =4890
                            Top =5835
                            Width =855
                            Height =405
                            Name ="Label15"
                            Caption ="ماليات كل سال"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2607
                    Top =6179
                    ColumnOrder =5
                    TabIndex =13
                    Name ="EPM"
                    ControlSource ="EPM"
                    Format ="#,###"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =5235
                            Top =6180
                            Width =525
                            Height =795
                            Name ="Label16"
                            Caption ="استراحت پزشكي"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =3852
                    Top =4025
                    ColumnWidth =660
                    ColumnOrder =6
                    TabIndex =14
                    Name ="ROZHA"
                    ControlSource ="ROZHA"
                    Format ="#,###"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2154
                            Top =4032
                            Width =1380
                            Height =240
                            Name ="Label18"
                            Caption ="كل روز "
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2664
                    ColumnOrder =14
                    TabIndex =15
                    Name ="EYDYMAL"
                    ControlSource ="EYDYMAL"
                    Format ="#,###"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4857
                            Top =6
                            Width =930
                            Height =405
                            Name ="Label20"
                            Caption ="تتمه ماليات عيدي"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Top =6630
                    ColumnOrder =9
                    TabIndex =16
                    Name ="darkol"
                    ControlSource ="darkol"
                    Format ="#,###"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2193
                            Top =6636
                            Width =930
                            Height =405
                            Name ="Label22"
                            Caption ="درآمد كل سال"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "SALARY_EYDY_SUB.cls"
