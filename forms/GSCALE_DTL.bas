Version =20
VersionRequired =20
Begin Form
    AllowDesignChanges = NotDefault
    DefaultView =2
    AllowUpdating =4
    ViewsAllowed =2
    TabularCharSet =178
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =8560
    RowHeight =400
    ItemSuffix =12
    Left =510
    Top =1935
    Right =9360
    Bottom =8115
    DatasheetForeColor =33554432
    DatasheetGridlinesColor =12632256
    MaxRecords =1000000
    RecSrcDt = Begin
        0x791423b2b88ee540
    End
    RecordSource ="GSCADTL"
    Caption ="GSCALE"
    BeforeUpdate ="[Event Procedure]"
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
            Height =6179
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    ColumnHidden = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6065
                    Top =113
                    Width =1521
                    ColumnOrder =0
                    Name ="GSCACOD"
                    ControlSource ="GSCACOD"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =7653
                            Top =113
                            Width =810
                            Height =240
                            Name ="Label1"
                            Caption ="کد اسکيل "
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1927
                    Top =453
                    Width =5670
                    ColumnWidth =6300
                    ColumnOrder =2
                    TabIndex =1
                    Name ="GSCANAME"
                    ControlSource ="GSCANAME"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =7643
                            Top =453
                            Width =795
                            Height =240
                            Name ="Label3"
                            Caption ="عنوان"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1927
                    Top =793
                    Width =5601
                    ColumnOrder =5
                    TabIndex =2
                    Name ="GSCAGRADE"
                    ControlSource ="GSCAGRADE"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =7643
                            Top =793
                            Width =885
                            Height =240
                            Name ="Label5"
                            Caption ="امتياز"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1927
                    Top =1133
                    Width =5601
                    ColumnWidth =1444
                    ColumnOrder =3
                    TabIndex =3
                    Name ="GSCAFROM"
                    ControlSource ="GSCAFROM"
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =85
                            Left =7643
                            Top =1133
                            Width =810
                            Height =240
                            Name ="Label7"
                            Caption ="از "
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1927
                    Top =1474
                    Width =5601
                    ColumnWidth =1515
                    ColumnOrder =4
                    TabIndex =4
                    Name ="GSCATO"
                    ControlSource ="GSCATO"
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =85
                            Left =7643
                            Top =1474
                            Width =585
                            Height =240
                            Name ="Label9"
                            Caption ="تا"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6009
                    Top =1814
                    Width =1521
                    ColumnWidth =390
                    ColumnOrder =1
                    TabIndex =5
                    Name ="GSCADTCOD"
                    ControlSource ="GSCADTCOD"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =7597
                            Top =1814
                            Width =810
                            Height =240
                            Name ="Label11"
                            Caption ="کد "
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "GSCALE_DTL.cls"
