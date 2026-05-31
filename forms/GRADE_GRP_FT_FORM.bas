Version =20
VersionRequired =20
Begin Form
    DefaultView =2
    AllowUpdating =4
    TabularCharSet =178
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =5808
    RowHeight =315
    ItemSuffix =13
    Left =1335
    Top =3615
    Right =14865
    Bottom =5505
    DatasheetForeColor =33554432
    DatasheetGridlinesColor =12632256
    MaxRecords =1000000
    RecSrcDt = Begin
        0x333fae36198ee540
    End
    RecordSource ="GRADE_GRP_FT"
    Caption ="فرمت طبقه"
    BeforeUpdate ="[Event Procedure]"
    AfterUpdate ="[Event Procedure]"
    DatasheetFontName ="Tahoma"
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
        Begin ComboBox
            SpecialEffect =2
            TextFontCharSet =178
            Width =1701
            LabelX =-1701
            FontName ="Tahoma"
        End
        Begin Subform
            SpecialEffect =2
            Width =1701
            Height =1701
        End
        Begin Section
            Height =4474
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1927
                    Top =113
                    ColumnOrder =1
                    Name ="GFID"
                    ControlSource ="GFTID"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =120
                            Width =960
                            Height =240
                            Name ="Label1"
                            Caption ="شماره طبقه"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1927
                    Top =453
                    Width =2475
                    ColumnWidth =7440
                    ColumnOrder =2
                    TabIndex =1
                    Name ="GFGRPNAMEFT"
                    ControlSource ="GFGRPNAMEFT"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =450
                            Width =1050
                            Height =240
                            Name ="Label3"
                            Caption ="عنوان سنجش"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1927
                    Top =793
                    ColumnOrder =5
                    TabIndex =2
                    Name ="GFGRPZARIB"
                    ControlSource ="GFGRPZARIB"
                    ValidationRule ="Not Is Null"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =795
                            Width =870
                            Height =240
                            Name ="Label5"
                            Caption ="ضريب "
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1927
                    Top =1133
                    ColumnWidth =930
                    ColumnOrder =0
                    TabIndex =3
                    Name ="GFTGRPID"
                    ControlSource ="GFTGRPID"
                    ValidationRule ="Not Is Null"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =1140
                            Width =870
                            Height =240
                            Name ="Label7"
                            Caption ="شماره گروه"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1927
                    Top =1590
                    ColumnOrder =4
                    TabIndex =4
                    Name ="GFGRPGRADE"
                    ControlSource ="GFGRPGRADE"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =1592
                            Width =870
                            Height =240
                            Name ="Label10"
                            Caption ="امتياز"
                        End
                    End
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =1927
                    Top =1952
                    ColumnWidth =2970
                    ColumnOrder =3
                    TabIndex =5
                    Name ="GVALUESCALE"
                    ControlSource ="GVALUESCALE"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT GSCACOD, GSCANAME FROM GSCALE ORDER BY GSCANAME"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =1954
                            Width =870
                            Height =240
                            Name ="Label12"
                            Caption ="ارزش"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "GRADE_GRP_FT_FORM.cls"
