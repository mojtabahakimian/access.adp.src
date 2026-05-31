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
    RowHeight =405
    ItemSuffix =9
    Left =405
    Top =2325
    Right =15240
    Bottom =10155
    DatasheetForeColor =33554432
    DatasheetGridlinesColor =12632256
    MaxRecords =1000000
    RecSrcDt = Begin
        0x03e7b3ce188ee540
    End
    RecordSource ="GRADE_TAB_FT"
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
                    ControlSource ="GFID"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =120
                            Width =960
                            Height =240
                            Name ="Label1"
                            Caption ="شماره فرمت"
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
                    ColumnWidth =12195
                    ColumnOrder =2
                    TabIndex =1
                    Name ="GFNAMEFT"
                    ControlSource ="GFNAMEFT"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =453
                            Width =870
                            Height =240
                            Name ="Label3"
                            Caption ="نام گروه"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1977
                    Top =793
                    ColumnOrder =3
                    TabIndex =2
                    Name ="GFGZARIB"
                    ControlSource ="GFGZARIB"
                    ValidationRule ="Not Is Null"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =170
                            Top =795
                            Width =870
                            Height =240
                            Name ="Label5"
                            Caption ="ضريب طبقه"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1927
                    Top =1133
                    ColumnWidth =585
                    ColumnOrder =0
                    TabIndex =3
                    Name ="GFTID"
                    ControlSource ="GFTID"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =1133
                            Width =525
                            Height =240
                            Name ="Label7"
                            Caption ="شماره"
                        End
                    End
                End
                Begin Subform
                    OverlapFlags =85
                    Left =283
                    Top =1474
                    Width =5381
                    Height =2880
                    TabIndex =4
                    Name ="GRADE_GRP_FT_FORM"
                    SourceObject ="Form.GRADE_GRP_FT_FORM"
                    LinkChildFields ="GFTID"
                    LinkMasterFields ="GFTID"
                End
            End
        End
    End
End
CodeBehindForm
' See "GRADE_TAB_FT_FORM.cls"
