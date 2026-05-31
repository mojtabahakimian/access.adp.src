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
    ItemSuffix =17
    Left =525
    Top =3615
    Right =15090
    Bottom =9000
    DatasheetForeColor =33554432
    DatasheetGridlinesColor =12632256
    MaxRecords =1000000
    RecSrcDt = Begin
        0xe93d6d351e8ee540
    End
    RecordSource ="GRADE_CUST_TAB"
    Caption ="عناوين گروه"
    BeforeUpdate ="[Event Procedure]"
    AfterUpdate ="[Event Procedure]"
    DatasheetFontName ="Tahoma"
    OnError ="[Event Procedure]"
    RecordSourceQualifier ="dbo"
    Begin
        Begin Label
            BackStyle =0
            TextFontCharSet =178
            FontName ="Tahoma"
        End
        Begin CommandButton
            TextFontCharSet =178
            Width =1701
            Height =283
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
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
            Height =5792
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
                    ColumnWidth =675
                    ColumnOrder =1
                    Name ="GCTABID"
                    ControlSource ="GCTABID"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =120
                            Width =960
                            Height =240
                            Name ="Label1"
                            Caption ="شماره "
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
                    ColumnWidth =8610
                    ColumnOrder =2
                    TabIndex =1
                    Name ="GCNAME"
                    ControlSource ="GCNAME"
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
                    TextFontCharSet =0
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1977
                    Top =793
                    ColumnWidth =1200
                    ColumnOrder =3
                    TabIndex =2
                    Name ="GCZARIB"
                    ControlSource ="GCZARIB"
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
                Begin Subform
                    OverlapFlags =85
                    Left =170
                    Top =2381
                    Width =5381
                    Height =2880
                    TabIndex =3
                    Name ="GRADE_GRP_FORM"
                    SourceObject ="Form.GRADE_GRP_FORM"
                    LinkChildFields ="GCTABID"
                    LinkMasterFields ="GCTABID"
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1977
                    Top =1155
                    ColumnWidth =1444
                    TabIndex =4
                    Name ="GCCUST_HES"
                    ControlSource ="GCCUST_HES"
                    ValidationRule ="Not Is Null"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =170
                            Top =1157
                            Width =870
                            Height =240
                            Name ="Label10"
                            Caption ="مشتري"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1977
                    Top =1517
                    ColumnWidth =1275
                    TabIndex =5
                    Name ="GCDATE"
                    ControlSource ="GCDATE"
                    ValidationRule ="Not Is Null"
                    DefaultValue ="=farsidate(Date())"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =170
                            Top =1519
                            Width =870
                            Height =240
                            Name ="Label12"
                            Caption ="تاريخ ثبت"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1977
                    Top =1879
                    ColumnWidth =2370
                    TabIndex =6
                    Name ="USERNAME"
                    ControlSource ="USERNAME"
                    ValidationRule ="Not Is Null"
                    DefaultValue ="UCurrentUser()"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =170
                            Top =1881
                            Width =870
                            Height =240
                            Name ="Label14"
                            Caption ="نام کاربر"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "GRADE_TAB_FORM.cls"
