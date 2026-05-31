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
    Width =7510
    RowHeight =315
    ItemSuffix =16
    Left =510
    Top =3015
    Right =5610
    Bottom =5610
    DatasheetForeColor =33554432
    DatasheetGridlinesColor =12632256
    MaxRecords =1000000
    RecSrcDt = Begin
        0x2e578e221e8ee540
    End
    RecordSource ="GRADE_CUST_GRP"
    Caption ="معيار سنجش"
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
        Begin BoundObjectFrame
            SpecialEffect =2
            Width =4536
            Height =2835
            LabelX =-1701
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
            Height =5953
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =3629
                    Top =113
                    ColumnOrder =1
                    Name ="GCTABID"
                    ControlSource ="GCTABID"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1822
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
                    Left =3629
                    Top =453
                    Width =2475
                    ColumnWidth =4785
                    ColumnOrder =2
                    TabIndex =1
                    Name ="GCGRPNAME"
                    ControlSource ="GCGRPNAME"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1822
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
                    Left =3629
                    Top =793
                    ColumnOrder =6
                    TabIndex =2
                    Name ="GCGRPZARIB"
                    ControlSource ="GCGRPZARIB"
                    ValidationRule ="Not Is Null"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1822
                            Top =795
                            Width =870
                            Height =240
                            Name ="Label5"
                            Caption ="ضريب "
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =3629
                    Top =1133
                    ColumnWidth =930
                    ColumnOrder =0
                    TabIndex =3
                    Name ="GCGRPID"
                    ControlSource ="GCGRPID"
                    ValidationRule ="Not Is Null"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1822
                            Top =1140
                            Width =870
                            Height =240
                            Name ="Label7"
                            Caption ="شماره"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =3629
                    Top =1590
                    ColumnOrder =5
                    TabIndex =4
                    Name ="GCGRPGRADE"
                    ControlSource ="GCGRPGRADE"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1822
                            Top =1592
                            Width =870
                            Height =240
                            Name ="Label10"
                            Caption ="امتياز"
                        End
                    End
                End
                Begin ComboBox
                    ColumnHidden = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =3509
                    Top =1927
                    ColumnWidth =2070
                    ColumnOrder =3
                    TabIndex =5
                    Name ="GCVALUE"
                    ControlSource ="GCVALUE"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT GSCACOD, GSCANAME FROM GSCALE"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1702
                            Top =1929
                            Width =870
                            Height =240
                            Name ="Label12"
                            Caption ="ارزش"
                        End
                    End
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =3629
                    Top =2265
                    ColumnOrder =4
                    TabIndex =6
                    Name ="GCVALUESCAL"
                    ControlSource ="GCVALUESCAL"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT GSCADTCOD, GSCANAME FROM GSCADTL"
                    ColumnWidths ="0"
                    AfterUpdate ="[Event Procedure]"
                    OnEnter ="[Event Procedure]"
                    OnExit ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1822
                            Top =2267
                            Width =870
                            Height =240
                            Name ="Label14"
                            Caption ="مقدار"
                        End
                    End
                End
                Begin BoundObjectFrame
                    OverlapFlags =85
                    Top =3118
                    TabIndex =7
                    Name ="GCPS"
                    ControlSource ="GCPS"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5715
                            Top =3120
                            Width =570
                            Height =240
                            Name ="Label15"
                            Caption ="مدارک"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "GRADE_GRP_FORM.cls"
