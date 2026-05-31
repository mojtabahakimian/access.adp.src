Version =20
VersionRequired =20
Begin Form
    AutoCenter = NotDefault
    DefaultView =2
    AllowUpdating =4
    ViewsAllowed =2
    TabularCharSet =178
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =5100
    RowHeight =360
    ItemSuffix =10
    Left =570
    Right =6540
    Bottom =6300
    DatasheetForeColor =33554432
    DatasheetGridlinesColor =12632256
    MaxRecords =0
    RecSrcDt = Begin
        0xd9a9cc679205e540
    End
    RecordSource ="CUSTKIND_TF"
    Caption ="CUSTKIND_TF"
    DatasheetFontName ="Tahoma"
    OnError ="[Event Procedure]"
    AllowFormView =0
    RecordSourceQualifier ="dbo"
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
        End
        Begin CommandButton
            Width =1701
            Height =283
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Tahoma"
        End
        Begin OptionButton
            SpecialEffect =2
            LabelX =230
            LabelY =-30
        End
        Begin CheckBox
            SpecialEffect =2
            LabelX =230
            LabelY =-30
        End
        Begin OptionGroup
            SpecialEffect =3
            Width =1701
            Height =1701
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
            Width =1701
            LabelX =-1701
            FontName ="Tahoma"
            AsianLineBreak =255
        End
        Begin ListBox
            SpecialEffect =2
            Width =1701
            Height =1417
            LabelX =-1701
            FontName ="Tahoma"
        End
        Begin ComboBox
            SpecialEffect =2
            Width =1701
            LabelX =-1701
            FontName ="Tahoma"
        End
        Begin Section
            Height =1870
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1312
                    Top =453
                    TabIndex =1
                    Name ="TID"
                    ControlSource ="TID"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =4500
                            Top =450
                            Width =600
                            Height =405
                            Name ="Label3"
                            Caption ="كد تخفيف"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =737
                    Top =793
                    TabIndex =2
                    Name ="CREA"
                    ControlSource ="CREA"
                    DefaultValue ="Now()"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =3775
                            Top =795
                            Width =750
                            Height =240
                            Name ="Label5"
                            Caption ="تاريخ ايجاد"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    ReadingOrder =1
                    Left =1951
                    Top =1247
                    TabIndex =3
                    Name ="ACTIV"
                    ControlSource ="ACTIV"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =1380
                            Top =1215
                            Width =675
                            Height =240
                            Name ="Label9"
                            Caption ="غير فعال"
                        End
                    End
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =1312
                    Top =113
                    ColumnWidth =3975
                    BackColor =13434828
                    Name ="CUST_COD"
                    ControlSource ="CUST_COD"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT CUST_COD, CUSTKNAME FROM CUSTKIND ORDER BY CUSTKNAME"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3945
                            Top =120
                            Width =1065
                            Height =240
                            Name ="Label1"
                            Caption =" گروه مشتريان"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "CUSTKIND_TF_FORM.cls"
