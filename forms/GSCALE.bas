Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    DividingLines = NotDefault
    DefaultView =0
    AllowUpdating =4
    TabularCharSet =178
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =9758
    RowHeight =345
    ItemSuffix =14
    Left =270
    Top =600
    Right =10875
    Bottom =8415
    DatasheetForeColor =33554432
    DatasheetGridlinesColor =12632256
    MaxRecords =1000000
    RecSrcDt = Begin
        0xe4b2d0d9b88ee540
    End
    RecordSource ="GSCALE"
    Caption ="تعريف پارامترهاي کمي و کيفي رده بندي مشتريان"
    OnCurrent ="[Event Procedure]"
    BeforeUpdate ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
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
            CanGrow = NotDefault
            Height =8109
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6065
                    Top =113
                    Width =1521
                    Height =340
                    ColumnWidth =855
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
                    Left =1950
                    Top =510
                    Width =5670
                    Height =340
                    ColumnWidth =8130
                    TabIndex =1
                    BackColor =13434828
                    Name ="GSCANAME"
                    ControlSource ="GSCANAME"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =7666
                            Top =510
                            Width =795
                            Height =240
                            Name ="Label3"
                            Caption ="عنوان"
                        End
                    End
                End
                Begin Subform
                    OverlapFlags =85
                    Left =170
                    Top =1303
                    Width =9135
                    Height =6465
                    TabIndex =3
                    Name ="GSCALE_DTL"
                    SourceObject ="Form.GSCALE_DTL"
                    LinkChildFields ="GSCACOD"
                    LinkMasterFields ="GSCACOD"
                End
                Begin ComboBox
                    RowSourceTypeInt =1
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =6105
                    Top =915
                    Width =1521
                    Height =340
                    TabIndex =2
                    Name ="GSCAKIND"
                    ControlSource ="GSCAKIND"
                    RowSourceType ="Value List"
                    RowSource ="0;\"کمي\";1;\"کيفي\""
                    ColumnWidths ="0"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =7681
                            Top =915
                            Width =885
                            Height =240
                            Name ="Label5"
                            Caption ="نوع"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "GSCALE.cls"
