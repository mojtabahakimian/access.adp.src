Version =20
VersionRequired =20
Begin Form
    AutoCenter = NotDefault
    AllowDeletions = NotDefault
    AllowAdditions = NotDefault
    KeyPreview = NotDefault
    AllowEdits = NotDefault
    DefaultView =2
    AllowUpdating =4
    ViewsAllowed =2
    TabularCharSet =178
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =4940
    RowHeight =360
    DatasheetFontHeight =10
    ItemSuffix =124
    Left =5505
    Top =1095
    Right =13440
    Bottom =5085
    DatasheetForeColor =33554432
    DatasheetGridlinesColor =12632256
    MaxRecords =100000000
    RecSrcDt = Begin
        0x8c0eaa49acb4e540
    End
    RecordSource ="BUGET_MAIN"
    Caption ="ليست بودجه هاي گرفته شده"
    OnCurrent ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Tahoma"
    OnKeyPress ="[Event Procedure]"
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
        Begin ComboBox
            SpecialEffect =2
            Width =1701
            LabelX =-1701
            FontName ="Tahoma"
        End
        Begin Section
            Height =1587
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1927
                    Top =113
                    Width =2490
                    ColumnWidth =720
                    Name ="BGID"
                    ControlSource ="BGID"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =120
                            Top =120
                            Width =525
                            Height =240
                            Name ="Label1"
                            Caption ="شماره"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1927
                    Top =453
                    Width =2490
                    TabIndex =1
                    Name ="BGCDATE"
                    ControlSource ="BGCDATE"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =450
                            Width =825
                            Height =240
                            Name ="Label3"
                            Caption ="تاريخ بودجه"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1927
                    Top =1133
                    TabIndex =3
                    Name ="BGSAL"
                    ControlSource ="BGSAL"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =1140
                            Width =825
                            Height =240
                            Name ="Label7"
                            Caption ="سال بودجه"
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
                    Top =793
                    TabIndex =2
                    Name ="BGMAH"
                    ControlSource ="BGMAH"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT MON_ID, MON FROM MON"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =795
                            Width =720
                            Height =240
                            Name ="Label5"
                            Caption ="ماه بودجه"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =215
                    TextAlign =3
                    IMESentenceMode =3
                    ColumnWidth =2085
                    TabIndex =4
                    Name ="BGBUGETMON"
                    ControlSource ="BGBUGETMON"
                    Format ="Standard"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =2247
                            Width =1155
                            Height =240
                            Name ="Label124"
                            Caption ="مبلغ ورود  جريان نقد"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "BUGET_MAIN_list.cls"
