Version =20
VersionRequired =20
Begin Form
    AutoCenter = NotDefault
    DefaultView =2
    AllowUpdating =4
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =4940
    RowHeight =330
    ItemSuffix =37
    Left =555
    Top =3975
    Right =15120
    Bottom =10440
    DatasheetGridlinesColor =12632256
    MaxRecords =100000000
    RecSrcDt = Begin
        0x1eb799ef77c1e540
    End
    RecordSource ="ENHESAR_KALA"
    Caption ="کالاهاي انحصاري"
    BeforeInsert ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Tahoma"
    OnError ="[Event Procedure]"
    RecordSourceQualifier ="dbo"
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
        End
        Begin CheckBox
            SpecialEffect =2
            LabelX =230
            LabelY =-30
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
            Height =5915
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
                    ColumnWidth =630
                    ColumnOrder =0
                    Name ="EN_IDD"
                    ControlSource ="EN_IDD"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =113
                            Width =630
                            Height =240
                            Name ="Label1"
                            Caption ="شماره"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1927
                    Top =453
                    Width =2490
                    ColumnWidth =945
                    ColumnOrder =1
                    TabIndex =1
                    Name ="EN_CUST_CO"
                    ControlSource ="EN_CUST_CO"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =453
                            Width =1065
                            Height =240
                            Name ="Label3"
                            Caption ="کد مشتري"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1927
                    Top =5215
                    ColumnWidth =690
                    ColumnOrder =13
                    TabIndex =13
                    Name ="USERID"
                    ControlSource ="USERID"
                    DefaultValue ="[Forms]![baseknow]![USERCOD]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =5215
                            Width =630
                            Height =240
                            Name ="Label31"
                            Caption ="کد کاربر"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1927
                    Top =5555
                    ColumnWidth =2190
                    ColumnOrder =16
                    TabIndex =14
                    Name ="EN_CDATE"
                    ControlSource ="EN_CDATE"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =5555
                            Width =855
                            Height =240
                            Name ="Label33"
                            Caption ="تاريخ ايجاد"
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
                    Top =1814
                    ColumnWidth =1470
                    ColumnOrder =2
                    TabIndex =3
                    Name ="EN_GRCODE1"
                    ControlSource ="EN_GRCODE1"
                    RowSourceType ="Table/View/StoredProc"
                    ColumnWidths ="0"
                    BeforeUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =1814
                            Width =1080
                            Height =240
                            Name ="LCOL1"
                            Caption ="کد گروه?"
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
                    Top =2154
                    ColumnOrder =3
                    TabIndex =4
                    Name ="EN_GRCODE2"
                    ControlSource ="EN_GRCODE2"
                    RowSourceType ="Table/View/StoredProc"
                    ColumnWidths ="0"
                    BeforeUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =2154
                            Width =1080
                            Height =240
                            Name ="LCOL2"
                            Caption ="کد گروه2"
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
                    Top =2494
                    ColumnOrder =4
                    TabIndex =5
                    Name ="EN_GRCODE3"
                    ControlSource ="EN_GRCODE3"
                    RowSourceType ="Table/View/StoredProc"
                    ColumnWidths ="0"
                    BeforeUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =2494
                            Width =1080
                            Height =240
                            Name ="LCOL3"
                            Caption ="کد گروه?"
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
                    Top =2834
                    ColumnOrder =5
                    TabIndex =6
                    Name ="EN_GRCODE4"
                    ControlSource ="EN_GRCODE4"
                    RowSourceType ="Table/View/StoredProc"
                    ColumnWidths ="0"
                    BeforeUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =2834
                            Width =1080
                            Height =240
                            Name ="LCOL4"
                            Caption ="کد گروه?"
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
                    Top =3174
                    ColumnOrder =6
                    TabIndex =7
                    Name ="EN_GRCODE5"
                    ControlSource ="EN_GRCODE5"
                    RowSourceType ="Table/View/StoredProc"
                    ColumnWidths ="0"
                    BeforeUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =3174
                            Width =1080
                            Height =240
                            Name ="LCOL5"
                            Caption ="کد گروه?"
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
                    Top =3514
                    ColumnOrder =7
                    TabIndex =8
                    Name ="EN_GRCODE6"
                    ControlSource ="EN_GRCODE6"
                    RowSourceType ="Table/View/StoredProc"
                    ColumnWidths ="0"
                    BeforeUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =3514
                            Width =1080
                            Height =240
                            Name ="LCOL6"
                            Caption ="کد گروه?"
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
                    Top =3855
                    ColumnOrder =8
                    TabIndex =9
                    Name ="EN_GRCODE7"
                    ControlSource ="EN_GRCODE7"
                    RowSourceType ="Table/View/StoredProc"
                    ColumnWidths ="0"
                    BeforeUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =3855
                            Width =1080
                            Height =240
                            Name ="LCOL7"
                            Caption ="کد گروه?"
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
                    Top =4195
                    ColumnOrder =9
                    TabIndex =10
                    Name ="EN_GRCODE8"
                    ControlSource ="EN_GRCODE8"
                    RowSourceType ="Table/View/StoredProc"
                    ColumnWidths ="0"
                    BeforeUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =4195
                            Width =1080
                            Height =240
                            Name ="LCOL8"
                            Caption ="کد گروه?"
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
                    Top =4535
                    ColumnOrder =10
                    TabIndex =11
                    Name ="EN_GRCODE9"
                    ControlSource ="EN_GRCODE9"
                    RowSourceType ="Table/View/StoredProc"
                    ColumnWidths ="0"
                    BeforeUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =4535
                            Width =1080
                            Height =240
                            Name ="LCOL9"
                            Caption ="کد گروه?"
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
                    Top =4875
                    ColumnOrder =11
                    TabIndex =12
                    Name ="EN_GRCODE10"
                    ControlSource ="EN_GRCODE10"
                    RowSourceType ="Table/View/StoredProc"
                    ColumnWidths ="0"
                    BeforeUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =4875
                            Width =1170
                            Height =240
                            Name ="LCOL10"
                            Caption ="کد گروه??"
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
                    Top =1133
                    Width =2490
                    ColumnWidth =4650
                    ColumnOrder =12
                    TabIndex =2
                    Name ="EN_KALA_COD"
                    ControlSource ="EN_KALA_COD"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT CODE, NAME FROM STUF_DEF ORDER BY NAME"
                    ColumnWidths ="0"
                    OnNotInList ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =1133
                            Width =1155
                            Height =240
                            Name ="Label7"
                            Caption =" کالا"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "ENHESAR_KALA_FORM.cls"
