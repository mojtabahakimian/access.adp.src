Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    DividingLines = NotDefault
    KeyPreview = NotDefault
    DefaultView =0
    AllowUpdating =4
    TabularCharSet =178
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =14966
    RowHeight =420
    ItemSuffix =23
    Left =2520
    Top =615
    Right =17490
    Bottom =9570
    DatasheetForeColor =33554432
    DatasheetGridlinesColor =12632256
    MaxRecords =1000000
    AfterInsert ="[Event Procedure]"
    RecSrcDt = Begin
        0x9683010b5ef6e440
    End
    RecordSource ="TOZIE"
    Caption ="توزيع دستي"
    OnCurrent ="[Event Procedure]"
    BeforeUpdate ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Tahoma"
    OnKeyPress ="[Event Procedure]"
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
        Begin CheckBox
            SpecialEffect =2
            LabelX =230
            LabelY =-30
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
            Height =8957
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =12870
                    Top =165
                    Width =1020
                    Height =317
                    ColumnWidth =960
                    ColumnOrder =0
                    BackColor =13408767
                    Name ="TID"
                    ControlSource ="TID"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =13952
                            Top =165
                            Width =960
                            Height =317
                            Name ="Label1"
                            Caption ="شماره برگه"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =9921
                    Top =1020
                    Width =1266
                    Height =317
                    ColumnWidth =1980
                    ColumnOrder =4
                    TabIndex =6
                    Name ="CDATE"
                    ControlSource ="CDATE"
                    DefaultValue ="Now()"
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =85
                            Left =11283
                            Top =1020
                            Width =750
                            Height =240
                            Name ="Label11"
                            Caption ="تاريخ ايجاد"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =9469
                    Top =623
                    Width =2250
                    Height =317
                    ColumnWidth =3270
                    ColumnOrder =6
                    TabIndex =7
                    Name ="USER_NAME"
                    ControlSource ="USER_NAME"
                    DefaultValue ="UCurrentUser()"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =11779
                            Top =623
                            Width =345
                            Height =330
                            Name ="Label13"
                            Caption ="كاربر"
                        End
                    End
                End
                Begin ComboBox
                    AutoExpand = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =6180
                    Top =165
                    Width =3021
                    Height =317
                    ColumnWidth =3795
                    ColumnOrder =5
                    TabIndex =2
                    BackColor =13434828
                    ForeColor =8388608
                    Name ="TDRIVER"
                    ControlSource ="TDRIVER"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TDRIVER FROM TOZIE GROUP BY TDRIVER ORDER BY TDRIVER"
                    BeforeUpdate ="[Event Procedure]"
                    OnNotInList ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =9240
                            Top =165
                            Width =840
                            Height =345
                            ForeColor =8388608
                            Name ="Label21"
                            Caption ="نام راننده"
                        End
                    End
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =279
                    Top =165
                    Width =4875
                    Height =317
                    ColumnWidth =3660
                    ColumnOrder =2
                    TabIndex =3
                    Name ="TMAMUR"
                    ControlSource ="TMAMUR"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TMAMUR FROM TOZIE GROUP BY TMAMUR ORDER BY TMAMUR"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5211
                            Top =165
                            Width =795
                            Height =285
                            Name ="Label5"
                            Caption ="مامور توزيع"
                        End
                    End
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =12188
                    Top =623
                    Width =1980
                    Height =317
                    ColumnWidth =3075
                    ColumnOrder =3
                    TabIndex =4
                    Name ="TCITY"
                    ControlSource ="TCITY"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TCITY FROM TOZIE GROUP BY TCITY ORDER BY TCITY"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =14389
                            Top =627
                            Width =525
                            Height =345
                            Name ="Label7"
                            Caption ="شهر"
                        End
                    End
                End
                Begin Subform
                    OverlapFlags =87
                    Left =61
                    Top =1417
                    Width =14850
                    Height =7440
                    TabIndex =8
                    Name ="TOZIE_SUB"
                    SourceObject ="Form.TOZIE_SUB"
                    LinkChildFields ="TID"
                    LinkMasterFields ="TID"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =12991
                            Top =1177
                            Width =1635
                            Height =240
                            Name ="Visit_route_dtl_FORM Label"
                            Caption ="مشتريان زير مجموعه"
                            EventProcPrefix ="Visit_route_dtl_FORM_Label"
                        End
                    End
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    AccessKey =134
                    TextFontCharSet =0
                    Left =340
                    Top =623
                    Width =1020
                    Height =465
                    FontWeight =700
                    TabIndex =9
                    Name ="Command100"
                    Caption ="&چاپ "
                    OnClick ="[Event Procedure]"
                    UnicodeAccessKey =1670
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =10261
                    Top =165
                    Width =1191
                    Height =317
                    ColumnWidth =840
                    ColumnOrder =1
                    TabIndex =1
                    BackColor =13434828
                    ForeColor =8388608
                    Name ="TDATE"
                    ControlSource ="TDATE"
                    AfterUpdate ="[Event Procedure]"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =11535
                            Top =165
                            Width =840
                            Height =345
                            ForeColor =8388608
                            Name ="Label19"
                            Caption ="تاريخ برگه"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontCharSet =0
                    Left =2777
                    Top =623
                    Width =4185
                    Height =465
                    FontWeight =700
                    TabIndex =5
                    Name ="Command23"
                    Caption ="دريافت ليست فاكتورها براي حواله"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    Left =1644
                    Top =623
                    Width =1020
                    Height =465
                    FontWeight =700
                    TabIndex =10
                    Name ="Command22"
                    Caption ="چاپ  حواله"
                    OnClick ="[Event Procedure]"
                End
            End
        End
    End
End
CodeBehindForm
' See "tozie.cls"
