Version =20
VersionRequired =20
Begin Form
    AllowDeletions = NotDefault
    KeyPreview = NotDefault
    DefaultView =2
    AllowUpdating =4
    ViewsAllowed =2
    TabularCharSet =178
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =7036
    RowHeight =630
    ItemSuffix =12
    Left =600
    Top =540
    Right =15345
    Bottom =7875
    DatasheetForeColor =33554432
    DatasheetGridlinesColor =12632256
    MaxRecords =10000000
    RecSrcDt = Begin
        0x9223ec85bceae340
    End
    RecordSource ="SQLSTATE"
    Caption ="ليست گزارشات ساخته شده توط كاربران"
    OnCurrent ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Tahoma"
    OnKeyPress ="[Event Procedure]"
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
        Begin ComboBox
            SpecialEffect =2
            TextFontCharSet =178
            Width =1701
            LabelX =-1701
            FontName ="Tahoma"
        End
        Begin Section
            Height =2741
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1927
                    Top =113
                    ColumnWidth =630
                    Name ="IDD"
                    ControlSource ="IDD"
                    Begin
                        Begin Label
                            OverlapFlags =85
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
                    ColumnHidden = NotDefault
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1927
                    Top =453
                    Width =4965
                    Height =810
                    ColumnWidth =14085
                    TabIndex =1
                    Name ="SQLST"
                    ControlSource ="SQLST"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =450
                            Width =1185
                            Height =240
                            Name ="Label3"
                            Caption ="SQL STATMENT"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1927
                    Top =1360
                    Width =2490
                    ColumnWidth =7095
                    TabIndex =2
                    Name ="TITEL"
                    ControlSource ="TITEL"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =1365
                            Width =960
                            Height =240
                            Name ="Label5"
                            Caption ="عنوان گزارش"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1927
                    Top =1700
                    Width =2490
                    ColumnWidth =1950
                    TabIndex =3
                    Name ="USER_NAME"
                    ControlSource ="USER_NAME"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =1700
                            Width =975
                            Height =240
                            Name ="Label7"
                            Caption ="كاربر سازنده"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1927
                    Top =2381
                    ColumnWidth =2085
                    TabIndex =5
                    Name ="CR_DATE"
                    ControlSource ="CR_DATE"
                    Format ="General Date"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =2381
                            Width =765
                            Height =240
                            Name ="Label11"
                            Caption ="زمان ايجاد"
                        End
                    End
                End
                Begin ComboBox
                    Locked = NotDefault
                    RowSourceTypeInt =1
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =1927
                    Top =2040
                    TabIndex =4
                    Name ="LETOTHER"
                    ControlSource ="LETOTHER"
                    RowSourceType ="Value List"
                    RowSource ="1;\"بله\";0;\"نه\""
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =2040
                            Width =1245
                            Height =240
                            Name ="Label9"
                            Caption ="اجازه براي ديگران"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "SQLSTATEFORM.cls"
