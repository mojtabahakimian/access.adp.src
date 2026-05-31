Version =20
VersionRequired =20
Begin Form
    AllowDeletions = NotDefault
    KeyPreview = NotDefault
    DefaultView =2
    AllowUpdating =4
    ViewsAllowed =2
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =5669
    RowHeight =330
    ItemSuffix =18
    Left =2730
    Top =1125
    Right =10170
    Bottom =4620
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x7ff3cef18e4be440
    End
    RecordSource ="SELECT USERCO, PRNAM, PCOUNT, ACT, IDP, PCD, REPNAM FROM USER_PRINT"
    BeforeUpdate ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Tahoma"
    OnError ="[Event Procedure]"
    Moveable =0
    AllowFormView =0
    FetchDefaults =0
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
            TextFontCharSet =178
            Width =1701
            LabelX =-1701
            FontName ="Tahoma"
        End
        Begin Section
            Height =2867
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2490
                    Top =1247
                    ColumnWidth =675
                    ColumnOrder =0
                    TabIndex =1
                    Name ="IDP"
                    ControlSource ="IDP"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =793
                            Top =1251
                            Width =570
                            Height =240
                            Name ="Label4"
                            Caption ="شماره "
                        End
                    End
                End
                Begin ComboBox
                    RowSourceTypeInt =1
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =1
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2325
                    Top =340
                    ColumnWidth =4470
                    ColumnOrder =1
                    Name ="PRNAM"
                    ControlSource ="PRNAM"
                    RowSourceType ="Value List"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =623
                            Top =340
                            Width =930
                            Height =240
                            Name ="Label2"
                            Caption ="نام چاپگر"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =2495
                    Top =1814
                    ColumnWidth =1020
                    ColumnOrder =4
                    TabIndex =2
                    Name ="ACT"
                    ControlSource ="ACT"
                    RowSourceType ="Value List"
                    RowSource ="0;\"فعال\";1;\"غير فعال\""
                    ColumnWidths ="0"
                    DefaultValue ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =793
                            Top =1814
                            Width =585
                            Height =240
                            Name ="Label13"
                            Caption ="وضعيت"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1870
                    Top =2267
                    ColumnWidth =780
                    ColumnOrder =2
                    TabIndex =3
                    Name ="PCOUNT"
                    ControlSource ="PCOUNT"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4522
                            Top =2267
                            Width =750
                            Height =240
                            Name ="Label14"
                            Caption ="تعدا چاپ"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =453
                    Top =793
                    ColumnWidth =480
                    ColumnOrder =3
                    TabIndex =4
                    Name ="USERCO"
                    ControlSource ="USERCO"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3105
                            Top =793
                            Width =750
                            Height =240
                            Name ="Label15"
                            Caption ="كاربر"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1870
                    Top =2627
                    ColumnWidth =960
                    TabIndex =5
                    Name ="REPNAM"
                    ControlSource ="REPNAM"
                    DefaultValue ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4515
                            Top =2625
                            Width =855
                            Height =240
                            Name ="Label17"
                            Caption ="شماره چاپ"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "user_priter_sub.cls"
