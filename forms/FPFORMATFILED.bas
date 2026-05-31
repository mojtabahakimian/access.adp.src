Version =20
VersionRequired =20
Begin Form
    AllowDeletions = NotDefault
    AllowAdditions = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =2
    AllowUpdating =4
    ViewsAllowed =2
    TabularCharSet =178
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =4940
    RowHeight =285
    ItemSuffix =48
    Left =405
    Top =1590
    Right =14760
    Bottom =9075
    DatasheetGridlinesColor =12632256
    MaxRecords =100000
    RecSrcDt = Begin
        0xacf23e2333ffe240
    End
    RecordSource ="FPFORMATFILELD"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Tahoma"
    AllowFormView =0
    RecordSourceQualifier ="dbo"
    Begin
        Begin Label
            BackStyle =0
            TextFontCharSet =178
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
                    ColumnOrder =0
                    Name ="FPID"
                    ControlSource ="FPID"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =113
                            Width =420
                            Height =240
                            Name ="Label1"
                            Caption ="FPID"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1927
                    Top =1814
                    Width =2496
                    ColumnWidth =915
                    ColumnOrder =4
                    TabIndex =2
                    Name ="FPTOP"
                    ControlSource ="FPTOP"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =120
                            Top =1815
                            Width =810
                            Height =240
                            Name ="Label11"
                            Caption ="فاصه از بالا"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1927
                    Top =2154
                    Width =2496
                    ColumnWidth =945
                    ColumnOrder =5
                    TabIndex =3
                    Name ="FPLEFT"
                    ControlSource ="FPLEFT"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =120
                            Top =2160
                            Width =930
                            Height =240
                            Name ="Label13"
                            Caption ="فاصله از چپ"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1927
                    Top =2494
                    Width =2496
                    ColumnWidth =465
                    ColumnOrder =6
                    TabIndex =4
                    Name ="FPWITH"
                    ControlSource ="FPWITH"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =113
                            Top =2494
                            Width =660
                            Height =240
                            Name ="Label15"
                            Caption ="طول"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1927
                    Top =2834
                    Width =2496
                    ColumnWidth =555
                    ColumnOrder =7
                    TabIndex =5
                    Name ="FPHIGHT"
                    ControlSource ="FPHIGHT"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =2834
                            Width =720
                            Height =240
                            Name ="Label17"
                            Caption ="عرض"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1927
                    Top =3514
                    Width =2496
                    ColumnWidth =930
                    ColumnOrder =9
                    TabIndex =7
                    Name ="FPFONTSIZE"
                    ControlSource ="FPFONTSIZE"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =3510
                            Width =1005
                            Height =240
                            Name ="Label21"
                            Caption ="اندازه فونت"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1927
                    Top =4081
                    ColumnWidth =885
                    ColumnOrder =14
                    TabIndex =8
                    Name ="FPFONTCOLOR"
                    ControlSource ="FPFONTCOLOR"
                    OnDblClick ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =4081
                            Width =1185
                            Height =240
                            Name ="Label29"
                            Caption ="رنگ متن"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1927
                    Top =4421
                    ColumnWidth =795
                    ColumnOrder =15
                    TabIndex =9
                    Name ="FPBKCOLOR"
                    ControlSource ="FPBKCOLOR"
                    OnDblClick ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =4421
                            Width =960
                            Height =240
                            Name ="Label31"
                            Caption ="رنگ زمينه"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1927
                    Top =4761
                    Width =2490
                    ColumnWidth =1995
                    ColumnOrder =16
                    TabIndex =10
                    Name ="FPFORMAT"
                    ControlSource ="FPFORMAT"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =4761
                            Width =885
                            Height =240
                            Name ="Label33"
                            Caption ="فرمت"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1927
                    Top =3174
                    Width =2490
                    ColumnWidth =2055
                    ColumnOrder =8
                    TabIndex =6
                    Name ="FPFONT"
                    ControlSource ="FPFONT"
                    RowSourceType ="Value List"
                    DefaultValue ="\"Titr\""
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =3174
                            Width =660
                            Height =240
                            Name ="Label19"
                            Caption ="فونت"
                        End
                    End
                End
                Begin ComboBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =1927
                    Top =453
                    ColumnWidth =3945
                    ColumnOrder =1
                    TabIndex =1
                    Name ="FPFIELD"
                    ControlSource ="FPFIELD"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT FPFLDID, FPFILEDNAME FROM FPFILEDS"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =453
                            Width =675
                            Height =240
                            Name ="Label3"
                            Caption ="نام فيلد"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    ReadingOrder =1
                    Left =1214
                    Top =1190
                    ColumnWidth =825
                    ColumnOrder =3
                    TabIndex =11
                    Name ="FPVISIBLE"
                    ControlSource ="FPVISIBLE"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =555
                            Top =1155
                            Width =720
                            Height =240
                            Name ="Label35"
                            Caption ="چاپ شود"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    ReadingOrder =1
                    Left =1214
                    Top =1585
                    ColumnWidth =570
                    ColumnOrder =10
                    TabIndex =12
                    Name ="FPBOLD"
                    ControlSource ="FPBOLD"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =630
                            Top =1550
                            Width =645
                            Height =240
                            Name ="Label37"
                            Caption ="پررنگ"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    ReadingOrder =1
                    Left =1214
                    Top =1980
                    ColumnWidth =615
                    ColumnOrder =11
                    TabIndex =13
                    Name ="FPITALIC"
                    ControlSource ="FPITALIC"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =555
                            Top =1945
                            Width =720
                            Height =240
                            Name ="Label39"
                            Caption ="ايتاليك"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    ReadingOrder =1
                    Left =1214
                    Top =2375
                    ColumnWidth =630
                    ColumnOrder =12
                    TabIndex =14
                    Name ="FPUNDELINE"
                    ControlSource ="FPUNDELINE"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =555
                            Top =2340
                            Width =720
                            Height =240
                            Name ="Label41"
                            Caption ="زير خط"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1920
                    Top =1477
                    Width =2511
                    ColumnWidth =2145
                    ColumnOrder =2
                    TabIndex =15
                    Name ="FPCAPTION"
                    ControlSource ="FPCAPTION"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =427
                            Top =793
                            Width =990
                            Height =240
                            Name ="Label42"
                            Caption ="عنوان"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =1814
                    Top =5272
                    Width =1710
                    ColumnWidth =1185
                    ColumnOrder =13
                    TabIndex =16
                    Name ="FPTEXTALIGHN"
                    ControlSource ="FPTEXTALIGHN"
                    RowSourceType ="Value List"
                    RowSource ="0;\"معمولي\";1;\"چپ چين\";2;\"وسط چين\";3;\"راست چين\";4;\"جدا جدا\""
                    ColumnWidths ="0"
                    DefaultValue ="\"Titr\""
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Top =5265
                            Width =930
                            Height =240
                            Name ="Label46"
                            Caption ="جهت نوشتن"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    ReadingOrder =1
                    Left =2121
                    Top =1020
                    TabIndex =17
                    Name ="KADR"
                    ControlSource ="KADR"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =1671
                            Top =990
                            Width =480
                            Height =240
                            Name ="Label47"
                            Caption ="كادر"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "FPFORMATFILED.cls"
