Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    DividingLines = NotDefault
    DefaultView =0
    AllowUpdating =4
    ScrollBars =0
    ViewsAllowed =1
    TabularCharSet =178
    TabularFamily =0
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =14853
    DatasheetFontHeight =10
    ItemSuffix =25
    Left =900
    Top =270
    Right =15495
    Bottom =9060
    DatasheetGridlinesColor =12632256
    MaxRecords =100000
    RecSrcDt = Begin
        0x8a497e2733ffe240
    End
    RecordSource ="FPFORMAT"
    Caption ="تهيه و نتظيم چاپ فاكتور دلخواه"
    BeforeUpdate ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial"
    OnError ="[Event Procedure]"
    AllowDatasheetView =0
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
        Begin UnboundObjectFrame
            SpecialEffect =2
            OldBorderStyle =1
            Width =4536
            Height =2835
        End
        Begin Section
            Height =8787
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =13144
                    Top =113
                    Width =918
                    Height =345
                    Name ="FPID"
                    ControlSource ="FPID"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =14075
                            Top =113
                            Width =630
                            Height =240
                            Name ="Label1"
                            Caption ="شماره :"
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
                    Left =7927
                    Top =113
                    Width =3678
                    Height =345
                    TabIndex =1
                    Name ="FPONVAN"
                    ControlSource ="FPONVAN"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =11662
                            Top =117
                            Width =1365
                            Height =330
                            Name ="Label3"
                            Caption ="عنوان فرمت چاپي:"
                        End
                    End
                End
                Begin Subform
                    OverlapFlags =119
                    Left =60
                    Top =959
                    Width =14636
                    Height =7770
                    TabIndex =4
                    Name ="FPFORMATFILED"
                    SourceObject ="Form.FPFORMATFILED"
                    LinkChildFields ="FPID"
                    LinkMasterFields ="FPID"
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =2938
                    Top =623
                    Width =3156
                    Height =330
                    TabIndex =3
                    Name ="FPPAPER"
                    ControlSource ="FPPAPER"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT idp, namep FROM papersize"
                    ColumnWidths ="0"
                    DefaultValue ="9"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =6166
                            Top =623
                            Width =870
                            Height =285
                            Name ="Label11"
                            Caption ="نوع كاغذ:"
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
                    Left =12092
                    Top =585
                    Width =897
                    Height =255
                    TabIndex =2
                    Name ="FPPGHHIGH"
                    ControlSource ="FPPGHHIGH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =13152
                            Top =592
                            Width =1560
                            Height =240
                            Name ="Label13"
                            Caption ="ارتفاع سربرگ فاكتور :"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =3052
                    Top =113
                    Width =1815
                    Height =405
                    TabIndex =5
                    Name ="Command16"
                    Caption ="اضافه كردن فرمت جديد"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =5087
                    Top =30
                    Width =576
                    Height =516
                    TabIndex =6
                    Name ="Command17"
                    Caption ="Command17"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000020000000200000000100080000000000000400000000000000000000 ,
                        0x0000000000000000000000000000bf0000bf000000bfbf00bf000000bf00bf00 ,
                        0xbfbf0000c0c0c000c0dcc000f0c8a40000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0xf0fbff00a4a0a000808080000000ff0000ff000000ffff00ff000000ff00ff00 ,
                        0xffff0000ffffff00070707070707070707070707070707070707070707070707 ,
                        0x0707070707070707070707070707070707070707070707070707070707070707 ,
                        0x0707070707070707070707070707070707070707070707070707070707070707 ,
                        0x0707070707070707070707070700000000000000000000000000000000000000 ,
                        0x000000000707070707070707000000f807060706070607060706070607060706 ,
                        0x070607000707070707070700fffb030000070607060706070607060706070607 ,
                        0x060706000707070707070707000000f807060706070607060706070607060706 ,
                        0x070607000707070707070707070000f806070607060706070607060706070607 ,
                        0x060706000707070707070707000000f807060706070607060706070607060706 ,
                        0x070607000707070707070700fffb030000070607060706070607060706070607 ,
                        0x060706000707070707070707000000f807060706070607060706070607060706 ,
                        0x070607000707070707070707070000f806070607060706070607060706070607 ,
                        0x060706000707070707070707000000f807060706070607060706070607060706 ,
                        0x070607000707070707070700fffb030000070607060706070607060706070607 ,
                        0x060706000707070707070707000000f807060706070607060706070607060706 ,
                        0x070607000707070707070707070000f806070607060706070607060706070607 ,
                        0x060706000707070707070707000000f807060706070607060706070607060706 ,
                        0x070607000707070707070700fffb030000070607060706070607060706070607 ,
                        0x060706000707070707070707000000f807060706070607060706060607060706 ,
                        0x070607000707070707070707070000f806070600000000000000000000000000 ,
                        0x060706000707070707070707000000f807060700ffffffffffffffffffffff00 ,
                        0x070607000707070707070700fffb030000070600ffffffffffffffffffffff00 ,
                        0x060706000707070707070707000000f807060700ffffffffffffffffffffff00 ,
                        0x070607000707070707070707070000f806070600000000000000000000000000 ,
                        0x060706000707070707070707000000f807060706070607060706070607060706 ,
                        0x070607000707070707070700fffb030000070607060706070607060706070607 ,
                        0x060706000707070707070707000000f807060706070607060706070607060706 ,
                        0x070607000707070707070707070000f806070607060706070607060706070607 ,
                        0x0607060007070707070707070700000000000000000000000000000000000000 ,
                        0x0000000007070707070707070707070707070707070707070707070707070707 ,
                        0x0707070707070707070707070707070707070707070707070707070707070707 ,
                        0x0707070707070707070707070707070707070707070707070707070707070707 ,
                        0x0707070707070707000000000000000000000000000000000000000000000000
                    End
                    ObjectPalette = Begin
                        0x0003000100000000bf00000000bf0000bfbf00000000bf00bf00bf0000bfbf00 ,
                        0xc0c0c000c0dcc000a4c8f0000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000000000000000000000000000000000000000000000000000fffbf000 ,
                        0xa0a0a40080808000ff00000000ff0000ffff00000000ff00ff00ff0000ffff00 ,
                        0xffffff0000000000
                    End
                    ControlTipText ="Preview Report"
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =9632
                    Top =600
                    Width =897
                    Height =255
                    TabIndex =7
                    Name ="FPPGHFOOTER"
                    ControlSource ="FPPGHFOOTER"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =10587
                            Top =607
                            Width =1410
                            Height =240
                            Name ="Label19"
                            Caption ="ارتفاع پابرگ فاكتور :"
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
                    Left =7367
                    Top =600
                    Width =897
                    Height =255
                    TabIndex =8
                    Name ="FPPGHDETAIL"
                    ControlSource ="FPPGHDETAIL"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =8322
                            Top =607
                            Width =1215
                            Height =240
                            Name ="Label22"
                            Caption ="ارتفاع هر سطر :"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =93
                    ReadingOrder =1
                    Left =407
                    Top =686
                    Width =851
                    Height =283
                    TabIndex =9
                    Name ="LANDSC"
                    ControlSource ="LANDSC"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =195
                            Top =630
                            Width =840
                            Height =240
                            Name ="Label24"
                            Caption ="Landscape"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "FPFORMAT.cls"
