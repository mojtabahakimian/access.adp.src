Version =20
VersionRequired =20
Begin Form
    AutoResize = NotDefault
    PopUp = NotDefault
    RecordSelectors = NotDefault
    DividingLines = NotDefault
    DefaultView =2
    AllowUpdating =4
    ScrollBars =0
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =4044
    RowHeight =300
    ItemSuffix =12
    Left =435
    Top =1185
    Right =5490
    Bottom =4920
    DatasheetBackColor =-2147483633
    DatasheetGridlinesColor =12632256
    MaxRecords =1000000
    RecSrcDt = Begin
        0xa05750088cebe540
    End
    RecordSource ="SELECT Notes.* FROM Notes WHERE (Ndone = 0)"
    Caption ="Notes"
    BeforeUpdate ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Tahoma"
    OnLoad ="[Event Procedure]"
    OnError ="[Event Procedure]"
    RecordSourceQualifier ="dbo"
    Begin
        Begin Label
            BackStyle =0
            TextFontCharSet =178
            BackColor =-2147483633
            ForeColor =-2147483630
        End
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
            Width =850
            Height =850
        End
        Begin Line
            Width =1701
        End
        Begin Image
            BackStyle =0
            OldBorderStyle =0
            PictureAlignment =2
            Width =1701
            Height =1701
        End
        Begin CommandButton
            TextFontCharSet =178
            Width =1701
            Height =283
            FontSize =8
            FontWeight =400
            FontName ="MS Sans Serif"
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
            OldBorderStyle =0
            BackStyle =0
            Width =4536
            Height =2835
            LabelX =-1701
        End
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            TextFontCharSet =178
            Width =1701
            LabelX =-1701
            BackColor =-2147483643
            ForeColor =-2147483640
            AsianLineBreak =255
        End
        Begin ListBox
            SpecialEffect =2
            TextFontCharSet =178
            Width =1701
            Height =1417
            LabelX =-1701
            BackColor =-2147483643
            ForeColor =-2147483640
        End
        Begin ComboBox
            SpecialEffect =2
            TextFontCharSet =178
            Width =1701
            LabelX =-1701
            BackColor =-2147483643
            ForeColor =-2147483640
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
        Begin ToggleButton
            TextFontCharSet =178
            Width =283
            Height =283
            FontSize =8
            FontWeight =400
            FontName ="MS Sans Serif"
        End
        Begin Tab
            TextFontCharSet =178
            BackStyle =0
            Width =5103
            Height =3402
        End
        Begin FormHeader
            Height =0
            BackColor =-2147483633
            Name ="FormHeader"
        End
        Begin Section
            Height =2193
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    ColumnHidden = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1677
                    Top =114
                    Width =615
                    Height =255
                    ColumnWidth =615
                    Name ="idd"
                    ControlSource ="idd"
                    FontName ="Tahoma"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =57
                            Top =114
                            Width =1560
                            Height =255
                            Name ="idd_Label"
                            Caption ="شماره"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =1
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1677
                    Top =456
                    Width =2310
                    Height =255
                    ColumnWidth =4260
                    TabIndex =1
                    Name ="Note"
                    ControlSource ="Note"
                    OnDblClick ="[Event Procedure]"
                    FontName ="Tahoma"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =57
                            Top =456
                            Width =1560
                            Height =255
                            Name ="Note_Label"
                            Caption ="يادداشت"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    ColumnHidden = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1677
                    Top =798
                    Width =2310
                    Height =255
                    ColumnWidth =2310
                    TabIndex =2
                    Name ="Ndate"
                    ControlSource ="Ndate"
                    DefaultValue ="=farsidate(Date())"
                    FontName ="Tahoma"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =57
                            Top =798
                            Width =1560
                            Height =255
                            Name ="Ndate_Label"
                            Caption ="تاريخ"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    ColumnHidden = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1677
                    Top =1140
                    Width =2310
                    Height =255
                    ColumnWidth =2310
                    TabIndex =3
                    Name ="Ntime"
                    ControlSource ="Ntime"
                    DefaultValue ="=Now()"
                    FontName ="Tahoma"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =57
                            Top =1140
                            Width =1560
                            Height =255
                            Name ="Ntime_Label"
                            Caption ="ساعت"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    ColumnHidden = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1677
                    Top =1482
                    Width =615
                    Height =255
                    ColumnWidth =615
                    TabIndex =4
                    Name ="userid"
                    ControlSource ="userid"
                    DefaultValue ="=[Forms]![baseknow]![USERCOD]"
                    FontName ="Tahoma"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =57
                            Top =1482
                            Width =1560
                            Height =255
                            Name ="userid_Label"
                            Caption ="کاربر"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =1677
                    Top =1824
                    ColumnWidth =525
                    TabIndex =5
                    Name ="Ndone"
                    ControlSource ="Ndone"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =57
                            Top =1824
                            Width =1560
                            Height =255
                            Name ="Ndone_Label"
                            Caption ="اتمام"
                            FontName ="Tahoma"
                        End
                    End
                End
            End
        End
        Begin FormFooter
            Height =0
            BackColor =-2147483633
            Name ="FormFooter"
        End
    End
End
CodeBehindForm
' See "Notes_frm.cls"
