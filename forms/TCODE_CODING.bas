Version =20
VersionRequired =20
Begin Form
    AutoCenter = NotDefault
    KeyPreview = NotDefault
    DefaultView =0
    AllowUpdating =4
    ViewsAllowed =2
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =7937
    RowHeight =285
    ItemSuffix =8
    Left =4530
    Top =645
    Right =19095
    Bottom =8715
    DatasheetGridlinesColor =12632256
    MaxRecords =10000000
    Filter ="cityname like N'%زد%'"
    RecSrcDt = Begin
        0x0efdec259e90e540
    End
    RecordSource ="TCOD_CITY"
    Caption ="كدينگ مقصد"
    OnCurrent ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Tahoma"
    OnKeyPress ="[Event Procedure]"
    OnError ="[Event Procedure]"
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
            Height =1747
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =6021
                    Top =453
                    Width =1866
                    Name ="CITYCODE"
                    ControlSource ="CITYCODE"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =907
                            Top =453
                            Width =705
                            Height =240
                            Name ="Label3"
                            Caption ="كد "
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
                    Left =1927
                    Top =793
                    Width =5955
                    ColumnWidth =4740
                    TabIndex =1
                    Name ="CITYNAME"
                    ControlSource ="CITYNAME"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =793
                            Width =705
                            Height =240
                            Name ="Label5"
                            Caption ="نام"
                        End
                    End
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =6192
                    Top =1147
                    ColumnWidth =2265
                    TabIndex =2
                    Name ="OSCODE"
                    ControlSource ="OSCODE"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT OSCODE, OSNAME FROM TCOD_OSTAN"
                    ColumnWidths ="0"
                    OnDblClick ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5325
                            Top =1140
                            Width =585
                            Height =240
                            Name ="Label6"
                            Caption ="استان*"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "TCODE_CODING.cls"
