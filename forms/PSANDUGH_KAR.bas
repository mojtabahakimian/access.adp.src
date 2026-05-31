Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    DividingLines = NotDefault
    DefaultView =0
    AllowUpdating =4
    ScrollBars =0
    TabularCharSet =178
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =11281
    RowHeight =390
    ItemSuffix =16
    Left =240
    Top =90
    Right =11265
    Bottom =3810
    DatasheetForeColor =33554432
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x6a5129c249efe240
    End
    RecordSource ="PSANDUGH_KARKONAN"
    Caption ="ساير كسورات و صندوق كاركنان"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Tahoma"
    OnLoad ="[Event Procedure]"
    OnError ="[Event Procedure]"
    RecordSourceQualifier ="dbo"
    Begin
        Begin Label
            BackStyle =0
            TextFontCharSet =178
            FontName ="Tahoma"
        End
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
            Width =850
            Height =850
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
            Width =4536
            Height =2835
            LabelX =-1701
        End
        Begin TextBox
            SpecialEffect =2
            OldBorderStyle =0
            TextFontCharSet =178
            Width =1701
            LabelX =-1701
            FontName ="Tahoma"
        End
        Begin ListBox
            SpecialEffect =2
            TextFontCharSet =178
            Width =1701
            Height =1417
            LabelX =-1701
            FontName ="Tahoma"
        End
        Begin ComboBox
            SpecialEffect =2
            TextFontCharSet =178
            Width =1701
            LabelX =-1701
            FontName ="Tahoma"
        End
        Begin FormHeader
            Height =660
            BackColor =8421376
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    TextFontFamily =2
                    Left =3118
                    Top =15
                    Width =5130
                    Height =645
                    FontSize =18
                    ForeColor =10092543
                    Name ="Label8"
                    Caption ="ساير كسورات دائمي و صندوق كاركنان"
                    FontName ="Titr"
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    TextFontFamily =2
                    Left =3133
                    Width =5130
                    Height =645
                    FontSize =18
                    Name ="Label10"
                    Caption ="ساير كسورات دائمي و صندوق كاركنان"
                    FontName ="Titr"
                End
            End
        End
        Begin Section
            Height =3061
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =6165
                    Top =2190
                    Width =1956
                    Height =315
                    ColumnWidth =1605
                    ColumnOrder =2
                    TabIndex =2
                    BackColor =16777164
                    Name ="MABL"
                    ControlSource ="MABL"
                    BeforeUpdate ="[Event Procedure]"
                    AfterUpdate ="[Event Procedure]"
                    OnKeyPress ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextFontCharSet =0
                            TextAlign =1
                            Left =8207
                            Top =2190
                            Width =375
                            Height =240
                            Name ="Label5"
                            Caption ="مبلغ"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =3000
                    Top =285
                    Width =4821
                    Height =315
                    ColumnWidth =4260
                    ColumnOrder =0
                    Name ="CODE"
                    ControlSource ="CODE"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT CODE, PNAME + N' ' + PFAMILY + N' ' + RTRIM(CAST(CODE AS NVARCHAR)) AS PE"
                        "R FROM PERSONEL ORDER BY PNAME + N' ' + PFAMILY + N' ' + RTRIM(CAST(CODE AS NVAR"
                        "CHAR))"
                    ColumnWidths ="0"
                    OnNotInList ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =7922
                            Top =285
                            Width =675
                            Height =300
                            Name ="Label1"
                            Caption ="پرسنل"
                        End
                    End
                End
                Begin Rectangle
                    OverlapFlags =255
                    Left =1305
                    Top =1140
                    Width =8787
                    Height =1644
                    Name ="Box15"
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =247
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2610
                    Top =1650
                    Width =5511
                    Height =360
                    ColumnWidth =4545
                    ColumnOrder =1
                    TabIndex =1
                    Name ="SHARH"
                    ControlSource ="SHARH"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT SHARH FROM PSANDUGH_KARKONAN GROUP BY SHARH ORDER BY SHARH"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextFontCharSet =0
                            TextAlign =1
                            Left =8207
                            Top =1650
                            Width =420
                            Height =240
                            Name ="Label3"
                            Caption ="شرح"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =56
                    Top =3
                    ColumnWidth =2070
                    TabIndex =3
                    Name ="CRT"
                    ControlSource ="CRT"
                    DefaultValue ="Now()"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2261
                            Width =1515
                            Height =240
                            Name ="Label35"
                            Caption ="تاريخ ايجاد"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Top =283
                    TabIndex =4
                    Name ="UID"
                    ControlSource ="UID"
                    DefaultValue ="[Forms]![BASEKNOW]![USERCOD]"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =2205
                            Top =285
                            Width =1365
                            Height =240
                            Name ="Label36"
                            Caption ="کد کاربر ايجاد کننده"
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
' See "PSANDUGH_KAR.cls"
