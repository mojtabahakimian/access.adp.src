Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    DefaultView =0
    AllowUpdating =4
    ScrollBars =0
    ViewsAllowed =1
    TabularCharSet =178
    TabularFamily =0
    BorderStyle =3
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =7890
    DatasheetFontHeight =10
    ItemSuffix =18
    Left =3690
    Top =2850
    Right =11580
    Bottom =4935
    DatasheetGridlinesColor =12632256
    MaxRecords =1000000
    RecSrcDt = Begin
        0xa2cec8e098d1e340
    End
    Caption ="صدور سند وصولي"
    DatasheetFontName ="Arial"
    AllowDatasheetView =0
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
            Height =2097
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    Left =4155
                    Top =1425
                    Width =1810
                    Height =450
                    Name ="Command0"
                    Caption ="ايجاد فايل"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    Cancel = NotDefault
                    OverlapFlags =85
                    Left =2325
                    Top =1431
                    Width =1810
                    Height =450
                    TabIndex =1
                    Name ="Command1"
                    Caption ="لغو"
                    OnClick ="[Event Procedure]"
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    ReadingOrder =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =675
                    Top =698
                    Width =5767
                    Height =360
                    TabIndex =2
                    Name ="DBPATH"
                    DefaultValue ="\"C:\\CFTB.DAT\""
                    Begin
                        Begin Label
                            OverlapFlags =85
                            ReadingOrder =2
                            NumeralShapes =2
                            Left =6510
                            Top =690
                            Width =1305
                            Height =345
                            Name ="Label7"
                            Caption ="مسير ذخيره فايل:"
                        End
                    End
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =690
                    Top =178
                    Width =5332
                    Height =330
                    TabIndex =3
                    Name ="BANK"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT CODE, NAMES FROM TCOD_BANKS"
                    ColumnWidths ="0"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="8"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            ReadingOrder =2
                            NumeralShapes =2
                            Left =6081
                            Top =165
                            Width =1770
                            Height =345
                            Name ="Label11"
                            Caption ="فايل جهت ارسال به بانك:"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "DISKETBANKMASKAN.cls"
