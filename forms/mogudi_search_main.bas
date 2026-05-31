Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    DefaultView =0
    AllowUpdating =4
    ViewsAllowed =1
    TabularCharSet =178
    TabularFamily =138
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =18878
    DatasheetFontHeight =10
    ItemSuffix =8
    Left =240
    Top =60
    Right =18840
    Bottom =8550
    DatasheetGridlinesColor =12632256
    MaxRecords =1000000
    RecSrcDt = Begin
        0x25364665467de540
    End
    Caption ="جستجوي موجودي"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial"
    OnError ="[Event Procedure]"
    AllowDatasheetView =0
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
        Begin Subform
            SpecialEffect =2
            Width =1701
            Height =1701
        End
        Begin Tab
            TextFontCharSet =178
            Width =5103
            Height =3402
            FontName ="Tahoma"
        End
        Begin Page
            Width =1701
            Height =1701
        End
        Begin Section
            CanGrow = NotDefault
            Height =8503
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin ComboBox
                    LimitToList = NotDefault
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3283
                    Top =60
                    Width =1431
                    Height =375
                    TabIndex =2
                    Name ="CANBAR"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT CODE FROM TCOD_ANBAR WHERE     (CODE <> 0) AND (CODE IN (SELECT  ANBCO  F"
                        "ROM dbo.OPANBACCESS  WHERE     (USERCO = 6)))"
                    ValidationRule ="Is Not Null"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="8"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =4878
                            Top =62
                            Width =390
                            Height =375
                            Name ="Label4"
                            Caption ="انبار:"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =297
                    Top =74
                    Width =2841
                    Height =375
                    TabIndex =1
                    Name ="ANBAR"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT     CODE, NAMES FROM         TCOD_ANBAR WHERE     (CODE <> 0) AND (CODE I"
                        "N (SELECT  ANBCO  FROM dbo.OPANBACCESS  WHERE     (USERCO = 6)))"
                    ColumnWidths ="0"
                    ValidationRule ="Is Not Null"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="8"
                End
                Begin Subform
                    TabStop = NotDefault
                    OverlapFlags =85
                    Left =63
                    Top =626
                    Width =18645
                    Height =7763
                    TabIndex =3
                    Name ="mogudi_search"
                    SourceObject ="Form.mogudi_search"
                End
                Begin TextBox
                    AutoTab = NotDefault
                    FELineBreak = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =10980
                    Top =60
                    Width =6539
                    Height =375
                    Name ="set"
                    AfterUpdate ="[Event Procedure]"
                    OnChange ="[Event Procedure]"
                    AsianLineBreak =0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =17700
                            Top =60
                            Width =660
                            Height =375
                            Name ="Label3"
                            Caption ="جستجو:"
                        End
                    End
                End
                Begin CheckBox
                    TabStop = NotDefault
                    OverlapFlags =85
                    ReadingOrder =1
                    Left =7041
                    Top =125
                    Width =233
                    Height =326
                    TabIndex =4
                    Name ="ONS"
                    DefaultValue ="False"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =5593
                            Top =135
                            Width =1395
                            Height =240
                            ForeColor =255
                            Name ="Label7"
                            Caption ="جستجوي لحظه اي"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "mogudi_search_main.cls"
