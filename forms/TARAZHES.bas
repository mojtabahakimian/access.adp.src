Version =20
VersionRequired =20
Begin Form
    OrderByOn = NotDefault
    DefaultView =2
    AllowUpdating =4
    ViewsAllowed =2
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =4940
    RowHeight =345
    DatasheetFontHeight =10
    ItemSuffix =4
    DatasheetGridlinesColor =12632256
    MaxRecords =100000
    OrderBy ="TARAZHES.NUMBER"
    RecSrcDt = Begin
        0x4bc17c8610e8e240
    End
    RecordSource ="TARAZHES"
    Caption ="حسابهاي موجود در ترازنامه"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Tahoma"
    OnError ="[Event Procedure]"
    AllowFormView =0
    RecordSourceQualifier ="dbo"
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
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
            Width =1701
            LabelX =-1701
            FontName ="Tahoma"
        End
        Begin Section
            Height =735
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
                    Left =1927
                    Top =113
                    ColumnWidth =1035
                    Name ="NUMBER"
                    ControlSource ="NUMBER"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =120
                            Width =780
                            Height =240
                            Name ="Label1"
                            Caption ="حساب كل"
                        End
                    End
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =1927
                    Top =480
                    Width =2886
                    Height =255
                    ColumnWidth =4365
                    TabIndex =1
                    Name ="Text2"
                    ControlSource ="NUMBER"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT NUMBER, NAME FROM TOTA_HES"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =487
                            Width =780
                            Height =240
                            Name ="Label3"
                            Caption ="حساب كل"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "TARAZHES.cls"
