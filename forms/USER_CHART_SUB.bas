Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =2
    AllowUpdating =4
    ScrollBars =0
    ViewsAllowed =2
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =4940
    RowHeight =300
    DatasheetFontWeight =700
    ItemSuffix =7
    Left =8220
    Top =1740
    Right =15375
    Bottom =7860
    DatasheetGridlinesColor =12632256
    MaxRecords =100000
    RecSrcDt = Begin
        0x79adbfefd37fe540
    End
    RecordSource ="CHARTSAZMANI"
    Caption ="OPANBACCESS"
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
            Height =943
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1814
                    Name ="USERCO"
                    ControlSource ="USERCO"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Width =690
                            Height =240
                            Name ="Label1"
                            Caption ="كاربر"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =1814
                    Top =340
                    ColumnWidth =4980
                    TabIndex =1
                    Name ="SUBUSERCO"
                    ControlSource ="SUBUSERCO"
                    RowSourceType ="Value List"
                    RowSource ="SELECT SALA_DTL.IDD, DECODEUN([SAL_NAME]) AS Expr1 FROM SALA_DTL;"
                    ColumnWidths ="0"
                    OnDblClick ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =7
                            Top =337
                            Width =615
                            Height =240
                            Name ="Label3"
                            Caption ="نام كاربر"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "USER_CHART_SUB.cls"
