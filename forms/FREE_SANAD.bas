Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    AllowEdits = NotDefault
    DefaultView =0
    AllowUpdating =4
    ViewsAllowed =2
    TabularCharSet =178
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =5167
    RowHeight =375
    ItemSuffix =11
    Left =5070
    Top =1860
    Right =10260
    Bottom =4515
    DatasheetForeColor =33554432
    DatasheetGridlinesColor =12632256
    MaxRecords =100000
    RecSrcDt = Begin
        0xa789515a59efe240
    End
    RecordSource ="FREE_SANAD"
    Caption ="اسناد حسابداري خالي"
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
        Begin Section
            Height =2677
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
                    Left =2154
                    Top =453
                    ColumnWidth =825
                    ColumnOrder =0
                    Name ="N_S"
                    ControlSource ="N_S"
                    OnDblClick ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =345
                            Top =450
                            Width =975
                            Height =240
                            Name ="Label3"
                            Caption ="شماره سند*"
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
                    Left =680
                    Top =1757
                    ColumnWidth =6045
                    ColumnOrder =2
                    TabIndex =1
                    Name ="SHARH_S"
                    ControlSource ="SHARH_S"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3197
                            Top =1757
                            Width =885
                            Height =240
                            Name ="Label7"
                            Caption ="شرح  سند"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =680
                    Top =2097
                    ColumnWidth =690
                    ColumnOrder =3
                    TabIndex =2
                    Name ="NO_S"
                    ControlSource ="NO_S"
                    OnDblClick ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3362
                            Top =2097
                            Width =720
                            Height =240
                            Name ="Label8"
                            Caption ="نوع سند"
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
                    Left =680
                    Top =2437
                    ColumnWidth =1665
                    ColumnOrder =4
                    TabIndex =3
                    Name ="USER_NAME"
                    ControlSource ="USER_NAME"
                    OnDblClick ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3047
                            Top =2437
                            Width =1035
                            Height =240
                            Name ="Label9"
                            Caption ="*كاربر"
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
                    Left =690
                    Top =1410
                    ColumnWidth =1080
                    ColumnOrder =1
                    TabIndex =4
                    Name ="DATE_S"
                    ControlSource ="DATE_S"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3219
                            Top =1410
                            Width =810
                            Height =240
                            Name ="Label10"
                            Caption ="تاريخ سند"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "FREE_SANAD.cls"
