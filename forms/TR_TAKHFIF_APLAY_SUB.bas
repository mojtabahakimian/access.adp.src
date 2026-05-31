Version =20
VersionRequired =20
Begin Form
    AllowDeletions = NotDefault
    AllowAdditions = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =2
    AllowUpdating =4
    ViewsAllowed =2
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =5100
    RowHeight =330
    ItemSuffix =10
    Left =630
    Top =4980
    Right =5460
    Bottom =6420
    DatasheetGridlinesColor =12632256
    MaxRecords =0
    RecSrcDt = Begin
        0xd68ae76a9105e540
    End
    RecordSource ="TR_TAKHFIF_APLAY"
    Caption ="TAKHFIF_APLAY"
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
        Begin CommandButton
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
            FELineBreak = NotDefault
            SpecialEffect =2
            OldBorderStyle =0
            Width =1701
            LabelX =-1701
            FontName ="Tahoma"
            AsianLineBreak =255
        End
        Begin ListBox
            SpecialEffect =2
            Width =1701
            Height =1417
            LabelX =-1701
            FontName ="Tahoma"
        End
        Begin ComboBox
            SpecialEffect =2
            Width =1701
            LabelX =-1701
            FontName ="Tahoma"
        End
        Begin Section
            Height =3017
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1312
                    Top =453
                    TabIndex =1
                    Name ="NUMBER"
                    ControlSource ="NUMBER"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =4137
                            Top =453
                            Width =690
                            Height =240
                            Name ="Label3"
                            Caption ="شماره"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1312
                    Top =793
                    ColumnWidth =405
                    TabIndex =2
                    Name ="KIND"
                    ControlSource ="KIND"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4392
                            Top =793
                            Width =435
                            Height =240
                            Name ="Label5"
                            Caption ="نوع"
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
                    Left =1303
                    Top =1530
                    ColumnWidth =3705
                    TabIndex =3
                    Name ="Combo6"
                    ControlSource ="TID"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TID, TSHARH FROM TAKHFIF_DEF"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4491
                            Top =1537
                            Width =600
                            Height =405
                            Name ="Label7"
                            Caption ="تخفيف"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1312
                    Top =113
                    ColumnWidth =840
                    Name ="TID"
                    ControlSource ="TID"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =4500
                            Top =120
                            Width =600
                            Height =405
                            Name ="Label1"
                            Caption ="كد تخفيف"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =906
                    Top =2437
                    TabIndex =4
                    Name ="UP_DATE"
                    ControlSource ="UP_DATE"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3498
                            Top =2437
                            Width =810
                            Height =240
                            Name ="Label8"
                            Caption ="UP_DATE:"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =906
                    Top =2777
                    TabIndex =5
                    Name ="UP_TIME"
                    ControlSource ="UP_TIME"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3528
                            Top =2777
                            Width =780
                            Height =240
                            Name ="Label9"
                            Caption ="UP_TIME:"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "TR_TAKHFIF_APLAY_SUB.cls"
