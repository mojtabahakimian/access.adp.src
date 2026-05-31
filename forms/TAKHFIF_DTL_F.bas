Version =20
VersionRequired =20
Begin Form
    AllowDesignChanges = NotDefault
    DefaultView =2
    AllowUpdating =4
    ViewsAllowed =2
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =5115
    RowHeight =360
    ItemSuffix =10
    Left =7155
    Top =-15
    Right =13110
    Bottom =6285
    DatasheetGridlinesColor =12632256
    MaxRecords =0
    RecSrcDt = Begin
        0xa4edd48b5105e540
    End
    RecordSource ="TAKHFIF_DEF_DTL"
    Caption ="محدوده تخفيفات"
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
            Height =1875
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    ColumnHidden = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1312
                    Top =113
                    Name ="TID"
                    ControlSource ="TID"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4500
                            Top =120
                            Width =525
                            Height =240
                            Name ="Label1"
                            Caption ="شماره"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1312
                    Top =453
                    ColumnWidth =1545
                    TabIndex =1
                    Name ="AZMAB"
                    ControlSource ="AZMAB"
                    Format ="#,###"
                    OnKeyPress ="[Event Procedure]"
                    ConditionalFormat = Begin
                        0x0100000066000000010000000000000003000000000000000200000001000000 ,
                        0x00000000ff000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x300000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4242
                            Top =453
                            Width =585
                            Height =240
                            Name ="Label3"
                            Caption ="از مبلغ"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1312
                    Top =793
                    ColumnWidth =1695
                    TabIndex =2
                    Name ="TAMAB"
                    ControlSource ="TAMAB"
                    Format ="#,###"
                    OnKeyPress ="[Event Procedure]"
                    ConditionalFormat = Begin
                        0x0100000066000000010000000000000003000000000000000200000001000000 ,
                        0x00000000ff660000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x300000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4242
                            Top =793
                            Width =585
                            Height =240
                            Name ="Label5"
                            Caption ="تا مبلغ"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1312
                    Top =1133
                    ColumnWidth =1035
                    TabIndex =3
                    Name ="TPERS"
                    ControlSource ="TPERS"
                    OnKeyPress ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =4290
                            Top =1140
                            Width =810
                            Height =405
                            Name ="Label7"
                            Caption ="درصد تخفيف"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1312
                    Top =1474
                    ColumnWidth =1305
                    TabIndex =4
                    Name ="TMAB"
                    ControlSource ="TMAB"
                    OnKeyPress ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =4350
                            Top =1470
                            Width =765
                            Height =405
                            Name ="Label9"
                            Caption ="مبلغ ثابت تخفيف"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "TAKHFIF_DTL_F.cls"
