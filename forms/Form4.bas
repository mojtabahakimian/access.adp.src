Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    AllowUpdating =4
    TabularCharSet =178
    TabularFamily =255
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =17858
    DatasheetFontHeight =10
    ItemSuffix =4
    Left =360
    Top =675
    Right =20520
    Bottom =9420
    DatasheetGridlinesColor =12632256
    MaxRecords =1000000
    RecSrcDt = Begin
        0x0739b3ceb868e540
    End
    DatasheetFontName ="Arial"
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
        Begin Section
            Height =6406
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    Left =7200
                    Top =6009
                    Name ="Command0"
                    Caption ="Command0"
                    OnClick ="[Event Procedure]"
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    OverlapFlags =215
                    TextFontCharSet =0
                    IMESentenceMode =3
                    Left =340
                    Top =113
                    Width =17063
                    Height =5338
                    FontSize =10
                    TabIndex =1
                    Name ="txt"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Top =113
                            Width =555
                            Height =240
                            Name ="Label2"
                            Caption ="Text1:"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =2551
                    Top =5782
                    TabIndex =2
                    Name ="Command3"
                    Caption ="Command3"
                    OnClick ="[Event Procedure]"
                End
            End
        End
    End
End
CodeBehindForm
' See "Form4.cls"
