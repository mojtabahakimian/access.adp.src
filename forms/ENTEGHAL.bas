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
    ScrollBars =0
    ViewsAllowed =1
    TabularCharSet =178
    TabularFamily =1
    BorderStyle =3
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =5325
    DatasheetFontHeight =10
    ItemSuffix =7
    Left =1980
    Top =1110
    Right =7935
    Bottom =2850
    DatasheetGridlinesColor =12632256
    MaxRecords =100000
    RecSrcDt = Begin
        0x0b22b4aafed4e240
    End
    Caption ="انتقال اسناد از فايل ديگر"
    DatasheetFontName ="Arial"
    OnError ="[Event Procedure]"
    AllowDatasheetView =0
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
            Height =1765
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    Left =1531
                    Top =1133
                    Width =1935
                    Height =405
                    Name ="Command0"
                    Caption ="عمليات انتقال انجام شود"
                    OnClick ="[Event Procedure]"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =453
                    Top =396
                    Width =3515
                    Height =450
                    TabIndex =1
                    Name ="PATH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4075
                            Top =515
                            Width =570
                            Height =240
                            Name ="Label2"
                            Caption ="از فايل:"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "ENTEGHAL.cls"
