Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    AllowEdits = NotDefault
    DefaultView =0
    AllowUpdating =4
    ViewsAllowed =1
    TabularCharSet =178
    TabularFamily =48
    BorderStyle =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =15590
    DatasheetFontHeight =10
    ItemSuffix =12
    Left =1650
    Right =20535
    Bottom =9690
    DatasheetGridlinesColor =12632256
    MaxRecords =100000000
    RecSrcDt = Begin
        0x6727831bbfb8e540
    End
    Caption ="نبض مالي"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial"
    AllowDatasheetView =0
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
        End
        Begin Subform
            SpecialEffect =2
            Width =1701
            Height =1701
        End
        Begin Section
            CanGrow = NotDefault
            Height =9840
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin Subform
                    OverlapFlags =85
                    OldBorderStyle =0
                    SpecialEffect =0
                    Left =10425
                    Top =45
                    Width =5118
                    Height =4820
                    Name ="AMAR_PISH7"
                    SourceObject ="Form.nabz_sand7"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =10290
                            Width =5103
                            Height =30
                            Name ="AMAR_PISH7 Label"
                            Caption ="AMAR_PISH7"
                            EventProcPrefix ="AMAR_PISH7_Label"
                        End
                    End
                End
                Begin Subform
                    OverlapFlags =215
                    OldBorderStyle =0
                    SpecialEffect =0
                    Left =10422
                    Top =4935
                    Width =5118
                    Height =4820
                    TabIndex =1
                    Name ="nabzasnad7"
                    SourceObject ="Form.nabzasnad7"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =10287
                            Top =4950
                            Width =5118
                            Height =345
                            Name ="Label3"
                            Caption ="AMAR_PISH7"
                        End
                    End
                End
                Begin Subform
                    OverlapFlags =85
                    OldBorderStyle =0
                    SpecialEffect =0
                    Left =5280
                    Top =45
                    Width =5118
                    Height =4820
                    TabIndex =2
                    Name ="AMAR_PISH30"
                    SourceObject ="Form.nabz_sanad30"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5145
                            Width =5103
                            Height =30
                            Name ="Label5"
                            Caption ="AMAR_PISH7"
                        End
                    End
                End
                Begin Subform
                    OverlapFlags =215
                    OldBorderStyle =0
                    SpecialEffect =0
                    Left =5280
                    Top =4935
                    Width =5118
                    Height =4820
                    TabIndex =3
                    Name ="Child6"
                    SourceObject ="Form.nabzasnad30"
                End
                Begin Subform
                    OverlapFlags =85
                    OldBorderStyle =0
                    SpecialEffect =0
                    Left =120
                    Top =60
                    Width =5118
                    Height =4820
                    TabIndex =4
                    Name ="nabz_mali90"
                    SourceObject ="Form.nabz_mali90"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Width =75
                            Height =45
                            Name ="Label9"
                            Caption ="AMAR_PISH7"
                        End
                    End
                End
                Begin Subform
                    OverlapFlags =215
                    OldBorderStyle =0
                    SpecialEffect =0
                    Left =120
                    Top =4935
                    Width =5118
                    Height =4820
                    TabIndex =5
                    Name ="nabzasnad90"
                    SourceObject ="Form.nabzasnad90"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =120
                            Top =4875
                            Width =90
                            Height =360
                            Name ="Label11"
                            Caption ="AMAR_PISH7"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "nabzmali.cls"
