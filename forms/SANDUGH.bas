Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    KeyPreview = NotDefault
    DefaultView =0
    AllowUpdating =4
    ScrollBars =0
    ViewsAllowed =1
    TabularCharSet =178
    BorderStyle =3
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =4795
    RowHeight =360
    ItemSuffix =5
    Left =5235
    Top =510
    Right =10035
    Bottom =6105
    DatasheetForeColor =33554432
    DatasheetGridlinesColor =12632256
    MaxRecords =1000000
    RecSrcDt = Begin
        0x0553927b3058e440
    End
    Caption ="انتخاب موقعيت چك"
    DatasheetFontName ="Tahoma"
    AllowDatasheetView =0
    Begin
        Begin Label
            BackStyle =0
            TextFontCharSet =178
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
        Begin ListBox
            SpecialEffect =2
            TextFontCharSet =178
            Width =1701
            Height =1417
            LabelX =-1701
            FontName ="Tahoma"
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
        Begin Section
            CanGrow = NotDefault
            Height =5612
            Name ="Detail"
            Begin
                Begin Subform
                    OverlapFlags =215
                    Left =40
                    Top =56
                    Width =4710
                    Height =5550
                    Name ="SANDUGH_SUB"
                    SourceObject ="Form.SANDUGH_SUB"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =3025
                            Width =1170
                            Height =240
                            Name ="SANDUGH_SUB Label"
                            Caption ="SANDUGH_SUB"
                            EventProcPrefix ="SANDUGH_SUB_Label"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "SANDUGH.cls"
