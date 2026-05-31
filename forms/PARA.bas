Version =20
VersionRequired =20
Begin Form
    DefaultView =0
    AllowUpdating =4
    TabularCharSet =178
    TabularFamily =42
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =6351
    DatasheetFontHeight =10
    ItemSuffix =7
    Left =1485
    Top =990
    Right =8955
    Bottom =4725
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x27c28697d791e240
    End
    DatasheetFontName ="Arial (Arabic)"
    OnError ="[Event Procedure]"
    Begin
        Begin Label
            BackStyle =0
            TextFontCharSet =178
            FontName ="Tahoma"
        End
        Begin TextBox
            SpecialEffect =2
            OldBorderStyle =0
            TextFontCharSet =178
            Width =1701
            LabelX =-1701
            FontName ="Tahoma"
        End
        Begin Section
            Height =1903
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =965
                    Top =226
                    Name ="COD"
                End
                Begin TextBox
                    OverlapFlags =85
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =341
                    Top =907
                    TabIndex =1
                    Name ="ANBAR"
                End
                Begin TextBox
                    OverlapFlags =85
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Top =1303
                    TabIndex =2
                    Name ="DT"
                    Format ="yyyy/mm/dd"
                End
                Begin TextBox
                    OverlapFlags =85
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Top =1663
                    TabIndex =3
                    Name ="id"
                End
            End
        End
    End
End
CodeBehindForm
' See "PARA.cls"
