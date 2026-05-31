Version =20
VersionRequired =20
Begin Form
    AllowDeletions = NotDefault
    AllowAdditions = NotDefault
    AllowEdits = NotDefault
    DefaultView =2
    AllowUpdating =4
    ViewsAllowed =2
    TabularCharSet =178
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =4785
    RowHeight =330
    ItemSuffix =4
    Left =690
    Top =1890
    Right =16215
    Bottom =9765
    DatasheetForeColor =33554432
    DatasheetGridlinesColor =12632256
    MaxRecords =100000
    RecSrcDt = Begin
        0x00d79ee850f4e540
    End
    RecordSource ="SELECT        NAME, ADDRESS, ISNULL(TEL, N' ') + N' | ' + ISNULL(ADDRESS, N' ') "
        "+ N' ' + ISNULL(TOZIH, N' ') AS TEL, TOZIH FROM            dbo.CUST_HESAB"
    Caption ="تلفون"
    DatasheetFontName ="Tahoma"
    AllowFormView =0
    RecordSourceQualifier ="dbo"
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
            SpecialEffect =2
            OldBorderStyle =0
            TextFontCharSet =178
            Width =1701
            LabelX =-1701
            FontName ="Tahoma"
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
        Begin Section
            Height =1020
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =933
                    Top =113
                    ColumnWidth =11985
                    ColumnOrder =1
                    Name ="ID"
                    ControlSource ="TEL"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4260
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
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =144
                    Top =453
                    Width =2490
                    ColumnWidth =2985
                    ColumnOrder =0
                    TabIndex =1
                    Name ="Field1"
                    ControlSource ="name"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4005
                            Top =450
                            Width =765
                            Height =240
                            Name ="Label3"
                            Caption ="نام"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "TELSUB.cls"
