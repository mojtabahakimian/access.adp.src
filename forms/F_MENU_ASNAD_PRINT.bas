Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    ControlBox = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    DefaultView =0
    AllowUpdating =4
    ScrollBars =0
    ViewsAllowed =1
    TabularCharSet =178
    TabularFamily =42
    BorderStyle =3
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =5215
    DatasheetFontHeight =10
    ItemSuffix =20
    Left =3510
    Top =2910
    Right =8730
    Bottom =4200
    HelpContextId =300
    DatasheetGridlinesColor =12632256
    MaxRecords =1000000
    RecSrcDt = Begin
        0xbace0f5a1a4ce240
    End
    Caption ="چاپ گروهي اسناد"
    HelpFile ="amin98.hlp>main"
    DatasheetFontName ="Arial (Arabic)"
    OnLoad ="[Event Procedure]"
    OnError ="[Event Procedure]"
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
            SpecialEffect =2
            OldBorderStyle =0
            TextFontCharSet =178
            Width =1701
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
        Begin UnboundObjectFrame
            SpecialEffect =2
            OldBorderStyle =1
            Width =4536
            Height =2835
        End
        Begin Section
            Height =1303
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    TextFontCharSet =0
                    Left =2455
                    Top =850
                    Width =870
                    Height =345
                    FontWeight =700
                    TabIndex =2
                    Name ="Command5"
                    Caption ="تاييد"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    Cancel = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    Left =1587
                    Top =850
                    Width =870
                    Height =347
                    FontWeight =700
                    TabIndex =3
                    Name ="Command6"
                    Caption ="انصراف"
                    OnClick ="[Event Procedure]"
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4055
                    Top =623
                    Width =729
                    Height =255
                    TabIndex =4
                    Name ="DTT"
                End
                Begin ComboBox
                    AutoExpand = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    ReadingOrder =1
                    AllowedText =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2715
                    Top =180
                    Width =1431
                    Height =300
                    Name ="SNDNUM1"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT DEED_HED.N_S FROM DEED_HED GROUP BY DEED_HED.N_S ORDER BY DEED_HED.N_S;"
                    InputMask ="############"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =4230
                            Top =180
                            Width =945
                            Height =285
                            Name ="Label17"
                            Caption ="از سند شماره:"
                        End
                    End
                End
                Begin ComboBox
                    AutoExpand = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    ReadingOrder =1
                    AllowedText =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =115
                    Top =170
                    Width =1431
                    Height =300
                    TabIndex =1
                    Name ="SNDNUM2"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT DEED_HED.N_S FROM DEED_HED GROUP BY DEED_HED.N_S ORDER BY DEED_HED.N_S DE"
                        "SC;"
                    InputMask ="############"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =1645
                            Top =170
                            Width =960
                            Height =285
                            Name ="Label4"
                            Caption ="تا سند شماره:"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "F_MENU_ASNAD_PRINT.cls"
