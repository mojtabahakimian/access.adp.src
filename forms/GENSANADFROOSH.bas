Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    ControlBox = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    CloseButton = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    AllowUpdating =4
    ScrollBars =0
    ViewsAllowed =1
    TabularCharSet =178
    TabularFamily =119
    BorderStyle =3
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =5669
    DatasheetFontHeight =10
    Left =9120
    Top =3555
    Right =14790
    Bottom =5235
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xc1f3dbbe94ade240
    End
    Caption ="صدور اسناد  بصورت گروهي"
    DatasheetFontName =""
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
        Begin ComboBox
            SpecialEffect =2
            Width =1701
            LabelX =-1701
            FontName ="Tahoma"
        End
        Begin Section
            Height =1701
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    Left =2566
                    Top =1076
                    Width =870
                    Height =345
                    FontWeight =700
                    Name ="Command5"
                    Caption ="تاييد"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    Cancel = NotDefault
                    OverlapFlags =85
                    Left =1698
                    Top =1076
                    Width =870
                    Height =347
                    FontWeight =700
                    TabIndex =1
                    Name ="Command6"
                    Caption ="انصراف"
                    OnClick ="[Event Procedure]"
                End
                Begin TextBox
                    Visible = NotDefault
                    FELineBreak = NotDefault
                    OverlapFlags =85
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4166
                    Top =849
                    Width =729
                    Height =255
                    TabIndex =2
                    Name ="DTT"
                    AsianLineBreak =0
                End
                Begin ComboBox
                    AutoExpand = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2826
                    Top =406
                    Width =1431
                    Height =300
                    TabIndex =3
                    Name ="HF1"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT NUMBER, TAG, N_S FROM HEAD_LST WHERE (TAG = 2) AND (N_S IS NULL) ORDER BY"
                        " NUMBER"
                    InputMask ="############"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4341
                            Top =406
                            Width =1215
                            Height =285
                            Name ="Label17"
                            Caption ="از برگه شماره:"
                        End
                    End
                End
                Begin ComboBox
                    AutoExpand = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =75
                    Top =396
                    Width =1431
                    Height =300
                    TabIndex =4
                    Name ="HF2"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT NUMBER, TAG, N_S FROM HEAD_LST WHERE (TAG = 2) AND (N_S IS NULL) ORDER BY"
                        " NUMBER DESC"
                    InputMask ="############"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1604
                            Top =390
                            Width =1170
                            Height =285
                            Name ="Label4"
                            Caption ="تا برگه شماره:"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "GENSANADFROOSH.cls"
