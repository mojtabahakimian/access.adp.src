Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    ControlBox = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    CloseButton = NotDefault
    DividingLines = NotDefault
    KeyPreview = NotDefault
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
    Width =7483
    DatasheetFontHeight =10
    ItemSuffix =4
    Left =5775
    Top =2220
    Right =13260
    Bottom =3675
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x3969b88c8f62e240
    End
    Caption ="'"
    DatasheetFontName ="Arial (Arabic)"
    OnKeyPress ="[Event Procedure]"
    OnActivate ="[Event Procedure]"
    OnLoad ="[Event Procedure]"
    OnError ="[Event Procedure]"
    AllowDatasheetView =0
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
        Begin Section
            Height =1474
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =55
                    Top =283
                    Width =6306
                    Height =330
                    Name ="COMP_COD"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT hes, NAME FROM CUST_HESAB"
                    ColumnWidths ="0"
                    OnNotInList ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =6442
                            Top =283
                            Width =990
                            Height =330
                            Name ="Label24"
                            Caption ="تماس گيرنده"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =3796
                    Top =1023
                    Width =1531
                    Height =340
                    TabIndex =1
                    Name ="Command2"
                    Caption ="تاييد "
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    Cancel = NotDefault
                    OverlapFlags =85
                    Left =2235
                    Top =1020
                    Width =1531
                    Height =340
                    TabIndex =2
                    Name ="Command3"
                    Caption ="لغو "
                    OnClick ="[Event Procedure]"
                End
            End
        End
    End
End
CodeBehindForm
' See "parvad.cls"
