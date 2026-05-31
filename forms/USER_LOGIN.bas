Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    Modal = NotDefault
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
    KeyPreview = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    AllowUpdating =4
    ScrollBars =0
    ViewsAllowed =1
    TabularCharSet =178
    TabularFamily =0
    BorderStyle =3
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =6462
    DatasheetFontHeight =10
    ItemSuffix =13
    Left =4395
    Top =1830
    Right =10860
    Bottom =6750
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x73360a0e5dace240
    End
    Caption ="ورود كاربر"
    DatasheetFontName =""
    OnKeyPress ="[Event Procedure]"
    OnLoad ="[Event Procedure]"
    AllowDatasheetView =0
    Begin
        Begin Label
            BackStyle =0
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
        Begin Subform
            SpecialEffect =2
            Width =1701
            Height =1701
        End
        Begin Section
            CanGrow = NotDefault
            Height =4932
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin Subform
                    OverlapFlags =119
                    Left =113
                    Top =566
                    Width =6240
                    Height =3675
                    Name ="USER_LOGIN_SUB"
                    SourceObject ="Form.USER_LOGIN_SUB"
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    KeyboardLanguage =11
                    IMESentenceMode =3
                    Left =135
                    Top =120
                    Width =1813
                    Height =270
                    TabIndex =1
                    Name ="PSWORD"
                    AfterUpdate ="[Event Procedure]"
                    InputMask ="Password"
                    OnGotFocus ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2005
                            Top =120
                            Width =675
                            Height =255
                            Name ="Label4"
                            Caption ="رمز عبور:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2776
                    Top =113
                    Width =2893
                    Height =270
                    TabIndex =2
                    Name ="USNAME"
                    AfterUpdate ="[Event Procedure]"
                    OnDblClick ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5726
                            Top =113
                            Width =675
                            Height =255
                            Name ="Label6"
                            Caption ="نام كاربر:"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =255
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2775
                    Top =300
                    Width =2893
                    Height =270
                    TabIndex =3
                    Name ="cusname"
                End
                Begin TextBox
                    TabStop = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    TextAlign =1
                    ReadingOrder =1
                    ScrollBarAlign =2
                    BackStyle =0
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =135
                    Top =4365
                    Width =3088
                    Height =390
                    FontSize =12
                    FontWeight =700
                    TabIndex =4
                    Name ="Text10"
                    OnDblClick ="[Event Procedure]"
                    DefaultValue ="=[CurrentProject].[Connection].[Properties].[Item](\"initial catalog\").[Value]"
                End
                Begin Label
                    FontUnderline = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =178
                    TextAlign =2
                    Left =3345
                    Top =4425
                    Width =2955
                    Height =255
                    BackColor =12632256
                    ForeColor =255
                    Name ="version"
                    Caption ="Version :7.1.1: date : 1405/02/28  21:59"
                    FontName ="MS Sans Serif"
                    OnDblClick ="[Event Procedure]"
                    OnMouseDown ="[Event Procedure]"
                    HyperlinkAddress ="http://2.184.57.153:90/update/negin.zip"
                End
            End
        End
    End
End
CodeBehindForm
' See "USER_LOGIN.cls"
