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
    Width =5940
    DatasheetFontHeight =10
    ItemSuffix =2
    Left =6585
    Top =3345
    Right =12525
    Bottom =4290
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x358ea6b70e6fe240
    End
    Caption ="نقدي"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial (Arabic)"
    OnKeyPress ="[Event Procedure]"
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
            Height =963
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =1
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =30
                    Top =60
                    Width =4821
                    Height =765
                    FontSize =20
                    FontWeight =700
                    ForeColor =255
                    Name ="Text0"
                    DefaultValue ="[Forms]![HEAD_LST_FROOSH22]![MAN]"
                    FontName ="Titr"
                    OnKeyPress ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            TextFontFamily =10
                            Left =5070
                            Top =60
                            Width =840
                            Height =705
                            FontSize =20
                            Name ="Label1"
                            Caption ="نقد  :"
                            FontName ="Titr"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "NAGHDF.cls"
