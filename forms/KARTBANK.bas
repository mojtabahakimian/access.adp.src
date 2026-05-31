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
    Width =7653
    DatasheetFontHeight =10
    ItemSuffix =4
    Left =5730
    Top =3345
    Right =13380
    Bottom =4290
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x358ea6b70e6fe240
    End
    Caption ="كارت بانك"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial (Arabic)"
    OnKeyPress ="[Event Procedure]"
    OnError ="[Event Procedure]"
    AllowDatasheetView =0
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
        Begin ComboBox
            SpecialEffect =2
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
                    OverlapFlags =215
                    TextAlign =1
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1740
                    Top =60
                    Width =3891
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
                            OverlapFlags =93
                            TextAlign =3
                            TextFontFamily =10
                            Left =5610
                            Top =60
                            Width =2010
                            Height =705
                            FontSize =20
                            Name ="Label1"
                            Caption ="كارت بانك :"
                            FontName ="Titr"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    TabStop = NotDefault
                    RowSourceTypeInt =1
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =178
                    TextAlign =1
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    DecimalPlaces =0
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =56
                    Top =61
                    Width =1596
                    Height =765
                    FontSize =20
                    FontWeight =700
                    TabIndex =1
                    ForeColor =255
                    Name ="PORTP"
                    RowSourceType ="Value List"
                    RowSource ="COM1;COM2;COM3;COM4;COM5;COM6;COM7;COM8;COM9;COM10;COM11;COM12"
                    DefaultValue ="\"COM1\""
                    FontName ="Titr"
                End
            End
        End
    End
End
CodeBehindForm
' See "KARTBANK.cls"
