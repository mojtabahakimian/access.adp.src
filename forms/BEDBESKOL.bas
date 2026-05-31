Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    KeyPreview = NotDefault
    AllowEdits = NotDefault
    AllowUpdating =4
    Orientation =1
    PictureAlignment =5
    PictureSizeMode =1
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =14343
    RowHeight =315
    ItemSuffix =23
    Left =1995
    Right =16335
    Bottom =8025
    HelpContextId =5009
    DatasheetForeColor =33554432
    DatasheetGridlinesColor =12632256
    PaintPalette = Begin
        0x000359000000000000ccff0005cdff000aceff000fcfff0014d0ff001ad1ff00 ,
        0x1fd2ff0024d3ff0029d4ff002ed5ff0034d6ff0039d7ff003ed8ff0043d9ff00 ,
        0x48daff004edbff0053dcff0058ddff005ddeff0062dfff0068e0ff006de1ff00 ,
        0x72e2ff0077e3ff007ce4ff0082e6ff0087e7ff008ce8ff0091e9ff0096eaff00 ,
        0x9cebff00a1ecff00a6edff00abeeff00b0efff00b6f0ff00bbf1ff00c0f2ff00 ,
        0xc5f3ff00caf4ff00cccccc00cfcfcf00d0f5ff00d1d1d100d4d4d400d5f6ff00 ,
        0xd6d6d600d9d9d900daf7ff00dbdbdb00dedede00dff8ff00e0e0e000e3e3e300 ,
        0xe4f9ff00e5e5e500e8e8e800eafaff00ebebeb00ededed00effbff00f0f0f000 ,
        0xf2f2f200f4fcff00f5f5f500f7f7f700f9fdff00fafafa00fcfcfc00ffffff00 ,
        0x800000000080000080800000000080008000800000808000c0c0c000c0dcc000 ,
        0xa6caf000fffbf000a0a0a40080808000ff00000000ff0000ffff00000000ff00 ,
        0xff00ff0000ffff0000000000
    End
    Tag ="@Forms___FMENU_TARAZ_4___DT1 = [Forms]![FMENU_TARAZ_4]![DT1],@Forms___FMENU_TARA"
        "Z_4___DT2=[Forms]![FMENU_TARAZ_4]![DT2],@Forms___FMENU_TARAZ_4___SNDNUM1 = [Form"
        "s]![FMENU_TARAZ_4]![SNDNUM1],@Forms___FMENU_TARAZ_4___SNDNUM2 = [Forms]![FMENU_T"
        "ARAZ_4]![SNDNUM2]"
    RecSrcDt = Begin
        0x4003d75a91b6e440
    End
    Caption ="ليست حسابهاي كل محدود شده"
    OnOpen ="[Event Procedure]"
    HelpFile ="dena.hlp"
    DatasheetFontName ="Tahoma"
    OnLoad ="[Event Procedure]"
    Begin
        Begin Label
            BackStyle =0
            TextFontCharSet =178
            TextFontFamily =42
            FontName ="Arial (Arabic)"
        End
        Begin Rectangle
            BackStyle =0
            BorderWidth =1
            Width =850
            Height =850
        End
        Begin Line
            BorderWidth =1
            Width =1701
        End
        Begin Image
            SpecialEffect =1
            BackStyle =0
            OldBorderStyle =0
            PictureAlignment =2
            Width =1701
            Height =1701
        End
        Begin CommandButton
            TextFontCharSet =178
            TextFontFamily =42
            Width =1701
            Height =283
            FontSize =8
            FontWeight =400
            FontName ="Tahoma"
        End
        Begin OptionButton
            SpecialEffect =1
            LabelX =230
            LabelY =-30
        End
        Begin CheckBox
            SpecialEffect =1
            BorderWidth =3
            LabelX =230
            LabelY =-30
        End
        Begin OptionGroup
            Width =1701
            Height =1701
            BorderColor =12632256
        End
        Begin BoundObjectFrame
            SpecialEffect =1
            OldBorderStyle =0
            BackStyle =0
            Width =4536
            Height =2835
            LabelX =-1701
        End
        Begin TextBox
            TextFontCharSet =178
            TextFontFamily =42
            BackStyle =0
            Width =1701
            LabelX =-1701
            BorderColor =12632256
            FontName ="Arial (Arabic)"
        End
        Begin ListBox
            SpecialEffect =1
            TextFontCharSet =178
            TextFontFamily =42
            Width =1701
            Height =1417
            LabelX =-1701
            FontName ="Arial (Arabic)"
        End
        Begin ComboBox
            TextFontCharSet =178
            TextFontFamily =42
            BackStyle =0
            Width =1701
            LabelX =-1701
            BorderColor =12632256
            FontName ="Arial (Arabic)"
        End
        Begin Subform
            SpecialEffect =1
            Width =1701
            Height =1701
        End
        Begin UnboundObjectFrame
            SpecialEffect =1
            BackStyle =0
            Width =4536
            Height =2835
        End
        Begin ToggleButton
            TextFontCharSet =178
            TextFontFamily =42
            Width =283
            Height =283
            FontSize =8
            FontWeight =400
            FontName ="Arial (Arabic)"
        End
        Begin Tab
            TextFontCharSet =178
            TextFontFamily =42
            BackStyle =0
            Width =5103
            Height =3402
            FontName ="Tahoma"
        End
        Begin FormHeader
            Height =300
            BackColor =10079487
            Name ="FormHeader"
            Begin
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =2
                    TextFontFamily =34
                    Left =13358
                    Width =975
                    Height =285
                    BackColor =13434828
                    ForeColor =8388608
                    Name ="NUMBER Label"
                    Caption ="شماره "
                    FontName ="Tahoma"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="NUMBER_Label"
                End
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =2
                    TextFontFamily =34
                    Left =10204
                    Width =1020
                    Height =300
                    ForeColor =8388608
                    Name ="NAME Label"
                    Caption ="نام حساب"
                    FontName ="Tahoma"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="NAME_Label"
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =2
                    TextFontFamily =34
                    Left =5520
                    Width =1780
                    Height =300
                    BackColor =13434828
                    BorderColor =13434828
                    ForeColor =8388608
                    Name ="SumOfBED Label"
                    Caption ="گردش بدهكار"
                    FontName ="Tahoma"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="SumOfBED_Label"
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =2
                    TextFontFamily =34
                    Left =3690
                    Width =1780
                    Height =300
                    BackColor =16776960
                    ForeColor =8388608
                    Name ="SumOfBES Label"
                    Caption ="گردش بستانكار"
                    FontName ="Tahoma"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="SumOfBES_Label"
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =2
                    TextFontFamily =34
                    Left =1876
                    Width =1765
                    Height =300
                    BackColor =13434828
                    ForeColor =8388608
                    Name ="bed Label"
                    Caption ="مانده بدهكار"
                    FontName ="Tahoma"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="bed_Label"
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =2
                    TextFontFamily =34
                    Left =56
                    Width =1776
                    Height =300
                    BackColor =16776960
                    ForeColor =8388608
                    Name ="bes Label"
                    Caption ="مانده بستانكار"
                    FontName ="Tahoma"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="bes_Label"
                End
            End
        End
        Begin Section
            Height =312
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    SpecialEffect =2
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    TextFontFamily =34
                    ReadingOrder =1
                    ScrollBarAlign =2
                    BackStyle =1
                    NumeralShapes =2
                    Left =13389
                    Top =-3
                    Width =939
                    Height =315
                    ColumnWidth =1065
                    BackColor =13434828
                    Name ="شماره حساب"
                    ControlSource ="HES_K"
                    OnDblClick ="[Event Procedure]"
                    FontName ="Tahoma"
                    EventProcPrefix ="شماره_حساب"
                End
                Begin TextBox
                    SpecialEffect =2
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    TextFontFamily =34
                    ReadingOrder =1
                    ScrollBarAlign =2
                    BackStyle =1
                    NumeralShapes =2
                    Left =7369
                    Top =-3
                    Width =5949
                    Height =315
                    ColumnWidth =3675
                    TabIndex =1
                    Name ="نام حساب"
                    ControlSource ="NAME"
                    FontName ="Tahoma"
                    EventProcPrefix ="نام_حساب"
                End
                Begin TextBox
                    DecimalPlaces =0
                    SpecialEffect =2
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    TextFontFamily =34
                    ReadingOrder =1
                    ScrollBarAlign =2
                    BackStyle =1
                    NumeralShapes =2
                    Left =5520
                    Top =-3
                    Width =1779
                    Height =315
                    ColumnWidth =1800
                    TabIndex =2
                    BackColor =13434828
                    Name ="جمع بدهكار"
                    ControlSource ="=Fix([SumOfBED])"
                    Format ="Standard"
                    FontName ="Tahoma"
                    EventProcPrefix ="جمع_بدهكار"
                End
                Begin TextBox
                    DecimalPlaces =0
                    SpecialEffect =2
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    TextFontFamily =34
                    ReadingOrder =1
                    ScrollBarAlign =2
                    BackStyle =1
                    NumeralShapes =2
                    Left =3690
                    Top =-3
                    Width =1779
                    Height =315
                    ColumnWidth =1965
                    TabIndex =3
                    BackColor =16776960
                    Name ="جمع بستانكار"
                    ControlSource ="=Fix([SumOfBES])"
                    Format ="Standard"
                    FontName ="Tahoma"
                    EventProcPrefix ="جمع_بستانكار"
                End
                Begin TextBox
                    DecimalPlaces =0
                    SpecialEffect =2
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    TextFontFamily =34
                    ReadingOrder =1
                    ScrollBarAlign =2
                    BackStyle =1
                    NumeralShapes =2
                    Left =1875
                    Top =-3
                    Width =1764
                    Height =315
                    ColumnWidth =1755
                    TabIndex =4
                    BackColor =13434828
                    Name ="مانده بدهكار"
                    ControlSource ="=Fix([bed])"
                    Format ="Standard"
                    FontName ="Tahoma"
                    EventProcPrefix ="مانده_بدهكار"
                End
                Begin TextBox
                    DecimalPlaces =0
                    SpecialEffect =2
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    TextFontFamily =34
                    ReadingOrder =1
                    ScrollBarAlign =2
                    BackStyle =1
                    NumeralShapes =2
                    Left =56
                    Top =-3
                    Width =1776
                    Height =315
                    ColumnWidth =1710
                    TabIndex =5
                    BackColor =16776960
                    Name ="مانده بستانكار"
                    ControlSource ="=Fix([bes])"
                    Format ="Standard"
                    FontName ="Tahoma"
                    EventProcPrefix ="مانده_بستانكار"
                End
                Begin Line
                    OverlapFlags =85
                    SpecialEffect =2
                    Left =14340
                    Width =0
                    Height =283
                    Name ="Line22"
                End
            End
        End
        Begin FormFooter
            Height =0
            BackColor =-2147483633
            Name ="FormFooter"
        End
    End
End
CodeBehindForm
' See "BEDBESKOL.cls"
