Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    KeyPreview = NotDefault
    AllowUpdating =4
    TabularCharSet =178
    Orientation =1
    PictureAlignment =5
    PictureSizeMode =1
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =14685
    RowHeight =435
    ItemSuffix =33
    Left =90
    Top =885
    Right =14775
    Bottom =6900
    FrozenColumns =5
    HelpContextId =5009
    DatasheetForeColor =33554432
    DatasheetGridlinesColor =12632256
    MaxRecords =10000000
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
    RecSrcDt = Begin
        0xe2e4a52d8fb6e440
    End
    Caption ="ليست محدود شده حسابهاي معين"
    OnOpen ="[Event Procedure]"
    HelpFile ="dena.hlp"
    DatasheetFontName ="Tahoma"
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
            Height =350
            BackColor =-2147483633
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    TextFontFamily =2
                    Left =13892
                    Width =750
                    Height =300
                    FontSize =7
                    Name ="NUMBER Label"
                    Caption ="حساب كل"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="NUMBER_Label"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =2
                    Left =5460
                    Width =1080
                    Height =300
                    Name ="SumOfBED Label"
                    Caption ="جمع بدهكار"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="SumOfBED_Label"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =2
                    Left =3690
                    Width =1170
                    Height =300
                    Name ="SumOfBES Label"
                    Caption ="جمع بستانكار"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="SumOfBES_Label"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =2
                    Left =1950
                    Width =1140
                    Height =300
                    Name ="bed Label"
                    Caption ="مانده بدهكار"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="bed_Label"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =2
                    Left =330
                    Width =1230
                    Height =300
                    Name ="bes Label"
                    Caption ="مانده بستانكار"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="bes_Label"
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    TextFontFamily =2
                    Left =12925
                    Width =975
                    Height =300
                    Name ="Label23"
                    Caption =" حساب معين"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =2
                    Left =9070
                    Width =1185
                    Height =300
                    Name ="Label24"
                    Caption ="نام حساب معين"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
            End
        End
        Begin Section
            Height =315
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    SpecialEffect =2
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    BackStyle =1
                    NumeralShapes =2
                    Left =13911
                    Top =-3
                    Width =774
                    Height =315
                    ColumnWidth =840
                    ColumnOrder =0
                    Name ="حساب كل"
                    ControlSource ="HES_K"
                    FontName ="WeblogmaYekan"
                    EventProcPrefix ="حساب_كل"
                End
                Begin TextBox
                    DecimalPlaces =0
                    SpecialEffect =2
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    BackStyle =1
                    NumeralShapes =2
                    Left =5235
                    Width =1659
                    Height =315
                    ColumnWidth =1350
                    ColumnOrder =4
                    TabIndex =1
                    Name ="جمع بدهكار"
                    ControlSource ="SumOfBED"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                    EventProcPrefix ="جمع_بدهكار"
                End
                Begin TextBox
                    DecimalPlaces =0
                    SpecialEffect =2
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    BackStyle =1
                    NumeralShapes =2
                    Left =3480
                    Width =1659
                    Height =315
                    ColumnWidth =1635
                    ColumnOrder =5
                    TabIndex =2
                    Name ="جمع بستانكار"
                    ControlSource ="SumOfBES"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                    EventProcPrefix ="جمع_بستانكار"
                End
                Begin TextBox
                    DecimalPlaces =0
                    SpecialEffect =2
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    BackStyle =1
                    NumeralShapes =2
                    Left =1740
                    Width =1644
                    Height =315
                    ColumnWidth =1455
                    ColumnOrder =6
                    TabIndex =3
                    Name ="مانده بدهكار"
                    ControlSource ="bed"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                    EventProcPrefix ="مانده_بدهكار"
                End
                Begin TextBox
                    DecimalPlaces =0
                    SpecialEffect =2
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    BackStyle =1
                    NumeralShapes =2
                    Left =15
                    Width =1656
                    Height =315
                    ColumnWidth =1665
                    ColumnOrder =7
                    TabIndex =4
                    Name ="مانده بستانكار"
                    ControlSource ="bes"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                    EventProcPrefix ="مانده_بستانكار"
                End
                Begin Line
                    OverlapFlags =87
                    SpecialEffect =2
                    Left =14679
                    Width =0
                    Height =283
                    Name ="Line22"
                End
                Begin TextBox
                    SpecialEffect =2
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    BackStyle =1
                    NumeralShapes =2
                    Left =12925
                    Top =-3
                    Width =909
                    Height =315
                    ColumnWidth =600
                    ColumnOrder =1
                    TabIndex =5
                    BackColor =13434828
                    Name ="معين"
                    ControlSource ="HES_M"
                    OnDblClick ="[Event Procedure]"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    SpecialEffect =2
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    BackStyle =1
                    NumeralShapes =2
                    Left =6970
                    Top =-3
                    Width =5904
                    Height =315
                    ColumnWidth =3780
                    ColumnOrder =3
                    TabIndex =6
                    Name ="نام حساب معين"
                    ControlSource ="moin"
                    FontName ="WeblogmaYekan"
                    EventProcPrefix ="نام_حساب_معين"
                End
                Begin Line
                    OverlapFlags =85
                    SpecialEffect =2
                    Left =13871
                    Width =0
                    Height =283
                    Name ="Line27"
                End
                Begin TextBox
                    Visible = NotDefault
                    SpecialEffect =2
                    OldBorderStyle =0
                    OverlapFlags =247
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    BackStyle =1
                    NumeralShapes =2
                    Left =1462
                    Width =774
                    Height =195
                    TabIndex =7
                    Name ="N_KOL"
                    ControlSource ="HES_K"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    SpecialEffect =2
                    OldBorderStyle =0
                    OverlapFlags =247
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    BackStyle =1
                    NumeralShapes =2
                    Left =340
                    Width =909
                    Height =195
                    TabIndex =8
                    BackColor =13434828
                    Name ="NUMBER"
                    ControlSource ="HES_M"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    Visible = NotDefault
                    OverlapFlags =119
                    SpecialEffect =2
                    Left =1360
                    Width =0
                    Height =163
                    Name ="Line31"
                End
            End
        End
        Begin FormFooter
            Visible = NotDefault
            Height =0
            BackColor =-2147483633
            Name ="FormFooter"
        End
    End
End
CodeBehindForm
' See "BEDEHBESTANMOIN.cls"
