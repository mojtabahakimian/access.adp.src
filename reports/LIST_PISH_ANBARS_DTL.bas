Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    DefaultView =0
    AllowUpdating =4
    TabularCharSet =178
    TabularFamily =42
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =10836
    DatasheetFontHeight =10
    ItemSuffix =191
    Top =-11986
    DatasheetGridlinesColor =12632256
    Tag ="@DT1 BIGINT = [Forms]![F_MENU_GOZARESH_FROOSH]![DT1],@DT2 BIGINT = [Forms]![F_ME"
        "NU_GOZARESH_FROOSH]![DT2],@USER = [Forms]![F_MENU_GOZARESH_FROOSH]![USERR]"
    RecSrcDt = Begin
        0x710c646fb424e440
    End
    InputParameters ="@DT1 BIGINT = [Forms]![F_MENU_GOZARESH_FROOSH]![DT1],@DT2 BIGINT = [Forms]![F_ME"
        "NU_GOZARESH_FROOSH]![DT2],@USER = [Forms]![F_MENU_GOZARESH_FROOSH]![USERR]"
    RecordSource ="Q_LIST_PISHSH_DAYLY_DTL"
    Caption ="گزارش فروش كالاها به تفكيك انبارها"
    DatasheetFontName ="Arial (Arabic)"
    UnknownProp = {258 ,66 ,12 ,4 ,308 } Begin
        0x4000440054003100200042004900470049004e00540020003d0020005b004600 ,
        0x6f0072006d0073005d0021005b0046005f004d0045004e0055005f0047004f00 ,
        0x5a00410052004500530048005f00460052004f004f00530048005d0021005b00 ,
        0x4400540031005d002c004000440054003200200042004900470049004e005400 ,
        0x20003d0020005b0046006f0072006d0073005d0021005b0046005f004d004500 ,
        0x4e0055005f0047004f005a00410052004500530048005f00460052004f004f00 ,
        0x530048005d0021005b004400540032005d002c00400055005300450052002000 ,
        0x3d0020005b0046006f0072006d0073005d0021005b0046005f004d0045004e00 ,
        0x55005f0047004f005a00410052004500530048005f00460052004f004f005300 ,
        0x48005d0021005b00550053004500520052005d00
    End
    RecordSourceQualifier ="dbo"
    Begin
        Begin Label
            BackStyle =0
            TextFontCharSet =178
            TextAlign =1
            TextFontFamily =42
            FontSize =9
            FontWeight =700
            ForeColor =128
            FontName ="Arial (Arabic)"
        End
        Begin Rectangle
            BackStyle =0
            BorderWidth =1
            Width =850
            Height =850
        End
        Begin Line
            Width =1701
            BorderColor =128
        End
        Begin Image
            OldBorderStyle =0
            PictureAlignment =2
            Width =1701
            Height =1701
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
            LabelX =230
            LabelY =-30
        End
        Begin CheckBox
            LabelX =230
            LabelY =-30
        End
        Begin OptionGroup
            BackStyle =1
            Width =1701
            Height =1701
        End
        Begin BoundObjectFrame
            Width =4536
            Height =2835
            LabelX =-1701
        End
        Begin TextBox
            OldBorderStyle =0
            TextFontCharSet =178
            TextFontFamily =42
            BackStyle =0
            Width =1701
            LabelX =-1701
            FontSize =9
            FontName ="Arial (Arabic)"
        End
        Begin ListBox
            TextFontCharSet =178
            TextFontFamily =42
            OldBorderStyle =0
            Width =1701
            Height =1417
            LabelX =-1701
            FontSize =9
            FontName ="Arial (Arabic)"
        End
        Begin ComboBox
            OldBorderStyle =0
            TextFontCharSet =178
            TextFontFamily =42
            BackStyle =0
            Width =1701
            LabelX =-1701
            FontSize =9
            FontName ="Arial (Arabic)"
        End
        Begin Subform
            OldBorderStyle =0
            Width =1701
            Height =1701
        End
        Begin UnboundObjectFrame
            Width =4536
            Height =2835
        End
        Begin BreakLevel
            GroupFooter = NotDefault
            ControlSource ="SORTI"
        End
        Begin BreakLevel
            ControlSource ="NUMBER"
        End
        Begin BreakLevel
            ControlSource ="DATE_N"
        End
        Begin BreakLevel
            ControlSource ="CUST_NO"
        End
        Begin BreakLevel
            SortOrder = NotDefault
            ControlSource ="GHABEL"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =1038
            Name ="ReportHeader"
            Begin
                Begin Line
                    BorderWidth =3
                    Top =510
                    Width =10836
                    BorderColor =0
                    Name ="Line99"
                End
                Begin Line
                    BorderWidth =3
                    Left =10815
                    Top =525
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line130"
                End
                Begin Line
                    Left =10230
                    Top =525
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line131"
                End
                Begin Line
                    Left =10215
                    Top =525
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line132"
                End
                Begin Line
                    Left =5580
                    Top =528
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line133"
                End
                Begin Line
                    Left =2235
                    Top =525
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line135"
                End
                Begin Line
                    Left =1290
                    Top =525
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line137"
                End
                Begin Line
                    BorderWidth =3
                    Left =30
                    Top =525
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line138"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =75
                    Top =573
                    Width =1155
                    Height =405
                    FontSize =10
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label139"
                    Caption =" قابل پرداخت"
                    FontName ="Traffic"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =1380
                    Top =570
                    Width =810
                    Height =405
                    FontSize =10
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label141"
                    Caption ="مقدار كل"
                    FontName ="Traffic"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =3368
                    Top =570
                    Width =1710
                    Height =405
                    FontSize =10
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label144"
                    Caption ="شــــــــــرح  كالا"
                    FontName ="Traffic"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =10236
                    Top =566
                    Width =540
                    Height =405
                    FontSize =10
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label146"
                    Caption ="رديف"
                    FontName ="Traffic"
                End
                Begin Line
                    BorderWidth =3
                    Top =1035
                    Width =10836
                    BorderColor =0
                    Name ="Line147"
                End
                Begin Line
                    Left =9555
                    Top =525
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line165"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =8787
                    Top =566
                    Width =480
                    Height =405
                    FontSize =10
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label166"
                    Caption ="تاريخ"
                    FontName ="Traffic"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =9615
                    Top =573
                    Width =570
                    Height =405
                    FontSize =10
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label170"
                    Caption ="فاكتور"
                    FontName ="Traffic"
                End
                Begin Line
                    Left =8642
                    Top =528
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line185"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =6420
                    Top =573
                    Width =1725
                    Height =405
                    FontSize =10
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label186"
                    Caption ="نام مشتري"
                    FontName ="Traffic"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =4251
                    Width =2340
                    Height =465
                    FontSize =12
                    ForeColor =255
                    Name ="LIST_PISH_ANBARS_DTL Label"
                    Caption ="پيش فاكتور هاي صادر شده"
                    FontName ="Titr"
                    EventProcPrefix ="LIST_PISH_ANBARS_DTL_Label"
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =450
            Name ="Detail"
            Begin
                Begin TextBox
                    DecimalPlaces =0
                    RunningSum =1
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =10246
                    Top =16
                    Width =504
                    Height =405
                    FontSize =10
                    FontWeight =700
                    BorderColor =32768
                    Name ="RADIF"
                    ControlSource ="=1"
                    FontName ="Traffic"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2278
                    Top =16
                    Width =3264
                    Height =405
                    FontSize =10
                    FontWeight =700
                    TabIndex =1
                    BorderColor =32768
                    Name ="KALA"
                    ControlSource ="KALA"
                    FontName ="Traffic"
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =2
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1305
                    Top =16
                    Width =894
                    Height =405
                    FontSize =10
                    FontWeight =700
                    TabIndex =2
                    BorderColor =32768
                    Name ="MEGHk"
                    ControlSource ="MEGHk"
                    Format ="Fixed"
                    FontName ="Traffic"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =60
                    Top =16
                    Width =1194
                    Height =405
                    FontSize =10
                    FontWeight =700
                    TabIndex =3
                    BorderColor =32768
                    Name ="GHABEL"
                    ControlSource ="GHABEL"
                    Format ="#,###"
                    FontName ="Traffic"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Top =450
                    Width =10821
                    BorderColor =0
                    Name ="Line100"
                End
                Begin Line
                    BorderWidth =3
                    Left =10815
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line101"
                End
                Begin Line
                    BorderWidth =1
                    Left =10230
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line102"
                End
                Begin Line
                    BorderWidth =1
                    Left =10215
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line103"
                End
                Begin Line
                    BorderWidth =1
                    Left =2235
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line104"
                End
                Begin Line
                    BorderWidth =1
                    Left =5580
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line106"
                End
                Begin Line
                    BorderWidth =1
                    Left =1290
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line108"
                End
                Begin Line
                    BorderWidth =2
                    Left =15
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line109"
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =2
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =8670
                    Top =16
                    Width =849
                    Height =405
                    FontSize =10
                    FontWeight =700
                    TabIndex =4
                    BorderColor =32768
                    Name ="DATE_N"
                    ControlSource ="DATE_N"
                    FontName ="Traffic"
                    InputMask ="####/##/##"
                End
                Begin Line
                    BorderWidth =1
                    Left =9555
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line168"
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =2
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =9590
                    Top =16
                    Width =594
                    Height =405
                    FontSize =10
                    FontWeight =700
                    TabIndex =5
                    BorderColor =32768
                    Name ="NUMBER"
                    ControlSource ="NUMBER"
                    FontName ="Traffic"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5610
                    Top =16
                    Width =2994
                    Height =405
                    FontSize =10
                    FontWeight =700
                    TabIndex =6
                    BorderColor =32768
                    Name ="MOSHTARY"
                    ControlSource ="MOSHTARY"
                    FontName ="Traffic"
                End
                Begin Line
                    BorderWidth =1
                    Left =8642
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line188"
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            CanGrow = NotDefault
            ForceNewPage =2
            Height =465
            Name ="GroupFooter3"
            Begin
                Begin Line
                    BorderWidth =3
                    Left =30
                    Top =30
                    Width =10780
                    BorderColor =0
                    Name ="Line110"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =56
                    Width =1884
                    Height =405
                    FontSize =10
                    FontWeight =700
                    BorderColor =32768
                    Name ="GHABELG"
                    ControlSource ="=Sum([GHABEL])"
                    Format ="#,###"
                    FontName ="Traffic"
                End
                Begin Line
                    BorderWidth =3
                    Left =30
                    Top =465
                    Width =10795
                    BorderColor =0
                    Name ="Line176"
                End
                Begin Line
                    BorderWidth =3
                    Left =30
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line177"
                End
                Begin Line
                    BorderWidth =3
                    Left =10815
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line178"
                End
                Begin Line
                    BorderWidth =3
                    Left =2055
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line180"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =4815
                    Width =1575
                    Height =405
                    FontSize =10
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label181"
                    Caption ="جمع مبالغ "
                    FontName ="Traffic"
                End
            End
        End
        Begin FormFooter
            KeepTogether = NotDefault
            Height =0
            Name ="ReportFooter"
        End
    End
End
CodeBehindForm
' See "LIST_PISH_ANBARS_DTL.cls"
