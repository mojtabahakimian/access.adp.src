Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    OrderByOn = NotDefault
    DefaultView =0
    AllowUpdating =4
    TabularCharSet =178
    TabularFamily =42
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =10246
    DatasheetFontHeight =10
    ItemSuffix =141
    Top =330
    DatasheetGridlinesColor =12632256
    Tag ="@Forms___F_MENU_BIM___MMO=Forms![F_MENU_BIM]![MMO],@Forms___Baseknow___PERSONEL="
        "Forms![Baseknow]![PERSONEL] ,@Forms___Baseknow___PERVAM =Forms![Baseknow]![PERVA"
        "M]"
    RecSrcDt = Begin
        0xbd36919dd2b9e240
    End
    InputParameters ="@Forms___F_MENU_BIM___MMO=Forms![F_MENU_BIM]![MMO],@Forms___Baseknow___PERSONEL="
        "Forms![Baseknow]![PERSONEL] ,@Forms___Baseknow___PERVAM =Forms![Baseknow]![PERVA"
        "M]"
    RecordSource ="dbo.LIST_SALARY5"
    Caption ="فيش حقوق"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Arial (Arabic)"
    OnError ="[Event Procedure]"
    UnknownProp = {258 ,66 ,12 ,4 ,324 } Begin
        0x400046006f0072006d0073005f005f005f0046005f004d0045004e0055005f00 ,
        0x420049004d005f005f005f004d004d004f003d0046006f0072006d0073002100 ,
        0x5b0046005f004d0045004e0055005f00420049004d005d0021005b004d004d00 ,
        0x4f005d002c00400046006f0072006d0073005f005f005f004200610073006500 ,
        0x6b006e006f0077005f005f005f0050004500520053004f004e0045004c003d00 ,
        0x46006f0072006d00730021005b0042006100730065006b006e006f0077005d00 ,
        0x21005b0050004500520053004f004e0045004c005d0020002c00400046006f00 ,
        0x72006d0073005f005f005f0042006100730065006b006e006f0077005f005f00 ,
        0x5f00500045005200560041004d0020003d0046006f0072006d00730021005b00 ,
        0x42006100730065006b006e006f0077005d0021005b0050004500520056004100 ,
        0x4d005d00
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
            FontName ="Arial (Arabic)"
        End
        Begin Rectangle
            BackStyle =0
            Width =850
            Height =850
        End
        Begin Line
            Width =1701
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
            FontName ="Arial (Arabic)"
        End
        Begin ListBox
            TextFontCharSet =178
            TextFontFamily =42
            OldBorderStyle =0
            Width =1701
            Height =1417
            LabelX =-1701
            FontName ="Arial (Arabic)"
        End
        Begin ComboBox
            OldBorderStyle =0
            TextFontCharSet =178
            TextFontFamily =42
            BackStyle =0
            Width =1701
            LabelX =-1701
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
        Begin Section
            KeepTogether = NotDefault
            CanGrow = NotDefault
            Height =5113
            Name ="Detail"
            Begin
                Begin Subform
                    Width =10246
                    Height =5113
                    Name ="SALARYBEST"
                    SourceObject ="Report.SALARYBEST"
                    LinkChildFields ="HES_T"
                    LinkMasterFields ="HES_T"
                End
            End
        End
    End
End
CodeBehindForm
' See "FISH_HOGHUGH11.cls"
