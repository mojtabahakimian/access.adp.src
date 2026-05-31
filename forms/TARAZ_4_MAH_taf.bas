Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    DividingLines = NotDefault
    KeyPreview = NotDefault
    AllowUpdating =4
    Orientation =1
    PictureAlignment =5
    PictureSizeMode =1
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =14520
    RowHeight =315
    ItemSuffix =34
    Left =3390
    Top =765
    Right =17925
    Bottom =6435
    HelpContextId =5009
    DatasheetForeColor =33554432
    DatasheetGridlinesColor =12632256
    MaxRecords =1000000
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
    Filter ="(((NEMUDAR.HES_M=1))) AND ((NEMUDAR.HES_T=5))"
    RecSrcDt = Begin
        0xab58e86b11cfe340
    End
    InputParameters ="@KOL = [FORMS]![TARAZ_4_MAH]![NUMBER]"
    RecordSource ="NEMUDAR"
    Caption ="تراز حسابها به تفكيك ماه"
    HelpFile ="dena.hlp"
    DatasheetFontName ="Tahoma"
    UnknownProp = {258 ,66 ,12 ,4 ,74 } Begin
        0x40004b004f004c0020003d0020005b0046004f0052004d0053005d0021005b00 ,
        0x54004100520041005a005f0034005f004d00410048005d0021005b004e005500 ,
        0x4d004200450052005d00
    End
    PivotTableXML ="<xml xmlns:x=\"urn:schemas-microsoft-com:office:excel\">\015\012 <x:WorksheetOpt"
        "ions>\015\012  <x:DisplayRightToLeft/>\015\012 </x:WorksheetOptions>\015\012 <x:"
        "PivotTable>\015\012  <x:OWCVersion>10.0.0.6854         </x:OWCVersion>\015\012  "
        "<x:NoDisplayAlerts/>\015\012  <x:NoDisplayToolbar/>\015\012  <x:DisplayScreenTip"
        "s/>\015\012  <x:NoAutoFit/>\015\012  <x:Height>384</x:Height>\015\012  <x:Width>"
        "541</x:Width>\015\012  <x:CubeProvider>msolap.2</x:CubeProvider>\015\012  <x:Cac"
        "heDetails/>\015\012  <x:Name>Microsoft Office PivotTable 10.0</x:Name>\015\012  "
        "<x:PivotField>\015\012   <x:Name>شماره حساب</x:Name>\015\012   <x:SourceName>HES"
        "_K</x:SourceName>\015\012   <x:DataType>Integer</x:DataType>\015\012   <x:Encode"
        "dType>adInteger</x:EncodedType>\015\012   <x:CompareOrderedMembersBy>UniqueName<"
        "/x:CompareOrderedMembersBy>\015\012  </x:PivotField>\015\012  <x:PivotField>\015"
        "\012   <x:Name>ماه</x:Name>\015\012   <x:SourceName>MON</x:SourceName>\015\012  "
        " <x:DataType>Integer</x:DataType>\015\012   <x:Orientation>Row</x:Orientation>\015"
        "\012   <x:EncodedType>adInteger</x:EncodedType>\015\012   <x:CompareOrderedMembe"
        "rsBy>UniqueName</x:CompareOrderedMembersBy>\015\012   <x:Expanded/>\015\012  </x"
        ":PivotField>\015\012  <x:PivotField>\015\012   <x:Name>نام حساب</x:Name>\015\012"
        "   <x:SourceName>kol</x:SourceName>\015\012   <x:EncodedType>adVarWChar</x:Encod"
        "edType>\015\012   <x:CompareOrderedMembersBy>UniqueName</x:CompareOrderedMembers"
        "By>\015\012  </x:PivotField>\015\012  <x:PivotField>\015\012   <x:Name>مانده بده"
        "كار</x:Name>\015\012   <x:SourceName>BED</x:SourceName>\015\012   <x:DataType>Nu"
        "mber</x:DataType>\015\012   <x:EncodedType>adDouble</x:EncodedType>\015\012   <x"
        ":NumberFormat>Standard</x:NumberFormat>\015\012   <x:CompareOrderedMembersBy>Uni"
        "queName</x:CompareOrderedMembersBy>\015\012  </x:PivotField>\015\012  <x:PivotFi"
        "eld>\015\012   <x:Name>مانده بستانكار</x:Name>\015\012   <x:SourceName>BES</x:So"
        "urceName>\015\012   <x:DataType>Number</x:DataType>\015\012   <x:PLDataOrientati"
        "on/>\015\012   <x:EncodedType>adDouble</x:EncodedType>\015\012   <x:NumberFormat"
        ">Standard</x:NumberFormat>\015\012   <x:CompareOrderedMembersBy>UniqueName</x:Co"
        "mpareOrderedMembersBy>\015\012  </x:PivotField>\015\012  <x:PivotField>\015\012 "
        "  <x:Name>تفصيلي</x:Name>\015\012   <x:SourceName>TAFZ</x:SourceName>\015\012   "
        "<x:EncodedType>adVarWChar</x:EncodedType>\015\012   <x:CompareOrderedMembersBy>U"
        "niqueName</x:CompareOrderedMembersBy>\015\012  </x:PivotField>\015\012  <x:Pivot"
        "Field>\015\012   <x:Name>تفصيلي 1</x:Name>\015\012   <x:SourceName>HES_T</x:Sour"
        "ceName>\015\012   <x:DataType>Integer</x:DataType>\015\012   <x:EncodedType>adIn"
        "teger</x:EncodedType>\015\012   <x:CompareOrderedMembersBy>UniqueName</x:Compare"
        "OrderedMembersBy>\015\012  </x:PivotField>\015\012  <x:PivotField>\015\012   <x:"
        "Name>معين</x:Name>\015\012   <x:SourceName>HES_M</x:SourceName>\015\012   <x:Dat"
        "aType>Integer</x:DataType>\015\012   <x:EncodedType>adInteger</x:EncodedType>\015"
        "\012   <x:CompareOrderedMembersBy>UniqueName</x:CompareOrderedMembersBy>\015\012"
        "  </x:PivotField>\015\012  <x:PivotField>\015\012   <x:Name>Data</x:Name>\015\012"
        "   <x:Orientation>Column</x:Orientation>\015\012   <x:Position>-1</x:Position>\015"
        "\012   <x:DataField/>\015\012  </x:PivotField>\015\012  <x:PivotField>\015\012  "
        " <x:Name>Sum of مانده بدهكار</x:Name>\015\012   <x:PLName>Total1</x:PLName>\015\012"
        "   <x:TotalNumber>0</x:TotalNumber>\015\012   <x:Orientation>Data</x:Orientation"
        ">\015\012   <x:Position>1</x:Position>\015\012   <x:ParentField>مانده بدهكار</x:"
        "ParentField>\015\012   <x:NumberFormat>Standard</x:NumberFormat>\015\012  </x:Pi"
        "votField>\015\012  <x:PivotField>\015\012   <x:Name>Sum of مانده بستانكار</x:Nam"
        "e>\015\012   <x:PLName>Total2</x:PLName>\015\012   <x:TotalNumber>1</x:TotalNumb"
        "er>\015\012   <x:Orientation>Data</x:Orientation>\015\012   <x:Position>2</x:Pos"
        "ition>\015\012   <x:ParentField>مانده بستانكار</x:ParentField>\015\012   <x:Numb"
        "erFormat>Standard</x:NumberFormat>\015\012  </x:PivotField>\015\012  <x:PLTotal>"
        "\015\012   <x:Name>Total3</x:Name>\015\012   <x:Caption>Count of تفصيلي</x:Capti"
        "on>\015\012   <x:ParentField>تفصيلي</x:ParentField>\015\012   <x:TotalNumber>2</"
        "x:TotalNumber>\015\012   <x:Operator>Count</x:Operator>\015\012  </x:PLTotal>\015"
        "\012  <x:PLTotal>\015\012   <x:Name>Total4</x:Name>\015\012   <x:Caption>Sum of "
        "تفصيلي 1</x:Caption>\015\012   <x:ParentField>تفصيلي 1</x:ParentField>\015\012  "
        " <x:TotalNumber>3</x:TotalNumber>\015\012  </x:PLTotal>\015\012  <x:PivotData>\015"
        "\012   <x:Top>0.0</x:Top>\015\012   <x:TopOffset>0</x:TopOffset>\015\012   <x:Le"
        "ft>0</x:Left>\015\012   <x:LeftOffset>0</x:LeftOffset>\015\012   <x:SeqNum>19</x"
        ":SeqNum>\015\012  </x:PivotData>\015\012  <x:PivotView>\015\012   <x:IsNotFilter"
        "ed/>\015\012   <x:Label>\015\012    <x:Caption>Microsoft Office PivotTable 10.0<"
        "/x:Caption>\015\012    <x:NotVisible/>\015\012   </x:Label>\015\012   <x:AllowEd"
        "its>false</x:AllowEdits>\015\012   <x:AllowAdditions>false</x:AllowAdditions>\015"
        "\012   <x:AllowDeletions>false</x:AllowDeletions>\015\012  </x:PivotView>\015\012"
        " </x:PivotTable>\015\012</xml>"
    PivotChartXML ="<xml xmlns:x=\"urn:schemas-microsoft-com:office:excel\">\015\012 <x:ChartSpace>\015"
        "\012  <x:OWCVersion>10.0.0.6854         </x:OWCVersion>\015\012  <x:Width>26988<"
        "/x:Width>\015\012  <x:Height>15161</x:Height>\015\012  <x:AllowPropertyBrowser/>"
        "\015\012  <x:DataSource>\015\012   <x:Type>PivotList</x:Type>\015\012  </x:DataS"
        "ource>\015\012  <x:BoundSeries>\015\012   <x:DataSourceIndex>0</x:DataSourceInde"
        "x>\015\012  </x:BoundSeries>\015\012  <x:Category>\015\012   <x:DataSourceIndex>"
        "0</x:DataSourceIndex>\015\012  </x:Category>\015\012  <x:Value>\015\012   <x:Dat"
        "aSourceIndex>0</x:DataSourceIndex>\015\012  </x:Value>\015\012  <x:BoundCharts>\015"
        "\012   <x:DataSourceIndex>0</x:DataSourceIndex>\015\012  </x:BoundCharts>\015\012"
        "  <x:FormatValue>\015\012   <x:DataSourceIndex>-3</x:DataSourceIndex>\015\012   "
        "<x:Data>2</x:Data>\015\012  </x:FormatValue>\015\012  <x:PivotAggOrientation>Ser"
        "ies</x:PivotAggOrientation>\015\012  <x:DisplayFieldList/>\015\012  <x:Palette>\015"
        "\012   <x:Entry>#000000</x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015\012  "
        " <x:Entry>#000000</x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015\012   <x:En"
        "try>#000000</x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015\012   <x:Entry>#0"
        "00000</x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015\012   <x:Entry>#000000<"
        "/x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015\012   <x:Entry>#000000</x:Ent"
        "ry>\015\012   <x:Entry>#000000</x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015"
        "\012   <x:Entry>#000000</x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015\012  "
        " <x:Entry>#000000</x:Entry>\015\012   <x:Entry>#8080FF</x:Entry>\015\012   <x:En"
        "try>#802060</x:Entry>\015\012   <x:Entry>#FFFFA0</x:Entry>\015\012   <x:Entry>#A"
        "0E0E0</x:Entry>\015\012   <x:Entry>#600080</x:Entry>\015\012   <x:Entry>#FF8080<"
        "/x:Entry>\015\012   <x:Entry>#008080</x:Entry>\015\012   <x:Entry>#C0C0FF</x:Ent"
        "ry>\015\012   <x:Entry>#000080</x:Entry>\015\012   <x:Entry>#FF00FF</x:Entry>\015"
        "\012   <x:Entry>#80FFFF</x:Entry>\015\012   <x:Entry>#0080FF</x:Entry>\015\012  "
        " <x:Entry>#FF8080</x:Entry>\015\012   <x:Entry>#C0FF80</x:Entry>\015\012   <x:En"
        "try>#FFC0FF</x:Entry>\015\012   <x:Entry>#FF80FF</x:Entry>\015\012  </x:Palette>"
        "\015\012  <x:DefaultFont>Arial</x:DefaultFont>\015\012  <x:Chart>\015\012   <x:P"
        "lotArea>\015\012    <x:Graph>\015\012     <x:SubType>Clustered</x:SubType>\015\012"
        "     <x:Type>Column</x:Type>\015\012     <x:Series>\015\012      <x:FormatMap>\015"
        "\012      </x:FormatMap>\015\012      <x:Name>Sum of مانده بدهكار</x:Name>\015\012"
        "      <x:Caption>\015\012       <x:DataSourceIndex>-1</x:DataSourceIndex>\015\012"
        "       <x:Data>&quot;Sum of مانده بدهكار&quot;</x:Data>\015\012      </x:Caption"
        ">\015\012      <x:Index>0</x:Index>\015\012      <x:Category>\015\012       <x:D"
        "ataSourceIndex>0</x:DataSourceIndex>\015\012      </x:Category>\015\012      <x:"
        "Value>\015\012       <x:DataSourceIndex>0</x:DataSourceIndex>\015\012      </x:V"
        "alue>\015\012      <x:FormatValue>\015\012       <x:DataSourceIndex>-3</x:DataSo"
        "urceIndex>\015\012       <x:Data>2</x:Data>\015\012      </x:FormatValue>\015\012"
        "      <x:Marker>\015\012       <x:Symbol>None</x:Symbol>\015\012      </x:Marker"
        ">\015\012      <x:Explode>0</x:Explode>\015\012      <x:Thickness>10</x:Thicknes"
        "s>\015\012      <x:DataSourceIndex>0</x:DataSourceIndex>\015\012      <x:Identif"
        "ier>!.Sum of مانده بدهكار</x:Identifier>\015\012     </x:Series>\015\012     <x:"
        "Series>\015\012      <x:FormatMap>\015\012      </x:FormatMap>\015\012      <x:N"
        "ame>Sum of مانده بستانكار</x:Name>\015\012      <x:Caption>\015\012       <x:Dat"
        "aSourceIndex>-1</x:DataSourceIndex>\015\012       <x:Data>&quot;Sum of مانده بست"
        "انكار&quot;</x:Data>\015\012      </x:Caption>\015\012      <x:Index>1</x:Index>"
        "\015\012      <x:Category>\015\012       <x:DataSourceIndex>0</x:DataSourceIndex"
        ">\015\012      </x:Category>\015\012      <x:Value>\015\012       <x:DataSourceI"
        "ndex>0</x:DataSourceIndex>\015\012      </x:Value>\015\012      <x:FormatValue>\015"
        "\012       <x:DataSourceIndex>-3</x:DataSourceIndex>\015\012       <x:Data>2</x:"
        "Data>\015\012      </x:FormatValue>\015\012      <x:DataLabels>\015\012       <x"
        ":Border>\015\012        <x:ColorIndex>None</x:ColorIndex>\015\012       </x:Bord"
        "er>\015\012       <x:Font>\015\012        <x:Color>#006400</x:Color>\015\012    "
        "    <x:B>Automatic</x:B>\015\012        <x:I>Automatic</x:I>\015\012        <x:U"
        ">Automatic</x:U>\015\012       </x:Font>\015\012       <x:Number>\015\012       "
        " <x:FormatString>#,###</x:FormatString>\015\012       </x:Number>\015\012       "
        "<x:ShowValue/>\015\012       <x:Separator>;</x:Separator>\015\012       <x:Posit"
        "ion>OutsideEnd</x:Position>\015\012      </x:DataLabels>\015\012      <x:Marker>"
        "\015\012       <x:Symbol>None</x:Symbol>\015\012      </x:Marker>\015\012      <"
        "x:Explode>0</x:Explode>\015\012      <x:Thickness>10</x:Thickness>\015\012      "
        "<x:DataSourceIndex>0</x:DataSourceIndex>\015\012      <x:Identifier>!.Sum of مان"
        "ده بستانكار</x:Identifier>\015\012     </x:Series>\015\012     <x:Dimension>\015"
        "\012      <x:ScaleID>373767512</x:ScaleID>\015\012      <x:Index>Categories</x:I"
        "ndex>\015\012     </x:Dimension>\015\012     <x:Dimension>\015\012      <x:Scale"
        "ID>373767732</x:ScaleID>\015\012      <x:Index>Value</x:Index>\015\012     </x:D"
        "imension>\015\012     <x:Dimension>\015\012      <x:ScaleID>373767952</x:ScaleID"
        ">\015\012      <x:Index>FormatValue</x:Index>\015\012     </x:Dimension>\015\012"
        "     <x:Overlap>0</x:Overlap>\015\012     <x:GapWidth>150</x:GapWidth>\015\012  "
        "   <x:FirstSliceAngle>0</x:FirstSliceAngle>\015\012    </x:Graph>\015\012    <x:"
        "Axis>\015\012     <x:AxisID>373772476</x:AxisID>\015\012     <x:ScaleID>37376751"
        "2</x:ScaleID>\015\012     <x:Type>Category</x:Type>\015\012     <x:MajorTick>Out"
        "side</x:MajorTick>\015\012     <x:MinorTick>None</x:MinorTick>\015\012     <x:Ti"
        "tle>\015\012      <x:Font>\015\012       <x:FontName>Arial</x:FontName>\015\012 "
        "      <x:Size>10</x:Size>\015\012      </x:Font>\015\012      <x:Border>\015\012"
        "       <x:ColorIndex>None</x:ColorIndex>\015\012      </x:Border>\015\012      <"
        "x:Caption>\015\012       <x:DataSourceIndex>-1</x:DataSourceIndex>\015\012      "
        " <x:Data>&quot;ماه&quot;</x:Data>\015\012      </x:Caption>\015\012      <x:Posi"
        "tion>Bottom</x:Position>\015\012     </x:Title>\015\012     <x:Placement>Bottom<"
        "/x:Placement>\015\012     <x:GroupingEnum>Auto</x:GroupingEnum>\015\012    </x:A"
        "xis>\015\012    <x:Axis>\015\012     <x:AxisID>373773208</x:AxisID>\015\012     "
        "<x:ScaleID>373767732</x:ScaleID>\015\012     <x:Type>Value</x:Type>\015\012     "
        "<x:Number>\015\012      <x:FormatString>#,###</x:FormatString>\015\012     </x:N"
        "umber>\015\012     <x:MajorGridlines>\015\012     </x:MajorGridlines>\015\012   "
        "  <x:MajorTick>Outside</x:MajorTick>\015\012     <x:MinorTick>None</x:MinorTick>"
        "\015\012     <x:Title>\015\012      <x:Font>\015\012       <x:FontName>Traffic</"
        "x:FontName>\015\012       <x:Size>10</x:Size>\015\012      </x:Font>\015\012    "
        "  <x:Border>\015\012       <x:ColorIndex>None</x:ColorIndex>\015\012      </x:Bo"
        "rder>\015\012      <x:Caption>\015\012       <x:DataSourceIndex>-1</x:DataSource"
        "Index>\015\012       <x:Data>&quot;مبلغ&quot;</x:Data>\015\012      </x:Caption>"
        "\015\012      <x:Position>Left</x:Position>\015\012     </x:Title>\015\012     <"
        "x:Placement>Left</x:Placement>\015\012    </x:Axis>\015\012   </x:PlotArea>\015\012"
        "   <x:Identifier></x:Identifier>\015\012  </x:Chart>\015\012  <x:Scaling>\015\012"
        "   <x:ScaleID>373767512</x:ScaleID>\015\012  </x:Scaling>\015\012  <x:Scaling>\015"
        "\012   <x:ScaleID>373767732</x:ScaleID>\015\012  </x:Scaling>\015\012  <x:Scalin"
        "g>\015\012   <x:ScaleID>373767952</x:ScaleID>\015\012  </x:Scaling>\015\012  <x:"
        "HasSelectionFeedback>True</x:HasSelectionFeedback>\015\012 </x:ChartSpace>\015\012"
        "</xml>"
    RecordSourceQualifier ="dbo"
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
                    Left =13718
                    Width =795
                    Height =285
                    BackColor =13434828
                    ForeColor =8388608
                    Name ="NUMBER Label"
                    Caption ="کل"
                    FontName ="Tahoma"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="NUMBER_Label"
                End
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =2
                    TextFontFamily =34
                    Left =8893
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
                    Left =1845
                    Width =1765
                    Height =300
                    BackColor =13421619
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
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =2
                    TextFontFamily =34
                    Left =11341
                    Width =645
                    Height =285
                    BackColor =13434828
                    ForeColor =8388608
                    Name ="Label24"
                    Caption ="ماه "
                    FontName ="Tahoma"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =2
                    TextFontFamily =34
                    Left =5498
                    Width =1020
                    Height =300
                    ForeColor =8388608
                    Name ="Label26"
                    Caption ="نام حساب"
                    FontName ="Tahoma"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =2
                    TextFontFamily =34
                    Left =12885
                    Width =825
                    Height =285
                    BackColor =13434828
                    ForeColor =8388608
                    Name ="Label28"
                    Caption ="معين "
                    FontName ="Tahoma"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =2
                    TextFontFamily =34
                    Left =12000
                    Width =855
                    Height =285
                    BackColor =13434828
                    ForeColor =8388608
                    Name ="Label30"
                    Caption ="تفصيلي "
                    FontName ="Tahoma"
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
                    Left =13749
                    Top =-3
                    Width =759
                    Height =315
                    ColumnWidth =1065
                    ColumnOrder =0
                    FontWeight =700
                    BackColor =13434828
                    Name ="شماره حساب"
                    ControlSource ="HES_K"
                    FontName ="Titr"
                    EventProcPrefix ="شماره_حساب"
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
                    Left =7353
                    Top =-3
                    Width =3924
                    Height =315
                    ColumnWidth =3330
                    ColumnOrder =3
                    FontWeight =700
                    TabIndex =1
                    Name ="نام حساب"
                    ControlSource ="kol"
                    FontName ="Titr"
                    EventProcPrefix ="نام_حساب"
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
                    Left =1844
                    Top =-3
                    Width =1764
                    Height =315
                    ColumnWidth =1800
                    ColumnOrder =5
                    FontWeight =700
                    TabIndex =2
                    BackColor =13421619
                    Name ="مانده بدهكار"
                    ControlSource ="BED"
                    Format ="Standard"
                    FontName ="Titr"
                    EventProcPrefix ="مانده_بدهكار"
                End
                Begin TextBox
                    DecimalPlaces =0
                    SpecialEffect =2
                    OldBorderStyle =0
                    OverlapFlags =87
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    BackStyle =1
                    NumeralShapes =2
                    Left =56
                    Top =-3
                    Width =1776
                    Height =315
                    ColumnWidth =1530
                    ColumnOrder =6
                    FontWeight =700
                    TabIndex =3
                    BackColor =16776960
                    Name ="مانده بستانكار"
                    ControlSource ="BES"
                    Format ="Standard"
                    FontName ="Titr"
                    EventProcPrefix ="مانده_بستانكار"
                End
                Begin Line
                    OverlapFlags =87
                    SpecialEffect =2
                    Left =14520
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
                    Left =11402
                    Top =-3
                    Width =549
                    Height =315
                    ColumnOrder =7
                    FontWeight =700
                    TabIndex =4
                    BackColor =13434828
                    Name ="ماه"
                    ControlSource ="MON"
                    FontName ="Titr"
                End
                Begin TextBox
                    SpecialEffect =2
                    OldBorderStyle =0
                    OverlapFlags =95
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    BackStyle =1
                    NumeralShapes =2
                    Left =3628
                    Top =-3
                    Width =3684
                    Height =315
                    ColumnWidth =3495
                    ColumnOrder =4
                    FontWeight =700
                    TabIndex =5
                    Name ="تفصيلي"
                    ControlSource ="TAFZ"
                    FontName ="Titr"
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
                    Left =12886
                    Top =-3
                    Width =819
                    Height =315
                    ColumnWidth =540
                    ColumnOrder =1
                    FontWeight =700
                    TabIndex =6
                    BackColor =13434828
                    Name ="معين"
                    ControlSource ="HES_M"
                    FontName ="Titr"
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
                    Left =12046
                    Top =-3
                    Width =759
                    Height =315
                    ColumnWidth =870
                    ColumnOrder =2
                    FontWeight =700
                    TabIndex =7
                    BackColor =13434828
                    Name ="تفصيلي 1"
                    ControlSource ="HES_T"
                    FontName ="Titr"
                    EventProcPrefix ="تفصيلي_1"
                End
                Begin TextBox
                    Visible = NotDefault
                    ColumnHidden = NotDefault
                    SpecialEffect =2
                    OldBorderStyle =0
                    OverlapFlags =247
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    BackStyle =1
                    NumeralShapes =2
                    Left =5952
                    Width =759
                    Height =315
                    FontWeight =700
                    TabIndex =8
                    BackColor =13434828
                    Name ="HES_K"
                    ControlSource ="HES_K"
                    FontName ="Titr"
                End
            End
        End
        Begin FormFooter
            Height =345
            BackColor =-2147483633
            Name ="FormFooter"
            Begin
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    TextFontFamily =34
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1856
                    Width =1749
                    Height =330
                    ForeColor =8388608
                    Name ="Text14"
                    ControlSource ="=Sum([bed])"
                    Format ="Fixed"
                    FontName ="Tahoma"
                    InputMask ="#,###"
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    TextFontFamily =34
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =56
                    Width =1761
                    Height =330
                    TabIndex =1
                    ForeColor =8388608
                    Name ="Text15"
                    ControlSource ="=Sum([bes])"
                    Format ="Fixed"
                    FontName ="Tahoma"
                End
            End
        End
    End
End
CodeBehindForm
' See "TARAZ_4_MAH_taf.cls"
