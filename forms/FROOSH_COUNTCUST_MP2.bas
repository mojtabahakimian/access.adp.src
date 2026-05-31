Version =20
VersionRequired =20
Begin Form
    KeyPreview = NotDefault
    DefaultView =2
    AllowUpdating =4
    ViewsAllowed =2
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =7875
    RowHeight =330
    ItemSuffix =36
    Left =750
    Top =5265
    Right =8070
    Bottom =10305
    DatasheetGridlinesColor =12632256
    MaxRecords =100000000
    Tag ="@Forms___F_MENU_ANBAR___DT2 BIGINT =[Forms]![F_MENU_ANBAR]![DT2],@Forms___F_MENU"
        "_ANBAR___MANBAR nvarchar(10)=[FormS]![F_MENU_ANBAR]![MANBAR]"
    ShortcutMenuBar ="MNEMUDAR"
    RecSrcDt = Begin
        0x29b3415efee3e340
    End
    RecordSource ="SELECT SUM(MEGHKF) AS MEGHF, SUM(MABK) AS MABKK, SUM(MEGHkK) AS MEGHk, SAIZE, TA"
        "RH, CAST(TARH AS INT) AS TTARH, CUST_NO FROM FROOSH_COUNTALL_CITY_N_FANI(1, 9999"
        "99) FROOSH_COUNTALL_CITY_N_FANI GROUP BY SAIZE, TARH, CAST(TARH AS INT), CUST_NO"
    Caption ="گزارش فروش به تفكيك مشتري /سايز"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Tahoma"
    OnKeyPress ="[Event Procedure]"
    OnError ="[Event Procedure]"
    PivotTableXML ="<xml xmlns:x=\"urn:schemas-microsoft-com:office:excel\">\015\012 <x:WorksheetOpt"
        "ions>\015\012  <x:DisplayRightToLeft/>\015\012 </x:WorksheetOptions>\015\012 <x:"
        "PivotTable>\015\012  <x:OWCVersion>10.0.0.6854         </x:OWCVersion>\015\012  "
        "<x:NoDisplayAlerts/>\015\012  <x:NoDisplayToolbar/>\015\012  <x:DisplayScreenTip"
        "s/>\015\012  <x:NoAutoFit/>\015\012  <x:Height>384</x:Height>\015\012  <x:Width>"
        "541</x:Width>\015\012  <x:NoAllowGrouping/>\015\012  <x:NoAllowFiltering/>\015\012"
        "  <x:CubeProvider>msolap.2</x:CubeProvider>\015\012  <x:CacheDetails/>\015\012  "
        "<x:Name>Microsoft Office PivotTable 10.0</x:Name>\015\012  <x:DataAxisEmpty/>\015"
        "\012  <x:PivotField>\015\012   <x:Name>موجودي</x:Name>\015\012   <x:SourceName>M"
        "ANDS</x:SourceName>\015\012   <x:DataType>Number</x:DataType>\015\012   <x:Encod"
        "edType>adDouble</x:EncodedType>\015\012   <x:NumberFormat>Fixed</x:NumberFormat>"
        "\015\012   <x:CompareOrderedMembersBy>UniqueName</x:CompareOrderedMembersBy>\015"
        "\012  </x:PivotField>\015\012  <x:PivotField>\015\012   <x:Name>SAIZE</x:Name>\015"
        "\012   <x:EncodedType>adVarWChar</x:EncodedType>\015\012   <x:CompareOrderedMemb"
        "ersBy>UniqueName</x:CompareOrderedMembersBy>\015\012   <x:NoDisplayInFieldList/>"
        "\015\012  </x:PivotField>\015\012  <x:PivotField>\015\012   <x:Name>كد انبار</x:"
        "Name>\015\012   <x:SourceName>ANBAR</x:SourceName>\015\012   <x:DataType>Number<"
        "/x:DataType>\015\012   <x:EncodedType>adDouble</x:EncodedType>\015\012   <x:Comp"
        "areOrderedMembersBy>UniqueName</x:CompareOrderedMembersBy>\015\012   <x:NoDispla"
        "yInFieldList/>\015\012  </x:PivotField>\015\012  <x:PivotField>\015\012   <x:Nam"
        "e>كارتن</x:Name>\015\012   <x:SourceName>MANDFS</x:SourceName>\015\012   <x:Data"
        "Type>Number</x:DataType>\015\012   <x:EncodedType>adDouble</x:EncodedType>\015\012"
        "   <x:NumberFormat>Fixed</x:NumberFormat>\015\012   <x:CompareOrderedMembersBy>U"
        "niqueName</x:CompareOrderedMembersBy>\015\012  </x:PivotField>\015\012  <x:Pivot"
        "Field>\015\012   <x:Name>انبار</x:Name>\015\012   <x:SourceName>ANBARN</x:Source"
        "Name>\015\012   <x:EncodedType>adVarWChar</x:EncodedType>\015\012   <x:CompareOr"
        "deredMembersBy>UniqueName</x:CompareOrderedMembersBy>\015\012   <x:NoDisplayInFi"
        "eldList/>\015\012  </x:PivotField>\015\012  <x:PivotField>\015\012   <x:Name>نام"
        " واحد</x:Name>\015\012   <x:SourceName>NAMES</x:SourceName>\015\012   <x:Encoded"
        "Type>adVarWChar</x:EncodedType>\015\012   <x:CompareOrderedMembersBy>UniqueName<"
        "/x:CompareOrderedMembersBy>\015\012  </x:PivotField>\015\012  <x:PivotField>\015"
        "\012   <x:Name>طرح</x:Name>\015\012   <x:SourceName>TARH</x:SourceName>\015\012 "
        "  <x:Orientation>Row</x:Orientation>\015\012   <x:EncodedType>adVarWChar</x:Enco"
        "dedType>\015\012   <x:CompareOrderedMembersBy>UniqueName</x:CompareOrderedMember"
        "sBy>\015\012   <x:Expanded/>\015\012  </x:PivotField>\015\012  <x:PivotField>\015"
        "\012   <x:Name>طرح1</x:Name>\015\012   <x:SourceName>ttarh</x:SourceName>\015\012"
        "   <x:DataType>Integer</x:DataType>\015\012   <x:PLCaption>طرح</x:PLCaption>\015"
        "\012   <x:FilterCaption>طرح</x:FilterCaption>\015\012   <x:EncodedType>adInteger"
        "</x:EncodedType>\015\012   <x:CompareOrderedMembersBy>UniqueName</x:CompareOrder"
        "edMembersBy>\015\012  </x:PivotField>\015\012  <x:PivotField>\015\012   <x:Name>"
        "Data</x:Name>\015\012   <x:Orientation>Column</x:Orientation>\015\012   <x:Posit"
        "ion>-1</x:Position>\015\012   <x:DataField/>\015\012  </x:PivotField>\015\012  <"
        "x:PivotField>\015\012   <x:Name>Sum of كارتن</x:Name>\015\012   <x:PLName>Total2"
        "</x:PLName>\015\012   <x:TotalNumber>0</x:TotalNumber>\015\012   <x:Orientation>"
        "Data</x:Orientation>\015\012   <x:Position>1</x:Position>\015\012   <x:ParentFie"
        "ld>كارتن</x:ParentField>\015\012   <x:NumberFormat>Fixed</x:NumberFormat>\015\012"
        "  </x:PivotField>\015\012  <x:PivotData>\015\012   <x:Top>0.0</x:Top>\015\012   "
        "<x:TopOffset>0</x:TopOffset>\015\012   <x:Left>0</x:Left>\015\012   <x:LeftOffse"
        "t>0</x:LeftOffset>\015\012   <x:SeqNum>38</x:SeqNum>\015\012  </x:PivotData>\015"
        "\012  <x:PivotView>\015\012   <x:IsNotFiltered/>\015\012   <x:Label>\015\012    "
        "<x:Caption>Microsoft Office PivotTable 10.0</x:Caption>\015\012    <x:NotVisible"
        "/>\015\012   </x:Label>\015\012   <x:AllowEdits>false</x:AllowEdits>\015\012   <"
        "x:AllowAdditions>false</x:AllowAdditions>\015\012   <x:AllowDeletions>false</x:A"
        "llowDeletions>\015\012  </x:PivotView>\015\012 </x:PivotTable>\015\012</xml>"
    PivotChartXML ="<xml xmlns:x=\"urn:schemas-microsoft-com:office:excel\">\015\012 <x:ChartSpace>\015"
        "\012  <x:OWCVersion>10.0.0.6854         </x:OWCVersion>\015\012  <x:Width>26988<"
        "/x:Width>\015\012  <x:Height>15161</x:Height>\015\012  <x:DataSource>\015\012   "
        "<x:Type>PivotList</x:Type>\015\012  </x:DataSource>\015\012  <x:BoundSeries>\015"
        "\012   <x:DataSourceIndex>0</x:DataSourceIndex>\015\012  </x:BoundSeries>\015\012"
        "  <x:Category>\015\012   <x:DataSourceIndex>0</x:DataSourceIndex>\015\012  </x:C"
        "ategory>\015\012  <x:Value>\015\012   <x:DataSourceIndex>0</x:DataSourceIndex>\015"
        "\012   <x:Data>Total2</x:Data>\015\012  </x:Value>\015\012  <x:BoundCharts>\015\012"
        "   <x:DataSourceIndex>0</x:DataSourceIndex>\015\012  </x:BoundCharts>\015\012  <"
        "x:PivotAggOrientation>None</x:PivotAggOrientation>\015\012  <x:HidePivotFields/>"
        "\015\012  <x:NoGrouping/>\015\012  <x:NoFiltering/>\015\012  <x:Palette>\015\012"
        "   <x:Entry>#000000</x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015\012   <x:"
        "Entry>#000000</x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015\012   <x:Entry>"
        "#000000</x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015\012   <x:Entry>#00000"
        "0</x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015\012   <x:Entry>#000000</x:E"
        "ntry>\015\012   <x:Entry>#000000</x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015"
        "\012   <x:Entry>#000000</x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015\012  "
        " <x:Entry>#000000</x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015\012   <x:En"
        "try>#000000</x:Entry>\015\012   <x:Entry>#8080FF</x:Entry>\015\012   <x:Entry>#8"
        "02060</x:Entry>\015\012   <x:Entry>#FFFFA0</x:Entry>\015\012   <x:Entry>#A0E0E0<"
        "/x:Entry>\015\012   <x:Entry>#600080</x:Entry>\015\012   <x:Entry>#FF8080</x:Ent"
        "ry>\015\012   <x:Entry>#008080</x:Entry>\015\012   <x:Entry>#C0C0FF</x:Entry>\015"
        "\012   <x:Entry>#000080</x:Entry>\015\012   <x:Entry>#FF00FF</x:Entry>\015\012  "
        " <x:Entry>#80FFFF</x:Entry>\015\012   <x:Entry>#0080FF</x:Entry>\015\012   <x:En"
        "try>#FF8080</x:Entry>\015\012   <x:Entry>#C0FF80</x:Entry>\015\012   <x:Entry>#F"
        "FC0FF</x:Entry>\015\012   <x:Entry>#FF80FF</x:Entry>\015\012  </x:Palette>\015\012"
        "  <x:DefaultFont>Arial</x:DefaultFont>\015\012  <x:Border>\015\012   <x:Color>#2"
        "F4F4F</x:Color>\015\012  </x:Border>\015\012  <x:Interior>\015\012   <x:Color>#F"
        "FFFFF</x:Color>\015\012   <x:FillEffect>\015\012    <x:fill x:type=\"Solid\" x:c"
        "olor=\"#FFFFFF\"/>\015\012   </x:FillEffect>\015\012  </x:Interior>\015\012  <x:"
        "Chart>\015\012   <x:PlotArea>\015\012    <x:Graph>\015\012     <x:SubType>Cluste"
        "red</x:SubType>\015\012     <x:Type>Column</x:Type>\015\012     <x:SubType>3D</x"
        ":SubType>\015\012     <x:Series>\015\012      <x:FormatMap>\015\012      </x:For"
        "matMap>\015\012      <x:Name>Series</x:Name>\015\012      <x:Caption>\015\012   "
        "    <x:DataSourceIndex>-1</x:DataSourceIndex>\015\012       <x:Data>&quot;Series"
        "&quot;</x:Data>\015\012      </x:Caption>\015\012      <x:Index>0</x:Index>\015\012"
        "      <x:Category>\015\012       <x:DataSourceIndex>0</x:DataSourceIndex>\015\012"
        "      </x:Category>\015\012      <x:Value>\015\012       <x:DataSourceIndex>0</x"
        ":DataSourceIndex>\015\012       <x:Data>Total2</x:Data>\015\012      </x:Value>\015"
        "\012      <x:DataLabels>\015\012       <x:Border>\015\012        <x:ColorIndex>N"
        "one</x:ColorIndex>\015\012       </x:Border>\015\012       <x:Number>\015\012   "
        "     <x:FormatString>General</x:FormatString>\015\012       </x:Number>\015\012 "
        "      <x:ShowValue/>\015\012       <x:Separator>;</x:Separator>\015\012       <x"
        ":Position>Center</x:Position>\015\012       <x:LabelEntry>\015\012        <x:Ind"
        "ex>0</x:Index>\015\012        <x:Visible>False</x:Visible>\015\012        <x:Num"
        "ber>\015\012         <x:FormatString>General</x:FormatString>\015\012        </x"
        ":Number>\015\012       </x:LabelEntry>\015\012       <x:LabelEntry>\015\012     "
        "   <x:Index>1</x:Index>\015\012        <x:Visible>False</x:Visible>\015\012     "
        "   <x:Number>\015\012         <x:FormatString>General</x:FormatString>\015\012  "
        "      </x:Number>\015\012       </x:LabelEntry>\015\012      </x:DataLabels>\015"
        "\012      <x:DataLabels>\015\012       <x:Border>\015\012        <x:ColorIndex>N"
        "one</x:ColorIndex>\015\012       </x:Border>\015\012       <x:Font>\015\012     "
        "   <x:Size>10</x:Size>\015\012        <x:Color>#FF0000</x:Color>\015\012        "
        "<x:B/>\015\012        <x:I>Automatic</x:I>\015\012        <x:U>Automatic</x:U>\015"
        "\012       </x:Font>\015\012       <x:Number>\015\012        <x:FormatString>Gen"
        "eral</x:FormatString>\015\012       </x:Number>\015\012       <x:ShowValue/>\015"
        "\012       <x:Separator>;</x:Separator>\015\012       <x:Position>Center</x:Posi"
        "tion>\015\012       <x:LabelEntry>\015\012        <x:Index>0</x:Index>\015\012  "
        "      <x:Visible>False</x:Visible>\015\012        <x:Font>\015\012         <x:Si"
        "ze>10</x:Size>\015\012         <x:Color>#FF0000</x:Color>\015\012         <x:B/>"
        "\015\012         <x:I>Automatic</x:I>\015\012         <x:U>Automatic</x:U>\015\012"
        "        </x:Font>\015\012        <x:Number>\015\012         <x:FormatString>Gene"
        "ral</x:FormatString>\015\012        </x:Number>\015\012       </x:LabelEntry>\015"
        "\012       <x:LabelEntry>\015\012        <x:Index>1</x:Index>\015\012        <x:"
        "Visible>False</x:Visible>\015\012        <x:Font>\015\012         <x:Size>10</x:"
        "Size>\015\012         <x:Color>#FF0000</x:Color>\015\012         <x:B/>\015\012 "
        "        <x:I>Automatic</x:I>\015\012         <x:U>Automatic</x:U>\015\012       "
        " </x:Font>\015\012        <x:Number>\015\012         <x:FormatString>General</x:"
        "FormatString>\015\012        </x:Number>\015\012       </x:LabelEntry>\015\012  "
        "    </x:DataLabels>\015\012      <x:DataLabels>\015\012       <x:Border>\015\012"
        "        <x:ColorIndex>None</x:ColorIndex>\015\012       </x:Border>\015\012     "
        "  <x:Font>\015\012        <x:Size>10</x:Size>\015\012        <x:Color>#FF0000</x"
        ":Color>\015\012        <x:B/>\015\012        <x:I>Automatic</x:I>\015\012       "
        " <x:U>Automatic</x:U>\015\012       </x:Font>\015\012       <x:Number>\015\012  "
        "      <x:FormatString>General</x:FormatString>\015\012       </x:Number>\015\012"
        "       <x:ShowValue/>\015\012       <x:Separator>;</x:Separator>\015\012       <"
        "x:Position>OutsideEnd</x:Position>\015\012      </x:DataLabels>\015\012      <x:"
        "Marker>\015\012       <x:Symbol>None</x:Symbol>\015\012      </x:Marker>\015\012"
        "      <x:Explode>10</x:Explode>\015\012      <x:Thickness>10</x:Thickness>\015\012"
        "      <x:DataSourceIndex>0</x:DataSourceIndex>\015\012      <x:Identifier>!</x:I"
        "dentifier>\015\012     </x:Series>\015\012     <x:Dimension>\015\012      <x:Sca"
        "leID>372712120</x:ScaleID>\015\012      <x:Index>Categories</x:Index>\015\012   "
        "  </x:Dimension>\015\012     <x:Dimension>\015\012      <x:ScaleID>372712324</x:"
        "ScaleID>\015\012      <x:Index>Value</x:Index>\015\012     </x:Dimension>\015\012"
        "     <x:Dimension>\015\012      <x:ScaleID>372707772</x:ScaleID>\015\012      <x"
        ":Index>FormatValue</x:Index>\015\012     </x:Dimension>\015\012     <x:Dimension"
        ">\015\012      <x:ScaleID>372721760</x:ScaleID>\015\012      <x:Index>Series</x:"
        "Index>\015\012     </x:Dimension>\015\012     <x:Overlap>0</x:Overlap>\015\012  "
        "   <x:GapWidth>150</x:GapWidth>\015\012     <x:FirstSliceAngle>0</x:FirstSliceAn"
        "gle>\015\012    </x:Graph>\015\012    <x:Axis>\015\012     <x:AxisID>372711240</"
        "x:AxisID>\015\012     <x:ScaleID>372712120</x:ScaleID>\015\012     <x:Type>Categ"
        "ory</x:Type>\015\012     <x:MajorTick>Outside</x:MajorTick>\015\012     <x:Minor"
        "Tick>None</x:MinorTick>\015\012     <x:Placement>Bottom</x:Placement>\015\012   "
        "  <x:GroupingEnum>Auto</x:GroupingEnum>\015\012    </x:Axis>\015\012    <x:Axis>"
        "\015\012     <x:AxisID>372710508</x:AxisID>\015\012     <x:ScaleID>372712324</x:"
        "ScaleID>\015\012     <x:Type>Value</x:Type>\015\012     <x:Number>\015\012      "
        "<x:FormatString>Fixed</x:FormatString>\015\012     </x:Number>\015\012     <x:Ma"
        "jorGridlines>\015\012     </x:MajorGridlines>\015\012     <x:MajorTick>Outside</"
        "x:MajorTick>\015\012     <x:MinorTick>None</x:MinorTick>\015\012     <x:Placemen"
        "t>Left</x:Placement>\015\012    </x:Axis>\015\012   </x:PlotArea>\015\012   <x:V"
        "iew3D>\015\012    <x:GapDepth>150</x:GapDepth>\015\012    <x:Perspective>20.0</x"
        ":Perspective>\015\012    <x:Rotation>20.0</x:Rotation>\015\012    <x:Inclination"
        ">15.0</x:Inclination>\015\012    <x:Light>\015\012     <x:Rotation>315.0</x:Rota"
        "tion>\015\012     <x:Inclination>15.0</x:Inclination>\015\012     <x:IntensityDi"
        "ffuse>0.549019607843137</x:IntensityDiffuse>\015\012     <x:IntensityAmbient>0.6"
        "19607843137255</x:IntensityAmbient>\015\012     <x:Normal>0.5</x:Normal>\015\012"
        "    </x:Light>\015\012   </x:View3D>\015\012   <x:Identifier></x:Identifier>\015"
        "\012   <x:Walls>\015\012    <x:Index>0</x:Index>\015\012    <x:Thickness>6</x:Th"
        "ickness>\015\012   </x:Walls>\015\012   <x:Walls>\015\012    <x:Index>1</x:Index"
        ">\015\012    <x:Thickness>6</x:Thickness>\015\012   </x:Walls>\015\012   <x:Wall"
        "s>\015\012    <x:Index>2</x:Index>\015\012    <x:Thickness>6</x:Thickness>\015\012"
        "   </x:Walls>\015\012  </x:Chart>\015\012  <x:Title>\015\012   <x:Font>\015\012 "
        "   <x:FontName>Sina</x:FontName>\015\012    <x:Color>#228B22</x:Color>\015\012  "
        "  <x:B>Automatic</x:B>\015\012    <x:I>Automatic</x:I>\015\012    <x:U>Automatic"
        "</x:U>\015\012   </x:Font>\015\012   <x:Border>\015\012    <x:ColorIndex>None</x"
        ":ColorIndex>\015\012   </x:Border>\015\012   <x:Caption>\015\012    <x:DataSourc"
        "eIndex>-1</x:DataSourceIndex>\015\012    <x:Data>&quot;نمودار موجودي كالا ها&quo"
        "t;</x:Data>\015\012   </x:Caption>\015\012  </x:Title>\015\012  <x:Scaling>\015\012"
        "   <x:ScaleID>372712120</x:ScaleID>\015\012  </x:Scaling>\015\012  <x:Scaling>\015"
        "\012   <x:ScaleID>372712324</x:ScaleID>\015\012  </x:Scaling>\015\012  <x:Scalin"
        "g>\015\012   <x:ScaleID>372707772</x:ScaleID>\015\012  </x:Scaling>\015\012  <x:"
        "Scaling>\015\012   <x:ScaleID>372721760</x:ScaleID>\015\012  </x:Scaling>\015\012"
        "  <x:HidePassiveAlert/>\015\012  <x:HideScreenTips/>\015\012 </x:ChartSpace>\015"
        "\012</xml>"
    AllowFormView =0
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
        Begin OptionButton
            SpecialEffect =2
            LabelX =230
            LabelY =-30
        End
        Begin CheckBox
            SpecialEffect =2
            LabelX =230
            LabelY =-30
        End
        Begin OptionGroup
            SpecialEffect =3
            Width =1701
            Height =1701
        End
        Begin BoundObjectFrame
            SpecialEffect =2
            Width =4536
            Height =2835
            LabelX =-1701
        End
        Begin TextBox
            SpecialEffect =2
            OldBorderStyle =0
            TextFontCharSet =178
            Width =1701
            LabelX =-1701
            FontName ="Tahoma"
        End
        Begin ListBox
            SpecialEffect =2
            TextFontCharSet =178
            Width =1701
            Height =1417
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
        Begin Subform
            SpecialEffect =2
            Width =1701
            Height =1701
        End
        Begin Section
            CanGrow = NotDefault
            Height =9391
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =1
                    NumeralShapes =2
                    Left =3895
                    Top =1845
                    ColumnWidth =465
                    ColumnOrder =0
                    Name ="TARH"
                    ControlSource ="TARH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6652
                            Top =1845
                            Width =645
                            Height =240
                            Name ="Label31"
                            Caption ="طرح"
                        End
                    End
                End
                Begin Subform
                    OverlapFlags =87
                    Top =4081
                    Width =6165
                    Height =5310
                    TabIndex =1
                    Name ="FROOSH_COUNTCUST_MP3"
                    SourceObject ="Form.FROOSH_COUNTCUST_MP3"
                    LinkChildFields ="SAIZE;TARH;CUST_NO"
                    LinkMasterFields ="SAIZE;TARH;CUST_NO"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =3795
                            Top =3841
                            Width =2370
                            Height =240
                            Name ="AK_MOGUDI_ANBAR_LIST_MP3 Label"
                            Caption ="AK_MOGUDI_ANBAR_LIST_MP3"
                            EventProcPrefix ="AK_MOGUDI_ANBAR_LIST_MP3_Label"
                        End
                    End
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =3895
                    Top =2205
                    ColumnWidth =1785
                    ColumnOrder =1
                    TabIndex =2
                    Name ="ttarh"
                    ControlSource ="ttarh"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT MPCODE, MPNAME FROM TCOD_MAP WHERE (MPP = 3)"
                    ColumnWidths ="0"
                    OnDblClick ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6652
                            Top =2205
                            Width =645
                            Height =240
                            Name ="Label35"
                            Caption ="طرح"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =2
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =1
                    NumeralShapes =2
                    Left =49
                    Top =3
                    ColumnOrder =2
                    TabIndex =3
                    Name ="MEGHk"
                    ControlSource ="MEGHk"
                    Format ="Fixed"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3106
                            Width =645
                            Height =240
                            Name ="Label3"
                            Caption ="متراژ"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =1
                    NumeralShapes =2
                    Left =46
                    Top =345
                    ColumnOrder =5
                    TabIndex =4
                    Name ="MABKK"
                    ControlSource ="MABKK"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2803
                            Top =345
                            Width =645
                            Height =240
                            Name ="Label18"
                            Caption ="مبلغ"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    SpecialEffect =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =2
                    ScrollBarAlign =1
                    NumeralShapes =2
                    Top =1880
                    ColumnOrder =3
                    TabIndex =5
                    BackColor =13434828
                    Name ="MANDF"
                    ControlSource ="MEGHF"
                    Format ="Fixed"
                    ConditionalFormat = Begin
                        0x0100000066000000010000000000000004000000000000000200000001000000 ,
                        0x00000000ffcc9900000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x300000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =3055
                            Top =1871
                            Width =1035
                            Height =240
                            Name ="Label26"
                            Caption ="كارتن"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    ColumnHidden = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =55
                    Top =1092
                    ColumnOrder =4
                    TabIndex =6
                    Name ="SAIZE"
                    ControlSource ="SAIZE"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT MPCODE, MPNAME FROM TCOD_MAP WHERE (MPP = 1)"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2812
                            Top =1092
                            Width =645
                            Height =240
                            Name ="Label29"
                            Caption ="سايز"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =283
                    Top =2777
                    TabIndex =7
                    Name ="CUST_NO"
                    ControlSource ="CUST_NO"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2845
                            Top =2777
                            Width =840
                            Height =240
                            Name ="Label36"
                            Caption ="مشتري"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "FROOSH_COUNTCUST_MP2.cls"
