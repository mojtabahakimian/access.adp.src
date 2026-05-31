Version =20
VersionRequired =20
Begin Form
    OrderByOn = NotDefault
    DefaultView =2
    AllowUpdating =4
    ViewsAllowed =2
    TabularCharSet =178
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =10034
    RowHeight =360
    ItemSuffix =38
    Left =495
    Top =600
    Right =7470
    Bottom =7455
    DatasheetForeColor =33554432
    DatasheetGridlinesColor =12632256
    MaxRecords =1000000
    OrderBy ="VISITORS_KOL.GHABEL DESC"
    RecSrcDt = Begin
        0x4025adfb334de440
    End
    InputParameters ="@DT1 = FORMS![F_MENU_DATE_VISIT]![DT1],@DT2= FORMS![F_MENU_DATE_VISIT]![DT2],@VI"
        "SITOR= FORMS![F_MENU_DATE_VISIT]![CUST_NO]"
    RecordSource ="VISITORS_KOL"
    Caption ="عملكرد ويزيتورها"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Tahoma"
    UnknownProp = {258 ,66 ,12 ,4 ,244 } Begin
        0x400044005400310020003d00200046004f0052004d00530021005b0046005f00 ,
        0x4d0045004e0055005f0044004100540045005f00560049005300490054005d00 ,
        0x21005b004400540031005d002c0040004400540032003d00200046004f005200 ,
        0x4d00530021005b0046005f004d0045004e0055005f0044004100540045005f00 ,
        0x560049005300490054005d0021005b004400540032005d002c00400056004900 ,
        0x5300490054004f0052003d00200046004f0052004d00530021005b0046005f00 ,
        0x4d0045004e0055005f0044004100540045005f00560049005300490054005d00 ,
        0x21005b0043005500530054005f004e004f005d00
    End
    PivotTableXML ="<xml xmlns:x=\"urn:schemas-microsoft-com:office:excel\">\015\012 <x:WorksheetOpt"
        "ions>\015\012  <x:DisplayRightToLeft/>\015\012 </x:WorksheetOptions>\015\012 <x:"
        "PivotTable>\015\012  <x:OWCVersion>10.0.0.5605         </x:OWCVersion>\015\012  "
        "<x:NoDisplayAlerts/>\015\012  <x:NoDisplayToolbar/>\015\012  <x:DisplayScreenTip"
        "s/>\015\012  <x:NoAutoFit/>\015\012  <x:Height>384</x:Height>\015\012  <x:Width>"
        "541</x:Width>\015\012  <x:NoAllowGrouping/>\015\012  <x:NoAllowFiltering/>\015\012"
        "  <x:CubeProvider>msolap.2</x:CubeProvider>\015\012  <x:CacheDetails/>\015\012  "
        "<x:Name>Microsoft Office PivotTable 10.0</x:Name>\015\012  <x:PivotField>\015\012"
        "   <x:Name>كد ويزيتور</x:Name>\015\012   <x:SourceName>CUST_NO</x:SourceName>\015"
        "\012   <x:EncodedType>adVarWChar</x:EncodedType>\015\012   <x:CompareOrderedMemb"
        "ersBy>UniqueName</x:CompareOrderedMembersBy>\015\012  </x:PivotField>\015\012  <"
        "x:PivotField>\015\012   <x:Name>جمع پورسانت</x:Name>\015\012   <x:SourceName>mab"
        "pur</x:SourceName>\015\012   <x:DataType>Number</x:DataType>\015\012   <x:PLData"
        "Orientation/>\015\012   <x:PLPosition>2</x:PLPosition>\015\012   <x:EncodedType>"
        "adDouble</x:EncodedType>\015\012   <x:NumberFormat>Standard</x:NumberFormat>\015"
        "\012   <x:CompareOrderedMembersBy>UniqueName</x:CompareOrderedMembersBy>\015\012"
        "  </x:PivotField>\015\012  <x:PivotField>\015\012   <x:Name>مبلغ كل فروش</x:Name"
        ">\015\012   <x:SourceName>MABL_K</x:SourceName>\015\012   <x:DataType>Number</x:"
        "DataType>\015\012   <x:PLDataOrientation/>\015\012   <x:EncodedType>adDouble</x:"
        "EncodedType>\015\012   <x:NumberFormat>Standard</x:NumberFormat>\015\012   <x:Co"
        "mpareOrderedMembersBy>UniqueName</x:CompareOrderedMembersBy>\015\012  </x:PivotF"
        "ield>\015\012  <x:PivotField>\015\012   <x:Name>مقدار</x:Name>\015\012   <x:Sour"
        "ceName>MEGHk</x:SourceName>\015\012   <x:DataType>Number</x:DataType>\015\012   "
        "<x:EncodedType>adDouble</x:EncodedType>\015\012   <x:CompareOrderedMembersBy>Uni"
        "queName</x:CompareOrderedMembersBy>\015\012  </x:PivotField>\015\012  <x:PivotFi"
        "eld>\015\012   <x:Name>مرجوعي</x:Name>\015\012   <x:SourceName>MEGH_MAR</x:Sourc"
        "eName>\015\012   <x:DataType>Number</x:DataType>\015\012   <x:EncodedType>adDoub"
        "le</x:EncodedType>\015\012   <x:CompareOrderedMembersBy>UniqueName</x:CompareOrd"
        "eredMembersBy>\015\012  </x:PivotField>\015\012  <x:PivotField>\015\012   <x:Nam"
        "e>جمع تخفيف</x:Name>\015\012   <x:SourceName>TAKHF</x:SourceName>\015\012   <x:D"
        "ataType>Number</x:DataType>\015\012   <x:EncodedType>adDouble</x:EncodedType>\015"
        "\012   <x:CompareOrderedMembersBy>UniqueName</x:CompareOrderedMembersBy>\015\012"
        "  </x:PivotField>\015\012  <x:PivotField>\015\012   <x:Name>ارزش افزوده</x:Name>"
        "\015\012   <x:SourceName>MBAA</x:SourceName>\015\012   <x:DataType>Number</x:Dat"
        "aType>\015\012   <x:EncodedType>adDouble</x:EncodedType>\015\012   <x:NumberForm"
        "at>Standard</x:NumberFormat>\015\012   <x:CompareOrderedMembersBy>UniqueName</x:"
        "CompareOrderedMembersBy>\015\012  </x:PivotField>\015\012  <x:PivotField>\015\012"
        "   <x:Name>مبلغ مرجوعي</x:Name>\015\012   <x:SourceName>MABMAR</x:SourceName>\015"
        "\012   <x:DataType>Number</x:DataType>\015\012   <x:EncodedType>adDouble</x:Enco"
        "dedType>\015\012   <x:NumberFormat>Standard</x:NumberFormat>\015\012   <x:Compar"
        "eOrderedMembersBy>UniqueName</x:CompareOrderedMembersBy>\015\012  </x:PivotField"
        ">\015\012  <x:PivotField>\015\012   <x:Name>فروش خالص</x:Name>\015\012   <x:Sour"
        "ceName>GHABEL</x:SourceName>\015\012   <x:DataType>Number</x:DataType>\015\012  "
        " <x:EncodedType>adDouble</x:EncodedType>\015\012   <x:NumberFormat>Standard</x:N"
        "umberFormat>\015\012   <x:CompareOrderedMembersBy>UniqueName</x:CompareOrderedMe"
        "mbersBy>\015\012  </x:PivotField>\015\012  <x:PivotField>\015\012   <x:Name>نام "
        "ويزيتور</x:Name>\015\012   <x:SourceName>NAME</x:SourceName>\015\012   <x:Orient"
        "ation>Row</x:Orientation>\015\012   <x:EncodedType>adVarWChar</x:EncodedType>\015"
        "\012   <x:CompareOrderedMembersBy>UniqueName</x:CompareOrderedMembersBy>\015\012"
        "   <x:Expanded/>\015\012  </x:PivotField>\015\012  <x:PivotField>\015\012   <x:N"
        "ame>آدرس</x:Name>\015\012   <x:SourceName>ADDRESS</x:SourceName>\015\012   <x:En"
        "codedType>adVarWChar</x:EncodedType>\015\012   <x:CompareOrderedMembersBy>Unique"
        "Name</x:CompareOrderedMembersBy>\015\012  </x:PivotField>\015\012  <x:PivotField"
        ">\015\012   <x:Name>تلفن</x:Name>\015\012   <x:SourceName>TEL</x:SourceName>\015"
        "\012   <x:EncodedType>adVarWChar</x:EncodedType>\015\012   <x:CompareOrderedMemb"
        "ersBy>UniqueName</x:CompareOrderedMembersBy>\015\012  </x:PivotField>\015\012  <"
        "x:PivotField>\015\012   <x:Name>توضيح</x:Name>\015\012   <x:SourceName>TOZIH</x:"
        "SourceName>\015\012   <x:EncodedType>adVarWChar</x:EncodedType>\015\012   <x:Com"
        "pareOrderedMembersBy>UniqueName</x:CompareOrderedMembersBy>\015\012  </x:PivotFi"
        "eld>\015\012  <x:PivotField>\015\012   <x:Name>موبايل</x:Name>\015\012   <x:Sour"
        "ceName>MOBILE</x:SourceName>\015\012   <x:EncodedType>adVarWChar</x:EncodedType>"
        "\015\012   <x:CompareOrderedMembersBy>UniqueName</x:CompareOrderedMembersBy>\015"
        "\012  </x:PivotField>\015\012  <x:PivotField>\015\012   <x:Name>تعداد فاكتور</x:"
        "Name>\015\012   <x:SourceName>tedad</x:SourceName>\015\012   <x:DataType>Integer"
        "</x:DataType>\015\012   <x:EncodedType>adInteger</x:EncodedType>\015\012   <x:Co"
        "mpareOrderedMembersBy>UniqueName</x:CompareOrderedMembersBy>\015\012  </x:PivotF"
        "ield>\015\012  <x:PivotField>\015\012   <x:Name>درصد</x:Name>\015\012   <x:Encod"
        "edType>adDouble</x:EncodedType>\015\012   <x:CompareOrderedMembersBy>UniqueName<"
        "/x:CompareOrderedMembersBy>\015\012   <x:FieldSetName>DARSAD</x:FieldSetName>\015"
        "\012   <x:CalculatedFieldName>DARSAD</x:CalculatedFieldName>\015\012   <x:Calcul"
        "atedFieldExpression>IIf([GHABEL]&lt;&gt;0,[mabpur]/[GHABEL]*100,0)</x:Calculated"
        "FieldExpression>\015\012  </x:PivotField>\015\012  <x:PivotField>\015\012   <x:N"
        "ame>Data</x:Name>\015\012   <x:Orientation>Column</x:Orientation>\015\012   <x:P"
        "osition>-1</x:Position>\015\012   <x:DataField/>\015\012  </x:PivotField>\015\012"
        "  <x:PivotField>\015\012   <x:Name>Sum of مبلغ كل فروش</x:Name>\015\012   <x:PLN"
        "ame>Total1</x:PLName>\015\012   <x:TotalNumber>0</x:TotalNumber>\015\012   <x:Or"
        "ientation>Data</x:Orientation>\015\012   <x:Position>1</x:Position>\015\012   <x"
        ":ParentField>مبلغ كل فروش</x:ParentField>\015\012   <x:NumberFormat>Standard</x:"
        "NumberFormat>\015\012  </x:PivotField>\015\012  <x:PivotField>\015\012   <x:Name"
        ">Sum of جمع پورسانت</x:Name>\015\012   <x:PLName>Total2</x:PLName>\015\012   <x:"
        "TotalNumber>1</x:TotalNumber>\015\012   <x:Orientation>Data</x:Orientation>\015\012"
        "   <x:Position>2</x:Position>\015\012   <x:ParentField>جمع پورسانت</x:ParentFiel"
        "d>\015\012   <x:NumberFormat>Standard</x:NumberFormat>\015\012  </x:PivotField>\015"
        "\012  <x:PivotData>\015\012   <x:Top>0.0</x:Top>\015\012   <x:TopOffset>0</x:Top"
        "Offset>\015\012   <x:Left>0</x:Left>\015\012   <x:LeftOffset>0</x:LeftOffset>\015"
        "\012   <x:SeqNum>32</x:SeqNum>\015\012  </x:PivotData>\015\012  <x:PivotView>\015"
        "\012   <x:IsNotFiltered/>\015\012   <x:Label>\015\012    <x:Caption>Microsoft Of"
        "fice PivotTable 10.0</x:Caption>\015\012    <x:NotVisible/>\015\012   </x:Label>"
        "\015\012   <x:AllowEdits>false</x:AllowEdits>\015\012   <x:AllowAdditions>false<"
        "/x:AllowAdditions>\015\012   <x:AllowDeletions>false</x:AllowDeletions>\015\012 "
        " </x:PivotView>\015\012 </x:PivotTable>\015\012</xml>"
    PivotChartXML ="<xml xmlns:x=\"urn:schemas-microsoft-com:office:excel\">\015\012 <x:ChartSpace>\015"
        "\012  <x:OWCVersion>10.0.0.5605         </x:OWCVersion>\015\012  <x:Width>26988<"
        "/x:Width>\015\012  <x:Height>15161</x:Height>\015\012  <x:DataSource>\015\012   "
        "<x:Type>PivotList</x:Type>\015\012  </x:DataSource>\015\012  <x:BoundSeries>\015"
        "\012   <x:DataSourceIndex>0</x:DataSourceIndex>\015\012  </x:BoundSeries>\015\012"
        "  <x:Category>\015\012   <x:DataSourceIndex>0</x:DataSourceIndex>\015\012  </x:C"
        "ategory>\015\012  <x:Value>\015\012   <x:DataSourceIndex>0</x:DataSourceIndex>\015"
        "\012  </x:Value>\015\012  <x:BoundCharts>\015\012   <x:DataSourceIndex>0</x:Data"
        "SourceIndex>\015\012  </x:BoundCharts>\015\012  <x:PivotAggOrientation>Series</x"
        ":PivotAggOrientation>\015\012  <x:HidePivotFields/>\015\012  <x:NoGrouping/>\015"
        "\012  <x:NoFiltering/>\015\012  <x:Palette>\015\012   <x:Entry>#000000</x:Entry>"
        "\015\012   <x:Entry>#000000</x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015\012"
        "   <x:Entry>#000000</x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015\012   <x:"
        "Entry>#000000</x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015\012   <x:Entry>"
        "#000000</x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015\012   <x:Entry>#00000"
        "0</x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015\012   <x:Entry>#000000</x:E"
        "ntry>\015\012   <x:Entry>#000000</x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015"
        "\012   <x:Entry>#000000</x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015\012  "
        " <x:Entry>#8080FF</x:Entry>\015\012   <x:Entry>#802060</x:Entry>\015\012   <x:En"
        "try>#FFFFA0</x:Entry>\015\012   <x:Entry>#A0E0E0</x:Entry>\015\012   <x:Entry>#6"
        "00080</x:Entry>\015\012   <x:Entry>#FF8080</x:Entry>\015\012   <x:Entry>#008080<"
        "/x:Entry>\015\012   <x:Entry>#C0C0FF</x:Entry>\015\012   <x:Entry>#000080</x:Ent"
        "ry>\015\012   <x:Entry>#FF00FF</x:Entry>\015\012   <x:Entry>#80FFFF</x:Entry>\015"
        "\012   <x:Entry>#0080FF</x:Entry>\015\012   <x:Entry>#FF8080</x:Entry>\015\012  "
        " <x:Entry>#C0FF80</x:Entry>\015\012   <x:Entry>#FFC0FF</x:Entry>\015\012   <x:En"
        "try>#FF80FF</x:Entry>\015\012  </x:Palette>\015\012  <x:DefaultFont>Arial</x:Def"
        "aultFont>\015\012  <x:Chart>\015\012   <x:PlotArea>\015\012    <x:Graph>\015\012"
        "     <x:SubType>Clustered</x:SubType>\015\012     <x:Type>Column</x:Type>\015\012"
        "     <x:Series>\015\012      <x:FormatMap>\015\012      </x:FormatMap>\015\012  "
        "    <x:Name>Sum of مبلغ كل فروش</x:Name>\015\012      <x:Caption>\015\012       "
        "<x:DataSourceIndex>-1</x:DataSourceIndex>\015\012       <x:Data>&quot;Sum of مبل"
        "غ كل فروش&quot;</x:Data>\015\012      </x:Caption>\015\012      <x:Index>0</x:In"
        "dex>\015\012      <x:Category>\015\012       <x:DataSourceIndex>0</x:DataSourceI"
        "ndex>\015\012      </x:Category>\015\012      <x:Value>\015\012       <x:DataSou"
        "rceIndex>0</x:DataSourceIndex>\015\012      </x:Value>\015\012      <x:DataLabel"
        "s>\015\012       <x:Border>\015\012        <x:ColorIndex>None</x:ColorIndex>\015"
        "\012       </x:Border>\015\012       <x:Number>\015\012        <x:FormatString>G"
        "eneral</x:FormatString>\015\012       </x:Number>\015\012       <x:ShowValue/>\015"
        "\012       <x:Separator>;</x:Separator>\015\012       <x:Position>OutsideEnd</x:"
        "Position>\015\012      </x:DataLabels>\015\012      <x:Marker>\015\012       <x:"
        "Symbol>None</x:Symbol>\015\012      </x:Marker>\015\012      <x:Explode>0</x:Exp"
        "lode>\015\012      <x:Thickness>10</x:Thickness>\015\012      <x:DataSourceIndex"
        ">0</x:DataSourceIndex>\015\012      <x:Identifier>!.Sum of مبلغ كل فروش</x:Ident"
        "ifier>\015\012     </x:Series>\015\012     <x:Series>\015\012      <x:FormatMap>"
        "\015\012      </x:FormatMap>\015\012      <x:Name>Sum of جمع پورسانت</x:Name>\015"
        "\012      <x:Caption>\015\012       <x:DataSourceIndex>-1</x:DataSourceIndex>\015"
        "\012       <x:Data>&quot;Sum of جمع پورسانت&quot;</x:Data>\015\012      </x:Capt"
        "ion>\015\012      <x:Index>1</x:Index>\015\012      <x:Category>\015\012       <"
        "x:DataSourceIndex>0</x:DataSourceIndex>\015\012      </x:Category>\015\012      "
        "<x:Value>\015\012       <x:DataSourceIndex>0</x:DataSourceIndex>\015\012      </"
        "x:Value>\015\012      <x:DataLabels>\015\012       <x:Border>\015\012        <x:"
        "ColorIndex>None</x:ColorIndex>\015\012       </x:Border>\015\012       <x:Number"
        ">\015\012        <x:FormatString>General</x:FormatString>\015\012       </x:Numb"
        "er>\015\012       <x:ShowValue/>\015\012       <x:Separator>;</x:Separator>\015\012"
        "       <x:Position>OutsideEnd</x:Position>\015\012      </x:DataLabels>\015\012 "
        "     <x:Marker>\015\012       <x:Symbol>None</x:Symbol>\015\012      </x:Marker>"
        "\015\012      <x:Explode>0</x:Explode>\015\012      <x:Thickness>10</x:Thickness"
        ">\015\012      <x:DataSourceIndex>0</x:DataSourceIndex>\015\012      <x:Identifi"
        "er>!.Sum of جمع پورسانت</x:Identifier>\015\012     </x:Series>\015\012     <x:Di"
        "mension>\015\012      <x:ScaleID>497960652</x:ScaleID>\015\012      <x:Index>Cat"
        "egories</x:Index>\015\012     </x:Dimension>\015\012     <x:Dimension>\015\012  "
        "    <x:ScaleID>497960872</x:ScaleID>\015\012      <x:Index>Value</x:Index>\015\012"
        "     </x:Dimension>\015\012     <x:Dimension>\015\012      <x:ScaleID>497961092<"
        "/x:ScaleID>\015\012      <x:Index>FormatValue</x:Index>\015\012     </x:Dimensio"
        "n>\015\012     <x:Overlap>0</x:Overlap>\015\012     <x:GapWidth>150</x:GapWidth>"
        "\015\012     <x:FirstSliceAngle>0</x:FirstSliceAngle>\015\012    </x:Graph>\015\012"
        "    <x:Axis>\015\012     <x:AxisID>497975508</x:AxisID>\015\012     <x:ScaleID>4"
        "97960652</x:ScaleID>\015\012     <x:Type>Category</x:Type>\015\012     <x:Alignm"
        "ent>\015\012      <x:Rotation>-90</x:Rotation>\015\012     </x:Alignment>\015\012"
        "     <x:MajorTick>Outside</x:MajorTick>\015\012     <x:MinorTick>None</x:MinorTi"
        "ck>\015\012     <x:Title>\015\012      <x:Font>\015\012       <x:FontName>Arial<"
        "/x:FontName>\015\012       <x:Size>10</x:Size>\015\012      </x:Font>\015\012   "
        "   <x:Border>\015\012       <x:ColorIndex>None</x:ColorIndex>\015\012      </x:B"
        "order>\015\012      <x:Caption>\015\012       <x:DataSourceIndex>-1</x:DataSourc"
        "eIndex>\015\012       <x:Data>&quot;ويزيتور&quot;</x:Data>\015\012      </x:Capt"
        "ion>\015\012      <x:Position>Bottom</x:Position>\015\012     </x:Title>\015\012"
        "     <x:Placement>Bottom</x:Placement>\015\012     <x:GroupingEnum>Auto</x:Group"
        "ingEnum>\015\012    </x:Axis>\015\012    <x:Axis>\015\012     <x:AxisID>49797624"
        "8</x:AxisID>\015\012     <x:ScaleID>497960872</x:ScaleID>\015\012     <x:Type>Va"
        "lue</x:Type>\015\012     <x:Number>\015\012      <x:FormatString>General</x:Form"
        "atString>\015\012     </x:Number>\015\012     <x:MajorGridlines>\015\012      <x"
        ":Line>\015\012       <x:LineJoin>JoinRound</x:LineJoin>\015\012      </x:Line>\015"
        "\012     </x:MajorGridlines>\015\012     <x:MajorTick>Outside</x:MajorTick>\015\012"
        "     <x:MinorTick>None</x:MinorTick>\015\012     <x:Title>\015\012      <x:Font>"
        "\015\012       <x:FontName>Arial</x:FontName>\015\012       <x:Size>10</x:Size>\015"
        "\012      </x:Font>\015\012      <x:Border>\015\012       <x:ColorIndex>None</x:"
        "ColorIndex>\015\012      </x:Border>\015\012      <x:Caption>\015\012       <x:D"
        "ataSourceIndex>-1</x:DataSourceIndex>\015\012       <x:Data>&quot;مبلغ&quot;</x:"
        "Data>\015\012      </x:Caption>\015\012      <x:Position>Left</x:Position>\015\012"
        "     </x:Title>\015\012     <x:Placement>Left</x:Placement>\015\012    </x:Axis>"
        "\015\012   </x:PlotArea>\015\012   <x:Identifier></x:Identifier>\015\012  </x:Ch"
        "art>\015\012  <x:Scaling>\015\012   <x:ScaleID>497960652</x:ScaleID>\015\012  </"
        "x:Scaling>\015\012  <x:Scaling>\015\012   <x:ScaleID>497960872</x:ScaleID>\015\012"
        "  </x:Scaling>\015\012  <x:Scaling>\015\012   <x:ScaleID>497961092</x:ScaleID>\015"
        "\012  </x:Scaling>\015\012  <x:HidePassiveAlert/>\015\012  <x:HideScreenTips/>\015"
        "\012 </x:ChartSpace>\015\012</xml>"
    AllowFormView =0
    RecordSourceQualifier ="dbo"
    Begin
        Begin Label
            BackStyle =0
            TextFontCharSet =178
            FontName ="Tahoma"
        End
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            OldBorderStyle =0
            TextFontCharSet =178
            Width =1701
            LabelX =-1701
            FontName ="Tahoma"
            AsianLineBreak =255
        End
        Begin Subform
            SpecialEffect =2
            Width =1701
            Height =1701
        End
        Begin Section
            CanGrow = NotDefault
            Height =7029
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =7031
                    Top =113
                    Width =2550
                    ColumnWidth =1290
                    ColumnOrder =0
                    Name ="CUST_NO"
                    ControlSource ="CUST_NO"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5217
                            Top =113
                            Width =780
                            Height =240
                            Name ="Label1"
                            Caption ="كد ويزيتور"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =7031
                    Top =453
                    Width =2556
                    ColumnWidth =1320
                    ColumnOrder =3
                    TabIndex =1
                    BackColor =13434828
                    Name ="mabpur"
                    ControlSource ="mabpur"
                    Format ="#,###"
                    ConditionalFormat = Begin
                        0x0100000066000000010000000000000004000000000000000200000001000000 ,
                        0x00000000ccffcc00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x300000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5217
                            Top =453
                            Width =615
                            Height =240
                            Name ="Label3"
                            Caption ="جمع پورسانت"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =7031
                    Top =793
                    Width =2556
                    ColumnWidth =1500
                    ColumnOrder =6
                    TabIndex =2
                    Name ="MABL_K"
                    ControlSource ="MABL_K"
                    Format ="#,###"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5217
                            Top =793
                            Width =1080
                            Height =240
                            Name ="Label5"
                            Caption ="مبلغ كل فروش"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =7031
                    Top =1133
                    Width =2556
                    ColumnWidth =780
                    ColumnOrder =5
                    TabIndex =3
                    Name ="MEGHk"
                    ControlSource ="MEGHk"
                    Format ="#,###"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5217
                            Top =1133
                            Width =570
                            Height =240
                            Name ="Label7"
                            Caption ="مقدار"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =7031
                    Top =1474
                    Width =2556
                    ColumnWidth =795
                    ColumnOrder =9
                    TabIndex =4
                    Name ="MEGH_MAR"
                    ControlSource ="MEGH_MAR"
                    Format ="#,###"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5217
                            Top =1474
                            Width =915
                            Height =240
                            Name ="Label9"
                            Caption ="مرجوعي"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =7031
                    Top =1814
                    Width =2556
                    ColumnWidth =1320
                    ColumnOrder =7
                    TabIndex =5
                    Name ="TAKHF"
                    ControlSource ="TAKHF"
                    Format ="#,###"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5217
                            Top =1814
                            Width =555
                            Height =240
                            Name ="Label11"
                            Caption ="جمع تخفيف"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =7031
                    Top =2154
                    Width =2556
                    ColumnWidth =1320
                    ColumnOrder =8
                    TabIndex =6
                    Name ="MBAA"
                    ControlSource ="MBAA"
                    Format ="#,###"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5217
                            Top =2154
                            Width =495
                            Height =240
                            Name ="Label13"
                            Caption ="ارزش افزوده"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =7031
                    Top =2494
                    Width =2556
                    ColumnWidth =1140
                    ColumnOrder =10
                    TabIndex =7
                    Name ="MABMAR"
                    ControlSource ="MABMAR"
                    Format ="#,###"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5217
                            Top =2494
                            Width =720
                            Height =240
                            Name ="Label15"
                            Caption ="مبلغ مرجوعي"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =7031
                    Top =2834
                    Width =2556
                    ColumnWidth =1500
                    ColumnOrder =11
                    TabIndex =8
                    BackColor =10092543
                    Name ="GHABEL"
                    ControlSource ="GHABEL"
                    Format ="#,###"
                    ConditionalFormat = Begin
                        0x0100000066000000010000000000000004000000000000000200000001000000 ,
                        0x00000000ffff9900000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x300000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4482
                            Top =2834
                            Width =1380
                            Height =240
                            Name ="Label17"
                            Caption ="فروش خالص"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =6974
                    Top =3628
                    Width =2490
                    ColumnWidth =4575
                    ColumnOrder =1
                    TabIndex =9
                    Name ="NAME"
                    ControlSource ="NAME"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5160
                            Top =3628
                            Width =1185
                            Height =240
                            Name ="Label19"
                            Caption ="نام ويزيتور"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =6974
                    Top =3968
                    Width =2490
                    ColumnOrder =14
                    TabIndex =10
                    Name ="ADDRESS"
                    ControlSource ="ADDRESS"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5160
                            Top =3968
                            Width =450
                            Height =240
                            Name ="Label21"
                            Caption ="آدرس"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =6974
                    Top =4309
                    Width =2490
                    ColumnOrder =15
                    TabIndex =11
                    Name ="TEL"
                    ControlSource ="TEL"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5160
                            Top =4309
                            Width =390
                            Height =240
                            Name ="Label23"
                            Caption ="تلفن"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =6974
                    Top =4649
                    Width =2490
                    ColumnOrder =16
                    TabIndex =12
                    Name ="TOZIH"
                    ControlSource ="TOZIH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5160
                            Top =4649
                            Width =540
                            Height =240
                            Name ="Label25"
                            Caption ="توضيح"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =6974
                    Top =4989
                    Width =2490
                    ColumnWidth =1905
                    ColumnOrder =13
                    TabIndex =13
                    Name ="MOBILE"
                    ControlSource ="MOBILE"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =5160
                            Top =4989
                            Width =630
                            Height =240
                            Name ="Label27"
                            Caption ="موبايل"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =7031
                    Top =3194
                    Width =2556
                    ColumnWidth =540
                    ColumnOrder =4
                    TabIndex =14
                    BackColor =16751052
                    Name ="DARSAD"
                    ControlSource ="=IIf(([GHABEL])<>0,([mabpur])/([GHABEL])*100,0)"
                    Format ="Standard"
                    ConditionalFormat = Begin
                        0x0100000066000000010000000000000004000000000000000200000001000000 ,
                        0x00000000cc99ff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x300000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5217
                            Top =3194
                            Width =645
                            Height =240
                            Name ="Label29"
                            Caption ="درصد"
                        End
                    End
                End
                Begin Subform
                    OverlapFlags =247
                    Left =5160
                    Top =5455
                    Width =4757
                    Height =619
                    TabIndex =15
                    Name ="VISITORS_CUST_F"
                    SourceObject ="Form.VISITORS_CUST_F"
                    LinkChildFields ="CUST_NO"
                    LinkMasterFields ="CUST_NO"
                    Begin
                        Begin Label
                            OverlapFlags =255
                            Left =5762
                            Top =5215
                            Width =1260
                            Height =240
                            Name ="VISITORS_CUST Label"
                            Caption ="VISITORS_CUST"
                            EventProcPrefix ="VISITORS_CUST_Label"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =6974
                    Top =6009
                    Width =2556
                    ColumnWidth =960
                    ColumnOrder =2
                    TabIndex =16
                    Name ="tedad"
                    ControlSource ="tedad"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =5160
                            Top =6015
                            Width =855
                            Height =240
                            Name ="Label35"
                            Caption ="تعداد فاكتور"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1814
                    Width =2556
                    ColumnWidth =1500
                    ColumnOrder =12
                    TabIndex =17
                    BackColor =10092543
                    Name ="Text36"
                    ControlSource ="=[GHABEL]-[MBAA]"
                    Format ="Standard"
                    ConditionalFormat = Begin
                        0x0100000066000000010000000000000004000000000000000200000001000000 ,
                        0x00000000ffff9900000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x300000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Width =1005
                            Height =240
                            Name ="Label37"
                            Caption ="مبلغ با تخفيف"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "VISITORS_KOL.cls"
