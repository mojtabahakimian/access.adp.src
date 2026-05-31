Version =20
VersionRequired =20
Begin Form
    DefaultView =2
    AllowUpdating =4
    ViewsAllowed =2
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =5612
    RowHeight =330
    ItemSuffix =35
    Left =2190
    Top =1500
    Right =10290
    Bottom =5880
    DatasheetGridlinesColor =12632256
    MaxRecords =100000000
    Tag ="@Forms___F_MENU_ANBAR___DT2 BIGINT =[Forms]![F_MENU_ANBAR]![DT2],@Forms___F_MENU"
        "_ANBAR___MANBAR nvarchar(10)=[FormS]![F_MENU_ANBAR]![MANBAR]"
    RecSrcDt = Begin
        0x27e78a0d1408e540
    End
    RecordSource ="SELECT CODE, GRCOD, GRNAME, KK, SUM(MABLK) AS MABLK, SUM(MAND) AS MAND, SUM(MAND"
        "F) AS MANDF, N_FANI, NAME, NAMES, NESBAT, TOZIH, VAZN, SUM(VAZNK) AS VAZNK, VCOD"
        ", AVG(FII) AS FII FROM AKMOGUDI_KOL_ANBAR(999999, '%') AKMOGUDI_KOL_ANBAR GROUP "
        "BY CODE, GRCOD, GRNAME, KK, N_FANI, NAME, NAMES, NESBAT, TOZIH, VCOD, VAZN"
    Caption ="موجودي  كالا ها به تفكيك انبارهاي تلفيق شده"
    OnCurrent ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Tahoma"
    OnMouseMove ="[Event Procedure]"
    OnError ="[Event Procedure]"
    PivotTableXML ="<xml xmlns:x=\"urn:schemas-microsoft-com:office:excel\">\015\012 <x:WorksheetOpt"
        "ions>\015\012  <x:DisplayRightToLeft/>\015\012 </x:WorksheetOptions>\015\012 <x:"
        "PivotTable>\015\012  <x:OWCVersion>10.0.0.5605         </x:OWCVersion>\015\012  "
        "<x:NoDisplayAlerts/>\015\012  <x:NoDisplayToolbar/>\015\012  <x:DisplayScreenTip"
        "s/>\015\012  <x:NoAutoFit/>\015\012  <x:Height>384</x:Height>\015\012  <x:Width>"
        "541</x:Width>\015\012  <x:NoAllowGrouping/>\015\012  <x:NoAllowFiltering/>\015\012"
        "  <x:CubeProvider>msolap.2</x:CubeProvider>\015\012  <x:CacheDetails/>\015\012  "
        "<x:Name>Microsoft Office PivotTable 10.0</x:Name>\015\012  <x:PivotField>\015\012"
        "   <x:Name>كد كالا</x:Name>\015\012   <x:SourceName>CODE</x:SourceName>\015\012 "
        "  <x:EncodedType>adVarWChar</x:EncodedType>\015\012   <x:CompareOrderedMembersBy"
        ">UniqueName</x:CompareOrderedMembersBy>\015\012  </x:PivotField>\015\012  <x:Piv"
        "otField>\015\012   <x:Name>موجودي</x:Name>\015\012   <x:SourceName>MAND</x:Sourc"
        "eName>\015\012   <x:DataType>Number</x:DataType>\015\012   <x:PLDataOrientation/"
        ">\015\012   <x:EncodedType>adDouble</x:EncodedType>\015\012   <x:NumberFormat>Fi"
        "xed</x:NumberFormat>\015\012   <x:CompareOrderedMembersBy>UniqueName</x:CompareO"
        "rderedMembersBy>\015\012  </x:PivotField>\015\012  <x:PivotField>\015\012   <x:N"
        "ame>MEGF</x:Name>\015\012   <x:DataType>Number</x:DataType>\015\012   <x:Encoded"
        "Type>adDouble</x:EncodedType>\015\012   <x:CompareOrderedMembersBy>UniqueName</x"
        ":CompareOrderedMembersBy>\015\012   <x:NoDisplayInFieldList/>\015\012  </x:Pivot"
        "Field>\015\012  <x:PivotField>\015\012   <x:Name>SMEGH</x:Name>\015\012   <x:Dat"
        "aType>Number</x:DataType>\015\012   <x:EncodedType>adDouble</x:EncodedType>\015\012"
        "   <x:CompareOrderedMembersBy>UniqueName</x:CompareOrderedMembersBy>\015\012   <"
        "x:NoDisplayInFieldList/>\015\012  </x:PivotField>\015\012  <x:PivotField>\015\012"
        "   <x:Name>في</x:Name>\015\012   <x:SourceName>FII</x:SourceName>\015\012   <x:D"
        "ataType>Number</x:DataType>\015\012   <x:EncodedType>adDouble</x:EncodedType>\015"
        "\012   <x:NumberFormat>Standard</x:NumberFormat>\015\012   <x:CompareOrderedMemb"
        "ersBy>UniqueName</x:CompareOrderedMembersBy>\015\012  </x:PivotField>\015\012  <"
        "x:PivotField>\015\012   <x:Name>مبلغ موجودي</x:Name>\015\012   <x:SourceName>MAB"
        "LK</x:SourceName>\015\012   <x:DataType>Number</x:DataType>\015\012   <x:Encoded"
        "Type>adDouble</x:EncodedType>\015\012   <x:NumberFormat>Standard</x:NumberFormat"
        ">\015\012   <x:CompareOrderedMembersBy>UniqueName</x:CompareOrderedMembersBy>\015"
        "\012  </x:PivotField>\015\012  <x:PivotField>\015\012   <x:Name>نام كالا</x:Name"
        ">\015\012   <x:SourceName>NAME</x:SourceName>\015\012   <x:Orientation>Row</x:Or"
        "ientation>\015\012   <x:EncodedType>adVarWChar</x:EncodedType>\015\012   <x:Comp"
        "areOrderedMembersBy>UniqueName</x:CompareOrderedMembersBy>\015\012   <x:Expanded"
        "/>\015\012  </x:PivotField>\015\012  <x:PivotField>\015\012   <x:Name>نام واحد</"
        "x:Name>\015\012   <x:SourceName>NAMES</x:SourceName>\015\012   <x:EncodedType>ad"
        "VarWChar</x:EncodedType>\015\012   <x:CompareOrderedMembersBy>UniqueName</x:Comp"
        "areOrderedMembersBy>\015\012  </x:PivotField>\015\012  <x:PivotField>\015\012   "
        "<x:Name>KK</x:Name>\015\012   <x:EncodedType>adVarChar</x:EncodedType>\015\012  "
        " <x:CompareOrderedMembersBy>UniqueName</x:CompareOrderedMembersBy>\015\012   <x:"
        "NoDisplayInFieldList/>\015\012  </x:PivotField>\015\012  <x:PivotField>\015\012 "
        "  <x:Name>كد انبار</x:Name>\015\012   <x:SourceName>ANBAR</x:SourceName>\015\012"
        "   <x:DataType>Number</x:DataType>\015\012   <x:EncodedType>adDouble</x:EncodedT"
        "ype>\015\012   <x:CompareOrderedMembersBy>UniqueName</x:CompareOrderedMembersBy>"
        "\015\012  </x:PivotField>\015\012  <x:PivotField>\015\012   <x:Name>انبار</x:Nam"
        "e>\015\012   <x:SourceName>ANBARN</x:SourceName>\015\012   <x:EncodedType>adVarW"
        "Char</x:EncodedType>\015\012   <x:CompareOrderedMembersBy>UniqueName</x:CompareO"
        "rderedMembersBy>\015\012  </x:PivotField>\015\012  <x:PivotField>\015\012   <x:N"
        "ame>كد كالا1</x:Name>\015\012   <x:SourceName>VCOD</x:SourceName>\015\012   <x:D"
        "ataType>Number</x:DataType>\015\012   <x:PLCaption>كد كالا</x:PLCaption>\015\012"
        "   <x:FilterCaption>كد كالا</x:FilterCaption>\015\012   <x:EncodedType>adBigInt<"
        "/x:EncodedType>\015\012   <x:CompareOrderedMembersBy>UniqueName</x:CompareOrdere"
        "dMembersBy>\015\012  </x:PivotField>\015\012  <x:PivotField>\015\012   <x:Name>G"
        "HEMAT</x:Name>\015\012   <x:DataType>Number</x:DataType>\015\012   <x:EncodedTyp"
        "e>adDouble</x:EncodedType>\015\012   <x:CompareOrderedMembersBy>UniqueName</x:Co"
        "mpareOrderedMembersBy>\015\012   <x:NoDisplayInFieldList/>\015\012  </x:PivotFie"
        "ld>\015\012  <x:PivotField>\015\012   <x:Name>كد گروه كالا</x:Name>\015\012   <x"
        ":SourceName>GRCOD</x:SourceName>\015\012   <x:DataType>Number</x:DataType>\015\012"
        "   <x:EncodedType>adDouble</x:EncodedType>\015\012   <x:CompareOrderedMembersBy>"
        "UniqueName</x:CompareOrderedMembersBy>\015\012  </x:PivotField>\015\012  <x:Pivo"
        "tField>\015\012   <x:Name>نام گروه كالا</x:Name>\015\012   <x:SourceName>GRNAME<"
        "/x:SourceName>\015\012   <x:EncodedType>adVarWChar</x:EncodedType>\015\012   <x:"
        "CompareOrderedMembersBy>UniqueName</x:CompareOrderedMembersBy>\015\012  </x:Pivo"
        "tField>\015\012  <x:PivotField>\015\012   <x:Name>شماره فني</x:Name>\015\012   <"
        "x:SourceName>N_FANI</x:SourceName>\015\012   <x:EncodedType>adVarWChar</x:Encode"
        "dType>\015\012   <x:CompareOrderedMembersBy>UniqueName</x:CompareOrderedMembersB"
        "y>\015\012  </x:PivotField>\015\012  <x:PivotField>\015\012   <x:Name>TOZIH</x:N"
        "ame>\015\012   <x:EncodedType>adVarWChar</x:EncodedType>\015\012   <x:CompareOrd"
        "eredMembersBy>UniqueName</x:CompareOrderedMembersBy>\015\012   <x:NoDisplayInFie"
        "ldList/>\015\012  </x:PivotField>\015\012  <x:PivotField>\015\012   <x:Name>ضريب"
        "</x:Name>\015\012   <x:SourceName>NESBAT</x:SourceName>\015\012   <x:DataType>Nu"
        "mber</x:DataType>\015\012   <x:EncodedType>adDouble</x:EncodedType>\015\012   <x"
        ":CompareOrderedMembersBy>UniqueName</x:CompareOrderedMembersBy>\015\012  </x:Piv"
        "otField>\015\012  <x:PivotField>\015\012   <x:Name>موجودي ثانويه</x:Name>\015\012"
        "   <x:SourceName>MANDF</x:SourceName>\015\012   <x:DataType>Number</x:DataType>\015"
        "\012   <x:EncodedType>adDouble</x:EncodedType>\015\012   <x:NumberFormat>Fixed</"
        "x:NumberFormat>\015\012   <x:CompareOrderedMembersBy>UniqueName</x:CompareOrdere"
        "dMembersBy>\015\012  </x:PivotField>\015\012  <x:PivotField>\015\012   <x:Name>ب"
        "ارگيري نشده</x:Name>\015\012   <x:SourceName>MEGHBAR</x:SourceName>\015\012   <x"
        ":DataType>Number</x:DataType>\015\012   <x:EncodedType>adDouble</x:EncodedType>\015"
        "\012   <x:NumberFormat>Fixed</x:NumberFormat>\015\012   <x:CompareOrderedMembers"
        "By>UniqueName</x:CompareOrderedMembersBy>\015\012  </x:PivotField>\015\012  <x:P"
        "ivotField>\015\012   <x:Name>B_SEF</x:Name>\015\012   <x:DataType>Number</x:Data"
        "Type>\015\012   <x:EncodedType>adDouble</x:EncodedType>\015\012   <x:CompareOrde"
        "redMembersBy>UniqueName</x:CompareOrderedMembersBy>\015\012   <x:NoDisplayInFiel"
        "dList/>\015\012  </x:PivotField>\015\012  <x:PivotField>\015\012   <x:Name>MABL_"
        "F</x:Name>\015\012   <x:DataType>Number</x:DataType>\015\012   <x:EncodedType>ad"
        "Double</x:EncodedType>\015\012   <x:CompareOrderedMembersBy>UniqueName</x:Compar"
        "eOrderedMembersBy>\015\012   <x:NoDisplayInFieldList/>\015\012  </x:PivotField>\015"
        "\012  <x:PivotField>\015\012   <x:Name>فاصله تا بهينه سفارش</x:Name>\015\012   <"
        "x:SourceName>bsef</x:SourceName>\015\012   <x:DataType>Number</x:DataType>\015\012"
        "   <x:EncodedType>adDouble</x:EncodedType>\015\012   <x:CompareOrderedMembersBy>"
        "UniqueName</x:CompareOrderedMembersBy>\015\012  </x:PivotField>\015\012  <x:Pivo"
        "tField>\015\012   <x:Name>فاصله تا نقطه سفارش</x:Name>\015\012   <x:SourceName>n"
        "sef</x:SourceName>\015\012   <x:DataType>Number</x:DataType>\015\012   <x:Encode"
        "dType>adDouble</x:EncodedType>\015\012   <x:CompareOrderedMembersBy>UniqueName</"
        "x:CompareOrderedMembersBy>\015\012  </x:PivotField>\015\012  <x:PivotField>\015\012"
        "   <x:Name>فاصله تا حداقل موجودي</x:Name>\015\012   <x:SourceName>minm</x:Source"
        "Name>\015\012   <x:DataType>Number</x:DataType>\015\012   <x:EncodedType>adDoubl"
        "e</x:EncodedType>\015\012   <x:CompareOrderedMembersBy>UniqueName</x:CompareOrde"
        "redMembersBy>\015\012  </x:PivotField>\015\012  <x:PivotField>\015\012   <x:Name"
        ">فاصله تا حد اكثر موجودي</x:Name>\015\012   <x:SourceName>maxm</x:SourceName>\015"
        "\012   <x:DataType>Number</x:DataType>\015\012   <x:EncodedType>adDouble</x:Enco"
        "dedType>\015\012   <x:CompareOrderedMembersBy>UniqueName</x:CompareOrderedMember"
        "sBy>\015\012  </x:PivotField>\015\012  <x:PivotField>\015\012   <x:Name>MAX_M</x"
        ":Name>\015\012   <x:DataType>Number</x:DataType>\015\012   <x:EncodedType>adDoub"
        "le</x:EncodedType>\015\012   <x:CompareOrderedMembersBy>UniqueName</x:CompareOrd"
        "eredMembersBy>\015\012   <x:NoDisplayInFieldList/>\015\012  </x:PivotField>\015\012"
        "  <x:PivotField>\015\012   <x:Name>وزن واحد</x:Name>\015\012   <x:SourceName>VAZ"
        "N</x:SourceName>\015\012   <x:DataType>Number</x:DataType>\015\012   <x:EncodedT"
        "ype>adDouble</x:EncodedType>\015\012   <x:CompareOrderedMembersBy>UniqueName</x:"
        "CompareOrderedMembersBy>\015\012  </x:PivotField>\015\012  <x:PivotField>\015\012"
        "   <x:Name>وزن كل</x:Name>\015\012   <x:SourceName>VAZNK</x:SourceName>\015\012 "
        "  <x:DataType>Number</x:DataType>\015\012   <x:EncodedType>adDouble</x:EncodedTy"
        "pe>\015\012   <x:CompareOrderedMembersBy>UniqueName</x:CompareOrderedMembersBy>\015"
        "\012  </x:PivotField>\015\012  <x:PivotField>\015\012   <x:Name>Data</x:Name>\015"
        "\012   <x:Orientation>Column</x:Orientation>\015\012   <x:Position>-1</x:Positio"
        "n>\015\012   <x:DataField/>\015\012  </x:PivotField>\015\012  <x:PivotField>\015"
        "\012   <x:Name>Sum of موجودي</x:Name>\015\012   <x:PLName>Total1</x:PLName>\015\012"
        "   <x:TotalNumber>0</x:TotalNumber>\015\012   <x:Orientation>Data</x:Orientation"
        ">\015\012   <x:Position>1</x:Position>\015\012   <x:ParentField>موجودي</x:Parent"
        "Field>\015\012  </x:PivotField>\015\012  <x:PivotData>\015\012   <x:Top>0.0</x:T"
        "op>\015\012   <x:TopOffset>0</x:TopOffset>\015\012   <x:Left>0</x:Left>\015\012 "
        "  <x:LeftOffset>0</x:LeftOffset>\015\012   <x:SeqNum>17</x:SeqNum>\015\012  </x:"
        "PivotData>\015\012  <x:PivotView>\015\012   <x:IsNotFiltered/>\015\012   <x:Labe"
        "l>\015\012    <x:Caption>Microsoft Office PivotTable 10.0</x:Caption>\015\012   "
        " <x:NotVisible/>\015\012   </x:Label>\015\012   <x:AllowEdits>false</x:AllowEdit"
        "s>\015\012   <x:AllowAdditions>false</x:AllowAdditions>\015\012   <x:AllowDeleti"
        "ons>false</x:AllowDeletions>\015\012  </x:PivotView>\015\012 </x:PivotTable>\015"
        "\012</xml>"
    PivotChartXML ="<xml xmlns:x=\"urn:schemas-microsoft-com:office:excel\">\015\012 <x:ChartSpace>\015"
        "\012  <x:OWCVersion>10.0.0.5605         </x:OWCVersion>\015\012  <x:Width>33761<"
        "/x:Width>\015\012  <x:Height>18494</x:Height>\015\012  <x:AllowPropertyBrowser/>"
        "\015\012  <x:DataSource>\015\012   <x:Type>PivotList</x:Type>\015\012  </x:DataS"
        "ource>\015\012  <x:BoundSeries>\015\012   <x:DataSourceIndex>0</x:DataSourceInde"
        "x>\015\012  </x:BoundSeries>\015\012  <x:Category>\015\012   <x:DataSourceIndex>"
        "0</x:DataSourceIndex>\015\012  </x:Category>\015\012  <x:Value>\015\012   <x:Dat"
        "aSourceIndex>0</x:DataSourceIndex>\015\012  </x:Value>\015\012  <x:BoundCharts>\015"
        "\012   <x:DataSourceIndex>0</x:DataSourceIndex>\015\012  </x:BoundCharts>\015\012"
        "  <x:FormatValue>\015\012   <x:DataSourceIndex>-3</x:DataSourceIndex>\015\012   "
        "<x:Data>2</x:Data>\015\012  </x:FormatValue>\015\012  <x:PivotAggOrientation>Ser"
        "ies</x:PivotAggOrientation>\015\012  <x:HidePivotFields/>\015\012  <x:NoGrouping"
        "/>\015\012  <x:NoFiltering/>\015\012  <x:Palette>\015\012   <x:Entry>#000000</x:"
        "Entry>\015\012   <x:Entry>#000000</x:Entry>\015\012   <x:Entry>#000000</x:Entry>"
        "\015\012   <x:Entry>#000000</x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015\012"
        "   <x:Entry>#000000</x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015\012   <x:"
        "Entry>#000000</x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015\012   <x:Entry>"
        "#000000</x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015\012   <x:Entry>#00000"
        "0</x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015\012   <x:Entry>#000000</x:E"
        "ntry>\015\012   <x:Entry>#000000</x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015"
        "\012   <x:Entry>#8080FF</x:Entry>\015\012   <x:Entry>#802060</x:Entry>\015\012  "
        " <x:Entry>#FFFFA0</x:Entry>\015\012   <x:Entry>#A0E0E0</x:Entry>\015\012   <x:En"
        "try>#600080</x:Entry>\015\012   <x:Entry>#FF8080</x:Entry>\015\012   <x:Entry>#0"
        "08080</x:Entry>\015\012   <x:Entry>#C0C0FF</x:Entry>\015\012   <x:Entry>#000080<"
        "/x:Entry>\015\012   <x:Entry>#FF00FF</x:Entry>\015\012   <x:Entry>#80FFFF</x:Ent"
        "ry>\015\012   <x:Entry>#0080FF</x:Entry>\015\012   <x:Entry>#FF8080</x:Entry>\015"
        "\012   <x:Entry>#C0FF80</x:Entry>\015\012   <x:Entry>#FFC0FF</x:Entry>\015\012  "
        " <x:Entry>#FF80FF</x:Entry>\015\012  </x:Palette>\015\012  <x:DefaultFont>Arial<"
        "/x:DefaultFont>\015\012  <x:Border>\015\012   <x:Color>#2F4F4F</x:Color>\015\012"
        "  </x:Border>\015\012  <x:Interior>\015\012   <x:Color>#FFFFFF</x:Color>\015\012"
        "   <x:FillEffect>\015\012    <x:fill x:type=\"Solid\" x:color=\"#FFFFFF\"/>\015\012"
        "   </x:FillEffect>\015\012  </x:Interior>\015\012  <x:Chart>\015\012   <x:PlotAr"
        "ea>\015\012    <x:Graph>\015\012     <x:SubType>Clustered</x:SubType>\015\012   "
        "  <x:Type>Column</x:Type>\015\012     <x:Series>\015\012      <x:FormatMap>\015\012"
        "      </x:FormatMap>\015\012      <x:Name>Series</x:Name>\015\012      <x:Captio"
        "n>\015\012       <x:DataSourceIndex>-1</x:DataSourceIndex>\015\012       <x:Data"
        ">&quot;Series&quot;</x:Data>\015\012      </x:Caption>\015\012      <x:Index>0</"
        "x:Index>\015\012      <x:Category>\015\012       <x:DataSourceIndex>0</x:DataSou"
        "rceIndex>\015\012      </x:Category>\015\012      <x:Value>\015\012       <x:Dat"
        "aSourceIndex>0</x:DataSourceIndex>\015\012      </x:Value>\015\012      <x:Forma"
        "tValue>\015\012       <x:DataSourceIndex>-3</x:DataSourceIndex>\015\012       <x"
        ":Data>2</x:Data>\015\012      </x:FormatValue>\015\012      <x:Marker>\015\012  "
        "     <x:Symbol>None</x:Symbol>\015\012      </x:Marker>\015\012      <x:Explode>"
        "0</x:Explode>\015\012      <x:Thickness>10</x:Thickness>\015\012      <x:DataSou"
        "rceIndex>0</x:DataSourceIndex>\015\012      <x:Identifier>!.Sum of موجودي</x:Ide"
        "ntifier>\015\012     </x:Series>\015\012     <x:Dimension>\015\012      <x:Scale"
        "ID>576131108</x:ScaleID>\015\012      <x:Index>Categories</x:Index>\015\012     "
        "</x:Dimension>\015\012     <x:Dimension>\015\012      <x:ScaleID>576131312</x:Sc"
        "aleID>\015\012      <x:Index>Value</x:Index>\015\012     </x:Dimension>\015\012 "
        "    <x:Dimension>\015\012      <x:ScaleID>576131516</x:ScaleID>\015\012      <x:"
        "Index>FormatValue</x:Index>\015\012     </x:Dimension>\015\012     <x:Overlap>0<"
        "/x:Overlap>\015\012     <x:GapWidth>150</x:GapWidth>\015\012     <x:FirstSliceAn"
        "gle>0</x:FirstSliceAngle>\015\012    </x:Graph>\015\012    <x:Axis>\015\012     "
        "<x:AxisID>576134704</x:AxisID>\015\012     <x:ScaleID>576131108</x:ScaleID>\015\012"
        "     <x:Type>Category</x:Type>\015\012     <x:Font>\015\012      <x:FontName>Tah"
        "oma</x:FontName>\015\012      <x:B>Automatic</x:B>\015\012      <x:I>Automatic</"
        "x:I>\015\012      <x:U>Automatic</x:U>\015\012     </x:Font>\015\012     <x:Majo"
        "rTick>Outside</x:MajorTick>\015\012     <x:MinorTick>None</x:MinorTick>\015\012 "
        "    <x:Title>\015\012      <x:Font>\015\012       <x:FontName>Tahoma</x:FontName"
        ">\015\012       <x:Size>8</x:Size>\015\012       <x:Color>#000000</x:Color>\015\012"
        "       <x:B>Automatic</x:B>\015\012       <x:I>Automatic</x:I>\015\012       <x:"
        "U>Automatic</x:U>\015\012      </x:Font>\015\012      <x:Border>\015\012       <"
        "x:ColorIndex>None</x:ColorIndex>\015\012      </x:Border>\015\012      <x:Captio"
        "n>\015\012       <x:DataSourceIndex>-1</x:DataSourceIndex>\015\012       <x:Data"
        ">&quot;نام كالا&quot;</x:Data>\015\012      </x:Caption>\015\012      <x:Positio"
        "n>Bottom</x:Position>\015\012     </x:Title>\015\012     <x:Placement>Bottom</x:"
        "Placement>\015\012     <x:GroupingEnum>Auto</x:GroupingEnum>\015\012    </x:Axis"
        ">\015\012    <x:Axis>\015\012     <x:AxisID>576136100</x:AxisID>\015\012     <x:"
        "ScaleID>576131312</x:ScaleID>\015\012     <x:Type>Value</x:Type>\015\012     <x:"
        "Number>\015\012      <x:FormatString>General</x:FormatString>\015\012     </x:Nu"
        "mber>\015\012     <x:MajorGridlines>\015\012     </x:MajorGridlines>\015\012    "
        " <x:MajorTick>Outside</x:MajorTick>\015\012     <x:MinorTick>None</x:MinorTick>\015"
        "\012     <x:Title>\015\012      <x:Font>\015\012       <x:FontName>Tahoma</x:Fon"
        "tName>\015\012       <x:Size>8</x:Size>\015\012       <x:B>Automatic</x:B>\015\012"
        "       <x:I>Automatic</x:I>\015\012       <x:U>Automatic</x:U>\015\012      </x:"
        "Font>\015\012      <x:Border>\015\012       <x:ColorIndex>None</x:ColorIndex>\015"
        "\012      </x:Border>\015\012      <x:Caption>\015\012       <x:DataSourceIndex>"
        "-1</x:DataSourceIndex>\015\012       <x:Data>&quot;موجودي&quot;</x:Data>\015\012"
        "      </x:Caption>\015\012      <x:Position>Left</x:Position>\015\012     </x:Ti"
        "tle>\015\012     <x:Placement>Left</x:Placement>\015\012    </x:Axis>\015\012   "
        "</x:PlotArea>\015\012   <x:Identifier></x:Identifier>\015\012  </x:Chart>\015\012"
        "  <x:Title>\015\012   <x:Font>\015\012    <x:FontName>Sina</x:FontName>\015\012 "
        "   <x:Color>#228B22</x:Color>\015\012    <x:B>Automatic</x:B>\015\012    <x:I>Au"
        "tomatic</x:I>\015\012    <x:U>Automatic</x:U>\015\012   </x:Font>\015\012   <x:B"
        "order>\015\012    <x:ColorIndex>None</x:ColorIndex>\015\012   </x:Border>\015\012"
        "   <x:Caption>\015\012    <x:DataSourceIndex>-1</x:DataSourceIndex>\015\012    <"
        "x:Data>&quot;نمودار موجودي كالا ها&quot;</x:Data>\015\012   </x:Caption>\015\012"
        "  </x:Title>\015\012  <x:Scaling>\015\012   <x:ScaleID>576131108</x:ScaleID>\015"
        "\012  </x:Scaling>\015\012  <x:Scaling>\015\012   <x:ScaleID>576131312</x:ScaleI"
        "D>\015\012  </x:Scaling>\015\012  <x:Scaling>\015\012   <x:ScaleID>576131516</x:"
        "ScaleID>\015\012  </x:Scaling>\015\012  <x:HidePassiveAlert/>\015\012  <x:HideSc"
        "reenTips/>\015\012 </x:ChartSpace>\015\012</xml>"
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
        Begin Section
            Height =8259
            BackColor =-2147483633
            Name ="Detail"
            OnMouseMove ="[Event Procedure]"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =960
                    Top =56
                    Width =2490
                    Height =255
                    ColumnWidth =900
                    ColumnOrder =2
                    Name ="CODE"
                    ControlSource ="CODE"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4821
                            Top =63
                            Width =510
                            Height =240
                            Name ="Label1"
                            Caption ="كد كالا"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    DecimalPlaces =2
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1749
                    Top =396
                    ColumnWidth =1500
                    ColumnOrder =6
                    TabIndex =1
                    Name ="MAND"
                    ControlSource ="MAND"
                    Format ="Fixed"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4806
                            Top =393
                            Width =645
                            Height =240
                            Name ="Label3"
                            Caption ="موجودي"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    DecimalPlaces =2
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1749
                    Top =1417
                    ColumnWidth =1680
                    ColumnOrder =12
                    TabIndex =2
                    Name ="FII1"
                    ControlSource ="FII"
                    Format ="Standard"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3971
                            Top =1417
                            Width =1350
                            Height =240
                            Name ="Label9"
                            Caption ="في"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1749
                    Top =1757
                    ColumnWidth =1680
                    ColumnOrder =13
                    TabIndex =3
                    Name ="MABLK"
                    ControlSource ="MABLK"
                    Format ="Standard"
                    OnDblClick ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4041
                            Top =1758
                            Width =1545
                            Height =225
                            Name ="Label11"
                            Caption ="مبلغ موجودي"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Top =2097
                    Width =3450
                    ColumnWidth =4440
                    ColumnOrder =4
                    TabIndex =4
                    Name ="NAME"
                    ControlSource ="NAME"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4821
                            Top =2103
                            Width =540
                            Height =240
                            Name ="Label13"
                            Caption ="نام كالا"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =660
                    Top =2437
                    Width =2790
                    ColumnWidth =735
                    ColumnOrder =5
                    TabIndex =5
                    Name ="NAMES"
                    ControlSource ="NAMES"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4691
                            Top =2437
                            Width =630
                            Height =240
                            Name ="Label15"
                            Caption ="نام واحد"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =876
                    Top =2777
                    Width =2490
                    ColumnWidth =1417
                    ColumnOrder =14
                    TabIndex =6
                    Name ="VCOD"
                    ControlSource ="VCOD"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4737
                            Top =2784
                            Width =510
                            Height =240
                            Name ="Label22"
                            Caption ="كد كالا"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1417
                    Top =3458
                    ColumnWidth =960
                    ColumnOrder =15
                    TabIndex =7
                    Name ="GRCOD"
                    ControlSource ="GRCOD"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3904
                            Top =3458
                            Width =915
                            Height =240
                            Name ="Label23"
                            Caption ="كد گروه كالا"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1417
                    Top =3798
                    ColumnWidth =2400
                    ColumnOrder =16
                    TabIndex =8
                    Name ="GRNAME"
                    ControlSource ="GRNAME"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3874
                            Top =3798
                            Width =945
                            Height =240
                            Name ="Label24"
                            Caption ="نام گروه كالا"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1190
                    Top =4254
                    ColumnWidth =1140
                    ColumnOrder =8
                    TabIndex =9
                    Name ="MANDF"
                    ControlSource ="MANDF"
                    Format ="Fixed"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4245
                            Top =4245
                            Width =1035
                            Height =240
                            Name ="Label26"
                            Caption ="موجودي ثانويه"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =396
                    Top =4648
                    ColumnWidth =1440
                    ColumnOrder =3
                    TabIndex =10
                    Name ="N_FANI"
                    ControlSource ="N_FANI"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3105
                            Top =4650
                            Width =870
                            Height =240
                            Name ="Label27"
                            Caption ="شماره فني"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1076
                    Top =5102
                    ColumnWidth =585
                    ColumnOrder =7
                    TabIndex =11
                    Name ="NESBAT"
                    ControlSource ="NESBAT"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3773
                            Top =5102
                            Width =480
                            Height =240
                            Name ="Label28"
                            Caption ="ضريب"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1700
                    Top =7944
                    Width =1296
                    Height =315
                    ColumnWidth =825
                    ColumnOrder =10
                    FontWeight =700
                    TabIndex =12
                    Name ="VAZNK"
                    ControlSource ="VAZNK"
                    FontName ="Arabic Transparent"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =3155
                            Top =7938
                            Width =1710
                            Height =240
                            Name ="Label32"
                            Caption ="وزن كل"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1701
                    Top =7535
                    Width =1296
                    Height =315
                    ColumnWidth =780
                    ColumnOrder =9
                    FontWeight =700
                    TabIndex =13
                    Name ="VAZN"
                    ControlSource ="VAZN"
                    FontName ="Arabic Transparent"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =3155
                            Top =7533
                            Width =1770
                            Height =240
                            Name ="Label34"
                            Caption ="وزن واحد"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "AK_MOGUDI_ANBAR_LIST_MERG.cls"
