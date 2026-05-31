Version =20
VersionRequired =20
Begin Form
    AutoCenter = NotDefault
    AllowDeletions = NotDefault
    AllowAdditions = NotDefault
    AllowEdits = NotDefault
    DefaultView =2
    AllowUpdating =4
    ViewsAllowed =2
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =5612
    RowHeight =330
    ItemSuffix =44
    Left =-3270
    Top =630
    Right =15120
    Bottom =8100
    DatasheetGridlinesColor =12632256
    MaxRecords =100000000
    Tag ="@Forms___F_MENU_ANBAR___DT2 BIGINT =[Forms]![F_MENU_ANBAR]![DT2],@Forms___F_MENU"
        "_ANBAR___MANBAR nvarchar(10)=[FormS]![F_MENU_ANBAR]![MANBAR]"
    RecSrcDt = Begin
        0x0600c8be5a7de540
    End
    RecordSource =" SELECT     CODE, MAND, ANBAR, ANBARN,  NAME, NAMES, VCOD, GRCOD, GRNAME, MANDF,"
        " N_FANI, NESBAT, MEGHBAR, bsef, nsef, maxm, minm, VAZN, VAZNK,menuit,MABL_F,B_SE"
        "F,fisiclymand,MAX_M,MEGHRES FROM         dbo.AKMOGUDI_KOL_ANBAR(99999999,8) AKMO"
        "GUDI_KOL_ANBAR where  NAME LIKE '%'  GROUP BY CODE, MAND, ANBAR, ANBARN, FII, MA"
        "BLK, NAME, NAMES, VCOD, GRCOD, GRNAME, MANDF, N_FANI, NESBAT, MEGHBAR, bsef, nse"
        "f, maxm, minm, VAZN, VAZNK ,menuit,MABL_F,B_SEF,fisiclymand,MAX_M,MEGHRes"
    Caption ="موجودي  كالا ها به تفكيك انبار"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Tahoma"
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
        "ivotField>\015\012   <x:Name>في خرده فروش</x:Name>\015\012   <x:SourceName>B_SEF"
        "</x:SourceName>\015\012   <x:DataType>Number</x:DataType>\015\012   <x:EncodedTy"
        "pe>adDouble</x:EncodedType>\015\012   <x:NumberFormat>#,##0.00&quot; ريال&quot;;"
        "-#,##0.00&quot; ريال&quot;</x:NumberFormat>\015\012   <x:CompareOrderedMembersBy"
        ">UniqueName</x:CompareOrderedMembersBy>\015\012  </x:PivotField>\015\012  <x:Piv"
        "otField>\015\012   <x:Name>في عمده فروش</x:Name>\015\012   <x:SourceName>MABL_F<"
        "/x:SourceName>\015\012   <x:DataType>Number</x:DataType>\015\012   <x:EncodedTyp"
        "e>adDouble</x:EncodedType>\015\012   <x:NumberFormat>#,##0.00&quot; ريال&quot;;-"
        "#,##0.00&quot; ريال&quot;</x:NumberFormat>\015\012   <x:CompareOrderedMembersBy>"
        "UniqueName</x:CompareOrderedMembersBy>\015\012  </x:PivotField>\015\012  <x:Pivo"
        "tField>\015\012   <x:Name>فاصله تا بهينه سفارش</x:Name>\015\012   <x:SourceName>"
        "bsef</x:SourceName>\015\012   <x:DataType>Number</x:DataType>\015\012   <x:Encod"
        "edType>adDouble</x:EncodedType>\015\012   <x:CompareOrderedMembersBy>UniqueName<"
        "/x:CompareOrderedMembersBy>\015\012  </x:PivotField>\015\012  <x:PivotField>\015"
        "\012   <x:Name>فاصله تا نقطه سفارش</x:Name>\015\012   <x:SourceName>nsef</x:Sour"
        "ceName>\015\012   <x:DataType>Number</x:DataType>\015\012   <x:EncodedType>adDou"
        "ble</x:EncodedType>\015\012   <x:CompareOrderedMembersBy>UniqueName</x:CompareOr"
        "deredMembersBy>\015\012  </x:PivotField>\015\012  <x:PivotField>\015\012   <x:Na"
        "me>فاصله تا حداقل موجودي</x:Name>\015\012   <x:SourceName>minm</x:SourceName>\015"
        "\012   <x:DataType>Number</x:DataType>\015\012   <x:EncodedType>adDouble</x:Enco"
        "dedType>\015\012   <x:CompareOrderedMembersBy>UniqueName</x:CompareOrderedMember"
        "sBy>\015\012  </x:PivotField>\015\012  <x:PivotField>\015\012   <x:Name>maxm</x:"
        "Name>\015\012   <x:DataType>Number</x:DataType>\015\012   <x:EncodedType>adDoubl"
        "e</x:EncodedType>\015\012   <x:CompareOrderedMembersBy>UniqueName</x:CompareOrde"
        "redMembersBy>\015\012   <x:NoDisplayInFieldList/>\015\012  </x:PivotField>\015\012"
        "  <x:PivotField>\015\012   <x:Name>قيمت مصرف كننده</x:Name>\015\012   <x:SourceN"
        "ame>MAX_M</x:SourceName>\015\012   <x:DataType>Number</x:DataType>\015\012   <x:"
        "EncodedType>adDouble</x:EncodedType>\015\012   <x:CompareOrderedMembersBy>Unique"
        "Name</x:CompareOrderedMembersBy>\015\012  </x:PivotField>\015\012  <x:PivotField"
        ">\015\012   <x:Name>VAZN</x:Name>\015\012   <x:DataType>Number</x:DataType>\015\012"
        "   <x:EncodedType>adDouble</x:EncodedType>\015\012   <x:CompareOrderedMembersBy>"
        "UniqueName</x:CompareOrderedMembersBy>\015\012   <x:NoDisplayInFieldList/>\015\012"
        "  </x:PivotField>\015\012  <x:PivotField>\015\012   <x:Name>VAZNK</x:Name>\015\012"
        "   <x:DataType>Number</x:DataType>\015\012   <x:EncodedType>adDouble</x:EncodedT"
        "ype>\015\012   <x:CompareOrderedMembersBy>UniqueName</x:CompareOrderedMembersBy>"
        "\015\012   <x:NoDisplayInFieldList/>\015\012  </x:PivotField>\015\012  <x:PivotF"
        "ield>\015\012   <x:Name>منوي  كالا</x:Name>\015\012   <x:SourceName>menuit</x:So"
        "urceName>\015\012   <x:EncodedType>adVarWChar</x:EncodedType>\015\012   <x:Compa"
        "reOrderedMembersBy>UniqueName</x:CompareOrderedMembersBy>\015\012  </x:PivotFiel"
        "d>\015\012  <x:PivotField>\015\012   <x:Name>Data</x:Name>\015\012   <x:Orientat"
        "ion>Column</x:Orientation>\015\012   <x:Position>-1</x:Position>\015\012   <x:Da"
        "taField/>\015\012  </x:PivotField>\015\012  <x:PivotField>\015\012   <x:Name>Sum"
        " of موجودي</x:Name>\015\012   <x:PLName>Total1</x:PLName>\015\012   <x:TotalNumb"
        "er>0</x:TotalNumber>\015\012   <x:Orientation>Data</x:Orientation>\015\012   <x:"
        "Position>1</x:Position>\015\012   <x:ParentField>موجودي</x:ParentField>\015\012 "
        " </x:PivotField>\015\012  <x:PivotData>\015\012   <x:Top>0.0</x:Top>\015\012   <"
        "x:TopOffset>0</x:TopOffset>\015\012   <x:Left>0</x:Left>\015\012   <x:LeftOffset"
        ">0</x:LeftOffset>\015\012   <x:SeqNum>17</x:SeqNum>\015\012  </x:PivotData>\015\012"
        "  <x:PivotView>\015\012   <x:IsNotFiltered/>\015\012   <x:Label>\015\012    <x:C"
        "aption>Microsoft Office PivotTable 10.0</x:Caption>\015\012    <x:NotVisible/>\015"
        "\012   </x:Label>\015\012   <x:AllowEdits>false</x:AllowEdits>\015\012   <x:Allo"
        "wAdditions>false</x:AllowAdditions>\015\012   <x:AllowDeletions>false</x:AllowDe"
        "letions>\015\012  </x:PivotView>\015\012 </x:PivotTable>\015\012</xml>"
    PivotChartXML ="<xml xmlns:x=\"urn:schemas-microsoft-com:office:excel\">\015\012 <x:ChartSpace>\015"
        "\012  <x:OWCVersion>10.0.0.5605         </x:OWCVersion>\015\012  <x:Width>26988<"
        "/x:Width>\015\012  <x:Height>17648</x:Height>\015\012  <x:AllowPropertyBrowser/>"
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
        "ID>414851196</x:ScaleID>\015\012      <x:Index>Categories</x:Index>\015\012     "
        "</x:Dimension>\015\012     <x:Dimension>\015\012      <x:ScaleID>414851400</x:Sc"
        "aleID>\015\012      <x:Index>Value</x:Index>\015\012     </x:Dimension>\015\012 "
        "    <x:Dimension>\015\012      <x:ScaleID>414851604</x:ScaleID>\015\012      <x:"
        "Index>FormatValue</x:Index>\015\012     </x:Dimension>\015\012     <x:Overlap>0<"
        "/x:Overlap>\015\012     <x:GapWidth>150</x:GapWidth>\015\012     <x:FirstSliceAn"
        "gle>0</x:FirstSliceAngle>\015\012    </x:Graph>\015\012    <x:Axis>\015\012     "
        "<x:AxisID>414854792</x:AxisID>\015\012     <x:ScaleID>414851196</x:ScaleID>\015\012"
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
        ">\015\012    <x:Axis>\015\012     <x:AxisID>414856188</x:AxisID>\015\012     <x:"
        "ScaleID>414851400</x:ScaleID>\015\012     <x:Type>Value</x:Type>\015\012     <x:"
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
        "  </x:Title>\015\012  <x:Scaling>\015\012   <x:ScaleID>414851196</x:ScaleID>\015"
        "\012  </x:Scaling>\015\012  <x:Scaling>\015\012   <x:ScaleID>414851400</x:ScaleI"
        "D>\015\012  </x:Scaling>\015\012  <x:Scaling>\015\012   <x:ScaleID>414851604</x:"
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
            Height =12075
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =960
                    Top =56
                    Width =2490
                    ColumnWidth =615
                    ColumnOrder =2
                    Name ="CODE"
                    ControlSource ="CODE"
                    OnDblClick ="[Event Procedure]"
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
                    DecimalPlaces =0
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1749
                    Top =396
                    ColumnWidth =780
                    ColumnOrder =6
                    TabIndex =1
                    Name ="MAND"
                    ControlSource ="MAND"
                    Format ="Fixed"
                    ConditionalFormat = Begin
                        0x010000006e000000010000000000000003000000000000000600000001000000 ,
                        0x00000000ccffcc00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x2d0030002e003100310000000000
                    End
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
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Top =2097
                    Width =3450
                    ColumnWidth =5775
                    ColumnOrder =4
                    TabIndex =2
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
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =660
                    Top =2437
                    Width =2790
                    ColumnWidth =1080
                    ColumnOrder =5
                    TabIndex =3
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
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1700
                    Top =737
                    ColumnWidth =690
                    ColumnOrder =1
                    TabIndex =4
                    Name ="ANBAR"
                    ControlSource ="ANBAR"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4457
                            Top =737
                            Width =645
                            Height =240
                            Name ="Label18"
                            Caption ="كد انبار"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1746
                    Top =1098
                    ColumnWidth =1455
                    ColumnOrder =0
                    TabIndex =5
                    Name ="ANBARN"
                    ControlSource ="ANBARN"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4503
                            Top =1098
                            Width =645
                            Height =240
                            Name ="Label20"
                            Caption ="انبار"
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
                    Left =876
                    Top =2777
                    Width =2490
                    ColumnWidth =1417
                    ColumnOrder =16
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
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1417
                    Top =3458
                    ColumnWidth =960
                    ColumnOrder =17
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
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1417
                    Top =3798
                    ColumnWidth =2400
                    ColumnOrder =18
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
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
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
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =396
                    Top =4648
                    ColumnWidth =1890
                    ColumnOrder =3
                    TabIndex =10
                    Name ="N_FANI"
                    ControlSource ="N_FANI"
                    OnDblClick ="[Event Procedure]"
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
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
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
                    DecimalPlaces =2
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =850
                    Top =6526
                    ColumnWidth =1125
                    ColumnOrder =10
                    TabIndex =12
                    Name ="MEGHBAR"
                    ControlSource ="MEGHBAR"
                    Format ="Fixed"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3392
                            Top =6519
                            Width =1020
                            Height =240
                            Name ="Label29"
                            Caption ="بارگيري نشده"
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
                    Left =1644
                    Top =6009
                    Width =1356
                    Height =315
                    ColumnWidth =1485
                    ColumnOrder =14
                    TabIndex =13
                    Name ="B_SEF"
                    ControlSource ="B_SEF"
                    Format ="#,##0.00\" ريال\";-#,##0.00\" ريال\""
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =3162
                            Top =6009
                            Width =1560
                            Height =285
                            Name ="Label30"
                            Caption ="في خرده فروش"
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
                    Left =1644
                    Top =5499
                    Width =1356
                    Height =315
                    ColumnWidth =1485
                    ColumnOrder =15
                    TabIndex =14
                    Name ="MABL_F"
                    ControlSource ="MABL_F"
                    Format ="#,##0.00\" ريال\";-#,##0.00\" ريال\""
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =3163
                            Top =5499
                            Width =1710
                            Height =300
                            Name ="Label31"
                            Caption ="في عمده فروش"
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
                    Left =660
                    Top =3120
                    Width =2790
                    ColumnWidth =3210
                    ColumnOrder =19
                    TabIndex =15
                    Name ="menuit"
                    ControlSource ="menuit"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4695
                            Top =3120
                            Width =765
                            Height =240
                            Name ="Label36"
                            Caption ="منوي  كالا"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =2
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1133
                    Top =7089
                    ColumnWidth =1305
                    ColumnOrder =11
                    TabIndex =16
                    Name ="fisiclymand"
                    ControlSource ="fisiclymand"
                    Format ="Fixed"
                    ConditionalFormat = Begin
                        0x010000006e000000010000000000000003000000000000000600000001000000 ,
                        0x00000000ffcc9900000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x2d0030002e003100310000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4185
                            Top =7080
                            Width =1200
                            Height =240
                            Name ="Label40"
                            Caption ="موجودي فيزيکي"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =2
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =15
                    Top =7494
                    ColumnWidth =810
                    ColumnOrder =9
                    TabIndex =17
                    Name ="MEGHRES"
                    ControlSource ="MEGHRES"
                    Format ="Fixed"
                    ConditionalFormat = Begin
                        0x010000006e000000010000000000000003000000000000000600000001000000 ,
                        0x00000000ff000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x2d0030002e003100310000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3067
                            Top =7485
                            Width =1200
                            Height =240
                            Name ="Label42"
                            Caption ="رزرو شده"
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
                    Top =411
                    Width =1296
                    Height =315
                    ColumnOrder =13
                    FontWeight =700
                    TabIndex =18
                    Name ="VAZNK"
                    ControlSource ="VAZNK"
                    FontName ="Arabic Transparent"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextFontCharSet =0
                            Left =1455
                            Top =405
                            Width =1710
                            Height =240
                            Name ="Label32"
                            Caption ="وزن كل"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =247
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1
                    Top =2
                    Width =1296
                    Height =315
                    ColumnOrder =12
                    FontWeight =700
                    TabIndex =19
                    Name ="VAZN"
                    ControlSource ="VAZN"
                    FontName ="Arabic Transparent"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextFontCharSet =0
                            Left =1455
                            Width =1770
                            Height =240
                            Name ="Label43"
                            Caption ="وزن واحد"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "mogudi_search.cls"
