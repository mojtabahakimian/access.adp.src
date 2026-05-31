Version =20
VersionRequired =20
Begin Form
    AllowDeletions = NotDefault
    AllowAdditions = NotDefault
    OrderByOn = NotDefault
    AllowEdits = NotDefault
    DefaultView =4
    AllowUpdating =4
    ViewsAllowed =2
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =5730
    RowHeight =270
    ItemSuffix =5
    Left =720
    Top =1380
    Right =6930
    Bottom =2790
    DatasheetGridlinesColor =12632256
    MaxRecords =0
    RecSrcDt = Begin
        0x0756722c5a2ee540
    End
    RecordSource ="sudmah"
    Caption ="نبض سود(سود ماه)"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Tahoma"
    PivotTableXML ="<xml xmlns:x=\"urn:schemas-microsoft-com:office:excel\">\015\012 <x:WorksheetOpt"
        "ions>\015\012  <x:DisplayRightToLeft/>\015\012 </x:WorksheetOptions>\015\012 <x:"
        "PivotTable>\015\012  <x:OWCVersion>10.0.0.5605         </x:OWCVersion>\015\012  "
        "<x:NoDisplayAlerts/>\015\012  <x:NoDisplayToolbar/>\015\012  <x:DisplayScreenTip"
        "s/>\015\012  <x:NoAutoFit/>\015\012  <x:Height>384</x:Height>\015\012  <x:Width>"
        "541</x:Width>\015\012  <x:NoAllowGrouping/>\015\012  <x:NoAllowFiltering/>\015\012"
        "  <x:CubeProvider>msolap.2</x:CubeProvider>\015\012  <x:CacheDetails/>\015\012  "
        "<x:Name>Microsoft Office PivotTable 10.0</x:Name>\015\012  <x:PivotField>\015\012"
        "   <x:Name>ماه</x:Name>\015\012   <x:SourceName>DATE_S</x:SourceName>\015\012   "
        "<x:DataType>Number</x:DataType>\015\012   <x:Orientation>Row</x:Orientation>\015"
        "\012   <x:EncodedType>adBigInt</x:EncodedType>\015\012   <x:CompareOrderedMember"
        "sBy>UniqueName</x:CompareOrderedMembersBy>\015\012   <x:Expanded/>\015\012  </x:"
        "PivotField>\015\012  <x:PivotField>\015\012   <x:Name>دارئيها</x:Name>\015\012  "
        " <x:SourceName>BEDS</x:SourceName>\015\012   <x:DataType>Number</x:DataType>\015"
        "\012   <x:EncodedType>adDouble</x:EncodedType>\015\012   <x:CompareOrderedMember"
        "sBy>UniqueName</x:CompareOrderedMembersBy>\015\012  </x:PivotField>\015\012  <x:"
        "PivotField>\015\012   <x:Name>بدهيها</x:Name>\015\012   <x:SourceName>BESS</x:So"
        "urceName>\015\012   <x:DataType>Number</x:DataType>\015\012   <x:EncodedType>adD"
        "ouble</x:EncodedType>\015\012   <x:CompareOrderedMembersBy>UniqueName</x:Compare"
        "OrderedMembersBy>\015\012  </x:PivotField>\015\012  <x:PivotField>\015\012   <x:"
        "Name>سود روزانه</x:Name>\015\012   <x:SourceName>sudday</x:SourceName>\015\012  "
        " <x:DataType>Number</x:DataType>\015\012   <x:PLDataOrientation/>\015\012   <x:E"
        "ncodedType>adDouble</x:EncodedType>\015\012   <x:CompareOrderedMembersBy>UniqueN"
        "ame</x:CompareOrderedMembersBy>\015\012  </x:PivotField>\015\012  <x:PivotField>"
        "\015\012   <x:Name>سود</x:Name>\015\012   <x:SourceName>SUD</x:SourceName>\015\012"
        "   <x:DataType>Number</x:DataType>\015\012   <x:EncodedType>adDouble</x:EncodedT"
        "ype>\015\012   <x:CompareOrderedMembersBy>UniqueName</x:CompareOrderedMembersBy>"
        "\015\012  </x:PivotField>\015\012  <x:PivotField>\015\012   <x:Name>Data</x:Name"
        ">\015\012   <x:Orientation>Column</x:Orientation>\015\012   <x:Position>-1</x:Po"
        "sition>\015\012   <x:DataField/>\015\012  </x:PivotField>\015\012  <x:PivotField"
        ">\015\012   <x:Name>Sum of سود روزانه</x:Name>\015\012   <x:PLName>Total1</x:PLN"
        "ame>\015\012   <x:TotalNumber>0</x:TotalNumber>\015\012   <x:Orientation>Data</x"
        ":Orientation>\015\012   <x:Position>1</x:Position>\015\012   <x:ParentField>سود "
        "روزانه</x:ParentField>\015\012  </x:PivotField>\015\012  <x:PLTotal>\015\012   <"
        "x:Name>Total2</x:Name>\015\012   <x:Caption>Sum of سود</x:Caption>\015\012   <x:"
        "ParentField>سود</x:ParentField>\015\012   <x:TotalNumber>1</x:TotalNumber>\015\012"
        "  </x:PLTotal>\015\012  <x:PLTotal>\015\012   <x:Name>Total3</x:Name>\015\012   "
        "<x:Caption>Sum of سود (2)</x:Caption>\015\012   <x:ParentField>سود</x:ParentFiel"
        "d>\015\012   <x:TotalNumber>2</x:TotalNumber>\015\012  </x:PLTotal>\015\012  <x:"
        "PLTotal>\015\012   <x:Name>Total4</x:Name>\015\012   <x:Caption>Sum of سود (3)</"
        "x:Caption>\015\012   <x:ParentField>سود</x:ParentField>\015\012   <x:TotalNumber"
        ">3</x:TotalNumber>\015\012  </x:PLTotal>\015\012  <x:PivotData>\015\012   <x:Top"
        ">0.0</x:Top>\015\012   <x:TopOffset>0</x:TopOffset>\015\012   <x:Left>0</x:Left>"
        "\015\012   <x:LeftOffset>0</x:LeftOffset>\015\012   <x:SeqNum>39</x:SeqNum>\015\012"
        "  </x:PivotData>\015\012  <x:PivotView>\015\012   <x:IsNotFiltered/>\015\012   <"
        "x:Label>\015\012    <x:Caption>Microsoft Office PivotTable 10.0</x:Caption>\015\012"
        "    <x:NotVisible/>\015\012   </x:Label>\015\012   <x:AllowEdits>false</x:AllowE"
        "dits>\015\012   <x:AllowAdditions>false</x:AllowAdditions>\015\012   <x:AllowDel"
        "etions>false</x:AllowDeletions>\015\012  </x:PivotView>\015\012 </x:PivotTable>\015"
        "\012</xml>"
    PivotChartXML ="<xml xmlns:x=\"urn:schemas-microsoft-com:office:excel\">\015\012 <x:ChartSpace>\015"
        "\012  <x:OWCVersion>10.0.0.5605         </x:OWCVersion>\015\012  <x:Width>33761<"
        "/x:Width>\015\012  <x:Height>16960</x:Height>\015\012  <x:AllowPropertyBrowser/>"
        "\015\012  <x:DataSource>\015\012   <x:Type>PivotList</x:Type>\015\012  </x:DataS"
        "ource>\015\012  <x:BoundSeries>\015\012   <x:DataSourceIndex>0</x:DataSourceInde"
        "x>\015\012  </x:BoundSeries>\015\012  <x:Category>\015\012   <x:DataSourceIndex>"
        "0</x:DataSourceIndex>\015\012  </x:Category>\015\012  <x:Value>\015\012   <x:Dat"
        "aSourceIndex>0</x:DataSourceIndex>\015\012   <x:Data>Total1</x:Data>\015\012  </"
        "x:Value>\015\012  <x:BoundCharts>\015\012   <x:DataSourceIndex>0</x:DataSourceIn"
        "dex>\015\012  </x:BoundCharts>\015\012  <x:PivotAggOrientation>None</x:PivotAggO"
        "rientation>\015\012  <x:HidePivotFields/>\015\012  <x:NoGrouping/>\015\012  <x:N"
        "oFiltering/>\015\012  <x:Palette>\015\012   <x:Entry>#000000</x:Entry>\015\012  "
        " <x:Entry>#000000</x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015\012   <x:En"
        "try>#000000</x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015\012   <x:Entry>#0"
        "00000</x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015\012   <x:Entry>#000000<"
        "/x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015\012   <x:Entry>#000000</x:Ent"
        "ry>\015\012   <x:Entry>#000000</x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015"
        "\012   <x:Entry>#000000</x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015\012  "
        " <x:Entry>#000000</x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015\012   <x:En"
        "try>#8080FF</x:Entry>\015\012   <x:Entry>#802060</x:Entry>\015\012   <x:Entry>#F"
        "FFFA0</x:Entry>\015\012   <x:Entry>#A0E0E0</x:Entry>\015\012   <x:Entry>#600080<"
        "/x:Entry>\015\012   <x:Entry>#FF8080</x:Entry>\015\012   <x:Entry>#008080</x:Ent"
        "ry>\015\012   <x:Entry>#C0C0FF</x:Entry>\015\012   <x:Entry>#000080</x:Entry>\015"
        "\012   <x:Entry>#FF00FF</x:Entry>\015\012   <x:Entry>#80FFFF</x:Entry>\015\012  "
        " <x:Entry>#0080FF</x:Entry>\015\012   <x:Entry>#FF8080</x:Entry>\015\012   <x:En"
        "try>#C0FF80</x:Entry>\015\012   <x:Entry>#FFC0FF</x:Entry>\015\012   <x:Entry>#F"
        "F80FF</x:Entry>\015\012  </x:Palette>\015\012  <x:DefaultFont>Arial</x:DefaultFo"
        "nt>\015\012  <x:Chart>\015\012   <x:PlotArea>\015\012    <x:Graph>\015\012     <"
        "x:SubType>Clustered</x:SubType>\015\012     <x:Type>Column</x:Type>\015\012     "
        "<x:Series>\015\012      <x:Interior>\015\012       <x:Color>#8B008B</x:Color>\015"
        "\012       <x:FillEffect>\015\012        <x:fill x:type=\"Solid\" x:color=\"#8B0"
        "08B\"/>\015\012       </x:FillEffect>\015\012      </x:Interior>\015\012      <x"
        ":FormatMap>\015\012      </x:FormatMap>\015\012      <x:Name>Series</x:Name>\015"
        "\012      <x:Caption>\015\012       <x:DataSourceIndex>-1</x:DataSourceIndex>\015"
        "\012       <x:Data>&quot;Series&quot;</x:Data>\015\012      </x:Caption>\015\012"
        "      <x:Index>0</x:Index>\015\012      <x:Category>\015\012       <x:DataSource"
        "Index>0</x:DataSourceIndex>\015\012      </x:Category>\015\012      <x:Value>\015"
        "\012       <x:DataSourceIndex>0</x:DataSourceIndex>\015\012       <x:Data>Total1"
        "</x:Data>\015\012      </x:Value>\015\012      <x:DataLabels>\015\012       <x:B"
        "order>\015\012        <x:ColorIndex>None</x:ColorIndex>\015\012       </x:Border"
        ">\015\012       <x:Font>\015\012        <x:Size>12</x:Size>\015\012        <x:B>"
        "Automatic</x:B>\015\012        <x:I>Automatic</x:I>\015\012        <x:U>Automati"
        "c</x:U>\015\012       </x:Font>\015\012       <x:Number>\015\012        <x:Forma"
        "tString>#,###</x:FormatString>\015\012       </x:Number>\015\012       <x:ShowVa"
        "lue/>\015\012       <x:Separator>;</x:Separator>\015\012       <x:Position>Outsi"
        "deEnd</x:Position>\015\012      </x:DataLabels>\015\012      <x:Marker>\015\012 "
        "      <x:Symbol>None</x:Symbol>\015\012      </x:Marker>\015\012      <x:Explode"
        ">0</x:Explode>\015\012      <x:Thickness>10</x:Thickness>\015\012      <x:DataSo"
        "urceIndex>0</x:DataSourceIndex>\015\012      <x:Identifier>!</x:Identifier>\015\012"
        "     </x:Series>\015\012     <x:Dimension>\015\012      <x:ScaleID>407830976</x:"
        "ScaleID>\015\012      <x:Index>Categories</x:Index>\015\012     </x:Dimension>\015"
        "\012     <x:Dimension>\015\012      <x:ScaleID>407831592</x:ScaleID>\015\012    "
        "  <x:Index>Value</x:Index>\015\012     </x:Dimension>\015\012     <x:Dimension>\015"
        "\012      <x:ScaleID>407831796</x:ScaleID>\015\012      <x:Index>FormatValue</x:"
        "Index>\015\012     </x:Dimension>\015\012     <x:Overlap>0</x:Overlap>\015\012  "
        "   <x:GapWidth>150</x:GapWidth>\015\012     <x:FirstSliceAngle>0</x:FirstSliceAn"
        "gle>\015\012    </x:Graph>\015\012    <x:Axis>\015\012     <x:AxisID>334446788</"
        "x:AxisID>\015\012     <x:ScaleID>407830976</x:ScaleID>\015\012     <x:Type>Categ"
        "ory</x:Type>\015\012     <x:MajorTick>Outside</x:MajorTick>\015\012     <x:Minor"
        "Tick>None</x:MinorTick>\015\012     <x:Placement>Bottom</x:Placement>\015\012   "
        "  <x:GroupingEnum>Auto</x:GroupingEnum>\015\012    </x:Axis>\015\012    <x:Axis>"
        "\015\012     <x:AxisID>334447260</x:AxisID>\015\012     <x:ScaleID>407831592</x:"
        "ScaleID>\015\012     <x:Type>Value</x:Type>\015\012     <x:Number>\015\012      "
        "<x:FormatString>#,###</x:FormatString>\015\012     </x:Number>\015\012     <x:Ma"
        "jorGridlines>\015\012     </x:MajorGridlines>\015\012     <x:MajorTick>Outside</"
        "x:MajorTick>\015\012     <x:MinorTick>None</x:MinorTick>\015\012     <x:Placemen"
        "t>Left</x:Placement>\015\012    </x:Axis>\015\012   </x:PlotArea>\015\012   <x:I"
        "dentifier></x:Identifier>\015\012  </x:Chart>\015\012  <x:Title>\015\012   <x:Fo"
        "nt>\015\012    <x:FontName>WeblogmaYekan</x:FontName>\015\012    <x:B>Automatic<"
        "/x:B>\015\012    <x:I>Automatic</x:I>\015\012    <x:U>Automatic</x:U>\015\012   "
        "</x:Font>\015\012   <x:Border>\015\012    <x:ColorIndex>None</x:ColorIndex>\015\012"
        "   </x:Border>\015\012   <x:Caption>\015\012    <x:DataSourceIndex>-1</x:DataSou"
        "rceIndex>\015\012    <x:Data>&quot;سود ماهيانه&quot;</x:Data>\015\012   </x:Capt"
        "ion>\015\012  </x:Title>\015\012  <x:Scaling>\015\012   <x:ScaleID>407830976</x:"
        "ScaleID>\015\012  </x:Scaling>\015\012  <x:Scaling>\015\012   <x:ScaleID>4078315"
        "92</x:ScaleID>\015\012  </x:Scaling>\015\012  <x:Scaling>\015\012   <x:ScaleID>4"
        "07831796</x:ScaleID>\015\012  </x:Scaling>\015\012  <x:HidePassiveAlert/>\015\012"
        "  <x:HideScreenTips/>\015\012 </x:ChartSpace>\015\012</xml>"
    AllowFormView =0
    RecordSourceQualifier ="dbo"
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
        End
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            OldBorderStyle =0
            Width =1701
            LabelX =-1701
            FontName ="Tahoma"
            AsianLineBreak =255
        End
        Begin Section
            Height =2394
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
                    Left =2267
                    Top =793
                    Name ="DATE_S"
                    ControlSource ="DATE_S"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4964
                            Top =793
                            Width =705
                            Height =240
                            Name ="Label0"
                            Caption ="ماه"
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
                    Left =2267
                    Top =1133
                    TabIndex =1
                    Name ="BEDS"
                    ControlSource ="BEDS"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =5160
                            Top =1140
                            Width =555
                            Height =405
                            Name ="Label1"
                            Caption ="دارئيها"
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
                    Left =2267
                    Top =1474
                    TabIndex =2
                    Name ="BESS"
                    ControlSource ="BESS"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =5175
                            Top =1470
                            Width =525
                            Height =240
                            Name ="Label2"
                            Caption ="بدهيها"
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
                    Left =2267
                    Top =1814
                    ColumnWidth =2280
                    TabIndex =3
                    Name ="sudday"
                    ControlSource ="sudday"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =5010
                            Top =1815
                            Width =720
                            Height =405
                            Name ="Label3"
                            Caption ="سود روزانه"
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
                    Left =2267
                    Top =2154
                    ColumnWidth =1875
                    TabIndex =4
                    Name ="SUD"
                    ControlSource ="SUD"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =5234
                            Top =2154
                            Width =435
                            Height =240
                            Name ="Label4"
                            Caption ="سود"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "sudmah.cls"
