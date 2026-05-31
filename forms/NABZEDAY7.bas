Version =20
VersionRequired =20
Begin Form
    AllowDeletions = NotDefault
    AllowAdditions = NotDefault
    AllowEdits = NotDefault
    DefaultView =4
    AllowUpdating =4
    ViewsAllowed =2
    TabularCharSet =178
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =5726
    RowHeight =375
    ItemSuffix =9
    Left =10755
    Top =660
    Right =15615
    Bottom =5220
    DatasheetForeColor =33554432
    DatasheetGridlinesColor =12632256
    MaxRecords =100000
    RecSrcDt = Begin
        0x2708a774b4f0e540
    End
    RecordSource ="SELECT TOP(7) COUNT(COMP_COD) AS CountOfCOMP_COD, STDATE, dbo.Uday(STDATE) AS DD"
        " FROM TASKS GROUP BY STDATE, dbo.Uday(STDATE) ORDER BY STDATE DESC, COUNT(COMP_C"
        "OD) DESC"
    Caption ="7 روزه"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Tahoma"
    PivotTableXML ="<xml xmlns:x=\"urn:schemas-microsoft-com:office:excel\">\015\012 <x:WorksheetOpt"
        "ions>\015\012  <x:DisplayRightToLeft/>\015\012 </x:WorksheetOptions>\015\012 <x:"
        "PivotTable>\015\012  <x:OWCVersion>10.0.0.6854         </x:OWCVersion>\015\012  "
        "<x:NoDisplayAlerts/>\015\012  <x:NoDisplayToolbar/>\015\012  <x:DisplayScreenTip"
        "s/>\015\012  <x:NoAutoFit/>\015\012  <x:Height>384</x:Height>\015\012  <x:Width>"
        "541</x:Width>\015\012  <x:NoAllowGrouping/>\015\012  <x:NoAllowFiltering/>\015\012"
        "  <x:CubeProvider>msolap.2</x:CubeProvider>\015\012  <x:CacheDetails/>\015\012  "
        "<x:Name>Microsoft Office PivotTable 10.0</x:Name>\015\012  <x:PivotField>\015\012"
        "   <x:Name>مقدار</x:Name>\015\012   <x:SourceName>CountOfCOMP_COD</x:SourceName>"
        "\015\012   <x:DataType>Integer</x:DataType>\015\012   <x:PLDataOrientation/>\015"
        "\012   <x:EncodedType>adInteger</x:EncodedType>\015\012   <x:CompareOrderedMembe"
        "rsBy>UniqueName</x:CompareOrderedMembersBy>\015\012  </x:PivotField>\015\012  <x"
        ":PivotField>\015\012   <x:Name>روز</x:Name>\015\012   <x:SourceName>STDATE</x:So"
        "urceName>\015\012   <x:DataType>Integer</x:DataType>\015\012   <x:EncodedType>ad"
        "Integer</x:EncodedType>\015\012   <x:CompareOrderedMembersBy>UniqueName</x:Compa"
        "reOrderedMembersBy>\015\012  </x:PivotField>\015\012  <x:PivotField>\015\012   <"
        "x:Name>روز1</x:Name>\015\012   <x:SourceName>DD</x:SourceName>\015\012   <x:Data"
        "Type>Integer</x:DataType>\015\012   <x:PLCaption>روز</x:PLCaption>\015\012   <x:"
        "FilterCaption>روز</x:FilterCaption>\015\012   <x:Orientation>Row</x:Orientation>"
        "\015\012   <x:EncodedType>adInteger</x:EncodedType>\015\012   <x:CompareOrderedM"
        "embersBy>UniqueName</x:CompareOrderedMembersBy>\015\012   <x:Expanded/>\015\012 "
        " </x:PivotField>\015\012  <x:PivotField>\015\012   <x:Name>Data</x:Name>\015\012"
        "   <x:Orientation>Column</x:Orientation>\015\012   <x:Position>-1</x:Position>\015"
        "\012   <x:DataField/>\015\012  </x:PivotField>\015\012  <x:PivotField>\015\012  "
        " <x:Name>Sum of مقدار</x:Name>\015\012   <x:PLName>Total1</x:PLName>\015\012   <"
        "x:TotalNumber>0</x:TotalNumber>\015\012   <x:Orientation>Data</x:Orientation>\015"
        "\012   <x:Position>1</x:Position>\015\012   <x:ParentField>مقدار</x:ParentField>"
        "\015\012  </x:PivotField>\015\012  <x:PivotData>\015\012   <x:Top>0.0</x:Top>\015"
        "\012   <x:TopOffset>0</x:TopOffset>\015\012   <x:Left>0</x:Left>\015\012   <x:Le"
        "ftOffset>0</x:LeftOffset>\015\012   <x:SeqNum>307</x:SeqNum>\015\012  </x:PivotD"
        "ata>\015\012  <x:PivotView>\015\012   <x:IsNotFiltered/>\015\012   <x:Label>\015"
        "\012    <x:Caption>Microsoft Office PivotTable 10.0</x:Caption>\015\012    <x:No"
        "tVisible/>\015\012   </x:Label>\015\012   <x:AllowEdits>false</x:AllowEdits>\015"
        "\012   <x:AllowAdditions>false</x:AllowAdditions>\015\012   <x:AllowDeletions>fa"
        "lse</x:AllowDeletions>\015\012  </x:PivotView>\015\012 </x:PivotTable>\015\012</"
        "xml>"
    PivotChartXML ="<xml xmlns:x=\"urn:schemas-microsoft-com:office:excel\">\015\012 <x:ChartSpace>\015"
        "\012  <x:OWCVersion>10.0.0.6854         </x:OWCVersion>\015\012  <x:Width>26882<"
        "/x:Width>\015\012  <x:Height>14658</x:Height>\015\012  <x:DisplayPropertyBrowser"
        "/>\015\012  <x:DataSource>\015\012   <x:Type>PivotList</x:Type>\015\012  </x:Dat"
        "aSource>\015\012  <x:BoundSeries>\015\012   <x:DataSourceIndex>0</x:DataSourceIn"
        "dex>\015\012  </x:BoundSeries>\015\012  <x:Category>\015\012   <x:DataSourceInde"
        "x>0</x:DataSourceIndex>\015\012  </x:Category>\015\012  <x:Value>\015\012   <x:D"
        "ataSourceIndex>0</x:DataSourceIndex>\015\012  </x:Value>\015\012  <x:BoundCharts"
        ">\015\012   <x:DataSourceIndex>0</x:DataSourceIndex>\015\012  </x:BoundCharts>\015"
        "\012  <x:PivotAggOrientation>Series</x:PivotAggOrientation>\015\012  <x:HidePivo"
        "tFields/>\015\012  <x:NoGrouping/>\015\012  <x:NoFiltering/>\015\012  <x:Palette"
        ">\015\012   <x:Entry>#000000</x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015\012"
        "   <x:Entry>#000000</x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015\012   <x:"
        "Entry>#000000</x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015\012   <x:Entry>"
        "#000000</x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015\012   <x:Entry>#00000"
        "0</x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015\012   <x:Entry>#000000</x:E"
        "ntry>\015\012   <x:Entry>#000000</x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015"
        "\012   <x:Entry>#000000</x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015\012  "
        " <x:Entry>#000000</x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015\012   <x:En"
        "try>#000000</x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015\012   <x:Entry>#0"
        "00000</x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015\012   <x:Entry>#000000<"
        "/x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015\012   <x:Entry>#000000</x:Ent"
        "ry>\015\012   <x:Entry>#8080FF</x:Entry>\015\012   <x:Entry>#802060</x:Entry>\015"
        "\012   <x:Entry>#FFFFA0</x:Entry>\015\012   <x:Entry>#A0E0E0</x:Entry>\015\012  "
        " <x:Entry>#600080</x:Entry>\015\012   <x:Entry>#FF8080</x:Entry>\015\012   <x:En"
        "try>#008080</x:Entry>\015\012   <x:Entry>#C0C0FF</x:Entry>\015\012   <x:Entry>#0"
        "00080</x:Entry>\015\012   <x:Entry>#FF00FF</x:Entry>\015\012   <x:Entry>#80FFFF<"
        "/x:Entry>\015\012   <x:Entry>#0080FF</x:Entry>\015\012   <x:Entry>#FF8080</x:Ent"
        "ry>\015\012   <x:Entry>#C0FF80</x:Entry>\015\012   <x:Entry>#FFC0FF</x:Entry>\015"
        "\012   <x:Entry>#FF80FF</x:Entry>\015\012  </x:Palette>\015\012  <x:DefaultFont>"
        "Arial</x:DefaultFont>\015\012  <x:Border>\015\012   <x:Color>#2F4F4F</x:Color>\015"
        "\012  </x:Border>\015\012  <x:Interior>\015\012   <x:Color>#008080</x:Color>\015"
        "\012   <x:FillEffect>\015\012    <x:fill x:type=\"Solid\" x:color=\"#008080\"/>\015"
        "\012   </x:FillEffect>\015\012  </x:Interior>\015\012  <x:Chart>\015\012   <x:Pl"
        "otArea>\015\012    <x:Interior>\015\012     <x:Color>#87CEEB</x:Color>\015\012  "
        "   <x:FillEffect>\015\012      <x:fill x:type=\"Solid\" x:color=\"#87CEEB\"/>\015"
        "\012     </x:FillEffect>\015\012    </x:Interior>\015\012    <x:Graph>\015\012  "
        "   <x:SubType>Smooth</x:SubType>\015\012     <x:Type>Line</x:Type>\015\012     <"
        "x:SubType>Standard</x:SubType>\015\012     <x:Series>\015\012      <x:Line>\015\012"
        "       <x:Color>#000000</x:Color>\015\012      </x:Line>\015\012      <x:FormatM"
        "ap>\015\012      </x:FormatMap>\015\012      <x:Name>Series</x:Name>\015\012    "
        "  <x:Caption>\015\012       <x:DataSourceIndex>-1</x:DataSourceIndex>\015\012   "
        "    <x:Data>&quot;Series&quot;</x:Data>\015\012      </x:Caption>\015\012      <"
        "x:Index>0</x:Index>\015\012      <x:Category>\015\012       <x:DataSourceIndex>0"
        "</x:DataSourceIndex>\015\012      </x:Category>\015\012      <x:Value>\015\012  "
        "     <x:DataSourceIndex>0</x:DataSourceIndex>\015\012      </x:Value>\015\012   "
        "   <x:Marker>\015\012       <x:Symbol>None</x:Symbol>\015\012      </x:Marker>\015"
        "\012      <x:Explode>0</x:Explode>\015\012      <x:Thickness>10</x:Thickness>\015"
        "\012      <x:Smooth/>\015\012      <x:DataSourceIndex>0</x:DataSourceIndex>\015\012"
        "      <x:Identifier>!.Sum of مقدار</x:Identifier>\015\012     </x:Series>\015\012"
        "     <x:Dimension>\015\012      <x:ScaleID>245909996</x:ScaleID>\015\012      <x"
        ":Index>Categories</x:Index>\015\012     </x:Dimension>\015\012     <x:Dimension>"
        "\015\012      <x:ScaleID>245910200</x:ScaleID>\015\012      <x:Index>Value</x:In"
        "dex>\015\012     </x:Dimension>\015\012     <x:Dimension>\015\012      <x:ScaleI"
        "D>245910404</x:ScaleID>\015\012      <x:Index>FormatValue</x:Index>\015\012     "
        "</x:Dimension>\015\012     <x:Overlap>100</x:Overlap>\015\012     <x:GapWidth>15"
        "0</x:GapWidth>\015\012     <x:FirstSliceAngle>0</x:FirstSliceAngle>\015\012    <"
        "/x:Graph>\015\012    <x:Axis>\015\012     <x:AxisID>245913292</x:AxisID>\015\012"
        "     <x:ScaleID>245909996</x:ScaleID>\015\012     <x:Type>Category</x:Type>\015\012"
        "     <x:Font>\015\012      <x:FontName>Tahoma</x:FontName>\015\012      <x:Size>"
        "7</x:Size>\015\012      <x:B>Automatic</x:B>\015\012      <x:I>Automatic</x:I>\015"
        "\012      <x:U>Automatic</x:U>\015\012     </x:Font>\015\012     <x:MajorTick>Ou"
        "tside</x:MajorTick>\015\012     <x:MinorTick>None</x:MinorTick>\015\012     <x:P"
        "lacement>Bottom</x:Placement>\015\012     <x:GroupingEnum>Auto</x:GroupingEnum>\015"
        "\012    </x:Axis>\015\012    <x:Axis>\015\012     <x:AxisID>307435752</x:AxisID>"
        "\015\012     <x:ScaleID>245910200</x:ScaleID>\015\012     <x:Type>Value</x:Type>"
        "\015\012     <x:Number>\015\012      <x:FormatString>General</x:FormatString>\015"
        "\012     </x:Number>\015\012     <x:MajorTick>Outside</x:MajorTick>\015\012     "
        "<x:MinorTick>None</x:MinorTick>\015\012     <x:Placement>Left</x:Placement>\015\012"
        "    </x:Axis>\015\012   </x:PlotArea>\015\012   <x:Identifier></x:Identifier>\015"
        "\012  </x:Chart>\015\012  <x:Title>\015\012   <x:Font>\015\012    <x:FontName>B "
        "Titr</x:FontName>\015\012    <x:Size>7</x:Size>\015\012    <x:Color>#FFA500</x:C"
        "olor>\015\012    <x:B>Automatic</x:B>\015\012    <x:I>Automatic</x:I>\015\012   "
        " <x:U>Automatic</x:U>\015\012   </x:Font>\015\012   <x:Border>\015\012    <x:Col"
        "orIndex>None</x:ColorIndex>\015\012   </x:Border>\015\012   <x:Caption>\015\012 "
        "   <x:DataSourceIndex>-1</x:DataSourceIndex>\015\012    <x:Data>&quot; فعاليتهاي"
        "  7 روزه&quot;</x:Data>\015\012   </x:Caption>\015\012  </x:Title>\015\012  <x:S"
        "caling>\015\012   <x:ScaleID>245909996</x:ScaleID>\015\012  </x:Scaling>\015\012"
        "  <x:Scaling>\015\012   <x:ScaleID>245910200</x:ScaleID>\015\012  </x:Scaling>\015"
        "\012  <x:Scaling>\015\012   <x:ScaleID>245910404</x:ScaleID>\015\012  </x:Scalin"
        "g>\015\012  <x:HidePassiveAlert/>\015\012  <x:HideScreenTips/>\015\012 </x:Chart"
        "Space>\015\012</xml>"
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
            Height =2734
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2551
                    Top =453
                    ColumnWidth =1590
                    ColumnOrder =2
                    Name ="CountOfCOMP_COD"
                    ControlSource ="CountOfCOMP_COD"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =850
                            Top =453
                            Width =645
                            Height =240
                            Name ="Label0"
                            Caption ="مقدار"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2551
                    Top =1834
                    ColumnOrder =1
                    TabIndex =1
                    Name ="DATE_N"
                    ControlSource ="STDATE"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =850
                            Top =1834
                            Width =480
                            Height =240
                            Name ="Label5"
                            Caption ="روز"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1927
                    Top =1247
                    TabIndex =2
                    Name ="DD"
                    ControlSource ="DD"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =226
                            Top =1247
                            Width =480
                            Height =240
                            Name ="Label10"
                            Caption ="روز"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "NABZEDAY7.cls"
