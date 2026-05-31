Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Open(CANCEL As Integer)
Dim objBrowser As Object, rst As New ADODB.Recordset
                     On Error Resume Next
    Dim HtmlCode As String
    Dim i, j, N As Integer
    Dim CenterLat As Double, CenterLng As Double, MapDraggable As Integer, Zoom As Integer, MapTypeControl As Integer, MapTypeControlID As String, DispZoomControl As Integer, ZoomControlPosition As String, ZoomControlStyle As String, DispScaleControl As Integer, ScaleControlPosition As String, DispMarkerCenter As Integer, DispMarkers As Integer, MarkersClickable As Integer, MarkersDragable As Integer, DispPolyline As Integer, PolyLineColor As String, PolyLineOpacity As Double, PolyLineWidth As Integer

    Zoom = 15
    MapDraggable = 1
    DispZoomControl = 1
    ZoomControlPosition = "LEFT_CENTER"
    ZoomControlStyle = "LARGE"
    DispScaleControl = 1
    ScaleControlPosition = "TOP_RIGHT"
    MapTypeControl = 1
    MapTypeControlID = "SATELLITE"
    DispMarkerCenter = 1
    DispPolyline = 1
    PolyLineColor = "fffaf0"
    PolyLineOpacity = 0.5
    PolyLineWidth = 3
    DispMarkers = 1
    MarkersClickable = 1
    MarkersDragable = 1
    rst.Open "SELECT     VISITOR, FDATE, EDATE, LATITUDE, LONGITUDE, SPEED FROM dbo.GPSLOCATION WHERE     (VISITOR = N'" & Forms![F_MENU_DATE_VISIT]![CUST_NO] & "') AND (FDATE BETWEEN " & Forms![F_MENU_DATE_VISIT]![DT1] & " AND " & Forms![F_MENU_DATE_VISIT]![DT2] & ") ORDER BY IDD", CurrentProject.Connection, adOpenKeyset, adLockOptimistic

    If rst.RecordCount > 0 Then
        CenterLat = rst.Fields("LATITUDE")
        CenterLng = rst.Fields("LONGITUDE")
    HtmlCode = "<!DOCTYPE HTML> <html lang=" & Chr(34) & "en" & Chr(34) & "> <head> <meta charset=" & Chr(34) & "UTF-8" & Chr(34) & " /> <title>Google Maps</title> <style type=" & Chr(34) & "text/css" & Chr(34) & "> #map-canvas{width: 900px;height: 600px;margin: 0 auto;}</style> <script type=" & Chr(34) & "text/javascript" & Chr(34) & " src=" & Chr(34) & "https://maps.googleapis.com/maps/api/js?v=3.exp&sensor=false" & Chr(34) & "></script><script type=" & Chr(34) & "text/javascript" & Chr(34) & ">" & _
                " function initialize() {var latlng = new google.maps.LatLng(" & CenterLat & "," & CenterLng & "); var myOptions = { zoom: " & Zoom & ", center: latlng, disableDefaultUI: true, draggable: " & MapDraggable & ", zoomControl: " & DispZoomControl & ", zoomControlOptions: {position: google.maps.ControlPosition." & ZoomControlPosition & ", style: google.maps.ZoomControlStyle." & ZoomControlStyle & "}, scaleControl: " & DispScaleControl & ", scaleControlOptions: {position: google.maps.ControlPosition." & ScaleControlPosition & "}, mapTypeControl: " & MapTypeControl & ", mapTypeId: google.maps.MapTypeId." & MapTypeControlID & _
                 " }; var map = new google.maps.Map(document.getElementById('map-canvas'),  myOptions)"

    ' adds tilesloaded event


    ' adds zoom_changed event, returns the zoom level

    ' displays a marker on the map center
    If DispMarkerCenter = 1 Then
        HtmlCode = HtmlCode & "; var markerc =         new google.maps.Marker({ position: new google.maps.LatLng(" & CenterLat & "," & CenterLng & "),map: map, title: '',clickable: true,icon: 'http://www.google.com/mapfiles/arrow.png' })"
    End If
 '   ----
     If rst.RecordCount > 0 Then
    ' displays a polyline between points
        j = rst.RecordCount - 1
        If DispPolyline = 1 And j > 0 Then
            ' defines an array 'points'
            HtmlCode = HtmlCode & "; var points = [];"
            For i = 0 To j
'               Pos = MarkerPos.Get(i)
                ' sets the current position to the 'point' variable defines an array 'points'
                HtmlCode = HtmlCode & "var point = new google.maps.LatLng(" & rst.Fields("LATITUDE") & "," & rst.Fields("LONGITUDE") & ");"
                ' adds the current point to the points array
                HtmlCode = HtmlCode & "points.push(point);"
                rst.MoveNext
            Next
            ' defines the polyline
            HtmlCode = HtmlCode & "; var polyline = new google.maps.Polyline({path: points, strokeColor: '" & PolyLineColor & "', strokeOpacity: " & PolyLineOpacity & ", strokeWeight: " & PolyLineWidth & "});"
            ' draws the polyline
'           HtmlCode = HtmlCode & "polyline.setEditable(true);"
            HtmlCode = HtmlCode & "polyline.setMap(map);"
        End If

        ' displays markers on the map
       ' GPSMarkerIndexes.Initialize
       N = 0
        If DispMarkers = 1 Then  ' checks if display of markers enabled
            rst.MoveFirst
            For i = 0 To j
'                Pos = MarkerPos.Get(I)

                If (i Mod 10) = 0 Then    ' checks if the current marker should be displayed
                    'GPSMarkerIndexes.ADD (I)
                    N = N + 1
                    ' sets the current position to a marker
                    HtmlCode = HtmlCode & "; var marker" & N & " = new google.maps.Marker({ position: new google.maps.LatLng(" & rst.Fields("LATITUDE") & "," & rst.Fields("LONGITUDE") & "),map: map, title: '" & "شروع مسير ويزيتور: " & GETHESNAME(rst.Fields("VISITOR")) & " نقطه :  " & i & " ',clickable: true , draggable: true , icon: 'http://chart.apis.google.com/"
                    If i = 0 Then
                        ' Displays a green A pin
                        ' A|00FF00|000000   A = character to display, 00FF00 fill color  000000 text color
                        HtmlCode = HtmlCode & "chart?chst=d_map_pin_letter&chld=A|00FF00|000000'});"
                    Else
                        If i = j Then
                        ' Displays a red
                            HtmlCode = HtmlCode & "chart?chst=d_map_pin_letter&chld=" & Chr((N Mod 26) + 65) & "|FF0000|000000'});"
                        Else
                        ' Displays a yellow pin
                            HtmlCode = HtmlCode & "chart?chst=d_map_pin_letter&chld=" & Chr((N Mod 26) + 65) & "|FFFF00|000000'});"
                        End If
                    End If
                    If MarkersClickable = 1 Then
                        ' adds an InfoWondow to each marker
                        Dim TXT As String
                        If i = 0 Then
                            TXT = "<b>" & GETHESNAME(rst.Fields("VISITOR")) & "</b>"      ' <br> = crlf
                            TXT = TXT & "<br><b>تاريخ : </b>" & Format(rst.Fields("FDATE"), "####/##/##")
                            TXT = TXT & "<br><b>ساعت : </b>" & Format(RIGHT("000000" & rst.Fields("EDATE"), 6), "00:00:00")
                            TXT = TXT & "<br><b>توضيح : </b>"
                            TXT = TXT & "<br>" & "سفير"
                  '      Else
                 '           txt = "<b>" & GETHESNAME(RST.Fields("VISITOR")) & "</b>"
                '            txt = txt & "<br><b>نقطه : </b>" & I
               '             txt = txt & "<br><b>طول : </b>" & RST.Fields("LATITUDE")
              '              txt = txt & "<br><b>عرض : </b>" & RST.Fields("LONGITUDE")
                        End If
                        HtmlCode = HtmlCode & "var infowindow" & i & " = new google.maps.InfoWindow("
                        HtmlCode = HtmlCode & "{content: '" & TXT & "', "
                        HtmlCode = HtmlCode & "disableAutoPan: false"
'                           HtmlCode = HtmlCode & "disableAutoPan: true"
                        HtmlCode = HtmlCode & "}); "

                        If i = 0 Then
                            ' opens the InfoWindow of the first marker directly
'                           HtmlCode = HtmlCode & "infowindow" & i & ".open(map, marker" & i & "); "
                        End If

                        ' adds click event to the marker and opens an InfoWindow
                        HtmlCode = HtmlCode & "google.maps.event.addListener(marker" & N & ", 'click', function(mouseEvent) {"
                        HtmlCode = HtmlCode & "infowindow" & i & ".open(map, marker" & N & "); "
                        HtmlCode = HtmlCode & "}); "

                        ' adds closeclick event to the InfoWindow
                    End If

                    If MarkersDragable = 1 Then
                        ' Adds the 'dragend' event listener
                        HtmlCode = HtmlCode & "; google.maps.event.addListener(marker" & N & ", 'dragend', function(mouseEvent) {"
                        ' mouseEvent.latLng = final drag position
                        HtmlCode = HtmlCode & "var lat = mouseEvent.latLng.lat();"
                        HtmlCode = HtmlCode & "var lng = mouseEvent.latLng.lng();"
                        HtmlCode = HtmlCode & "var latlng1 = new google.maps.LatLng(lat, lng); "
                        ' Centers the map to that location
                        HtmlCode = HtmlCode & "map.setCenter(latlng1);"
                        ' Adds the B4A 'MapViewer_MarkerDragged' event call

                        ' doesn't work
                        HtmlCode = HtmlCode & "points.setAt(" & i & ", latlng1);"
                        HtmlCode = HtmlCode & "polyline.setMap(map);"
                    End If

                    HtmlCode = HtmlCode & "})"
                End If
                rst.MoveNext
            Next
        End If
    End If

 '   ----
    HtmlCode = HtmlCode & "; }</script></head><body onload=" & Chr(34) & "initialize()" & Chr(34) & ">  <div id=" & Chr(34) & "map-canvas" & Chr(34) & " ></div></body></html>"
    Dim fs, a
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set a = fs.CREATETextFile("c:\masir.htm", True)
    a.writeline (HtmlCode)
    a.Close
        Set objBrowser = Me.WebBrowser0
        With Me.WebBrowser0
            .Navigate "c:\masir.htm"

        End With
       ' DoCmd.Hourglass False
    Else
        DoCmd.OpenForm "mesageform", acNormal, , , acFormReadOnly, acDialog, "ويزيتور فاقد فعاليت در اين محدوده تاريخي است"
    End If
End Sub

Private Sub Form_Close()
   DoCmd.Close acForm, "F_MENU_DATE_VISIT"
End Sub
