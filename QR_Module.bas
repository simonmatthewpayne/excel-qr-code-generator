Attribute VB_Name = "Module1"
Option Explicit

' ============================================================
' MAIN QR GENERATOR
' ============================================================
Public Sub GenerateBulkQRCodes()
    Dim srcRange As Range
    Dim cell As Range
    Dim outColInput As String
    Dim outCol As Long
    Dim ws As Worksheet
    
    If TypeName(Selection) <> "Range" Then
        MsgBox "Select the cells containing the QR values first.", vbExclamation
        Exit Sub
    End If
    
    Set srcRange = Selection
    Set ws = srcRange.Worksheet
    
    outColInput = InputBox("Enter output column letter for QR codes (e.g. F):", _
                           "QR Output Column", "F")
    If outColInput = "" Then Exit Sub
    
    On Error Resume Next
    outCol = ws.Range(outColInput & "1").Column
    On Error GoTo 0
    If outCol = 0 Then
        MsgBox "Invalid column letter.", vbCritical
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    
    Dim qrText As String
    Dim qrUrlEncoded As String
    Dim tempFile As String
    Dim pic As Picture
    Dim tgtCell As Range
    Dim s As Shape
    Dim side As Double
    
    For Each cell In srcRange.Cells
        qrText = cell.Text            'exactly what is displayed in the cell
        If Len(qrText) > 0 Then
            
            qrUrlEncoded = SafeUrlEncode(qrText)
            tempFile = DownloadQRToTempFile(qrUrlEncoded, 300)
            
            If tempFile <> "" Then
                Set tgtCell = ws.Cells(cell.Row, outCol)
                
                ' Insert the picture
                Set pic = ws.Pictures.Insert(tempFile)
                Set s = pic.ShapeRange(1)
                
                With s
                    ' Name encodes the cell it belongs to, e.g. QR_F12
                    .Name = "QR_" & tgtCell.Address(False, False)
                    .LockAspectRatio = msoTrue
                    
                    ' --- SIZE CONTROL ---
                    ' Make the QR a square controlled by ROW HEIGHT
                    side = tgtCell.Height     ' row height defines size
                    .Height = side
                    .Width = side
                    
                    ' Centre horizontally in the cell, top aligned
                    .Left = tgtCell.Left + (tgtCell.Width - .Width) / 2
                    .Top = tgtCell.Top
                    
                    ' Move & size with the cell (sorting / row-height changes)
                    .Placement = xlMoveAndSize
                    
                    ' Mark as locked so you can't drag it when sheet is protected
                    .Locked = True
                End With
                
                ' Delete temp file
                On Error Resume Next
                Kill tempFile
                On Error GoTo 0
            End If
        End If
    Next cell
    
    Application.ScreenUpdating = True
    MsgBox "QR codes generated.", vbInformation
End Sub


' ============================================================
' SNAP ALL QRS BACK TO THEIR CELLS
' (Run this after changing row heights / column widths)
' ============================================================
Public Sub SnapAllQRCodesToCells()
    Dim ws As Worksheet
    Dim s As Shape
    Dim addr As String
    Dim tgtCell As Range
    Dim side As Double
    
    Set ws = ActiveSheet
    
    For Each s In ws.Shapes
        If Left$(s.Name, 3) = "QR_" Then
            addr = Mid$(s.Name, 4)   'strip off "QR_"
            
            On Error Resume Next
            Set tgtCell = ws.Range(addr)
            On Error GoTo 0
            
            If Not tgtCell Is Nothing Then
                With s
                    .LockAspectRatio = msoTrue
                    
                    ' square = row height again
                    side = tgtCell.Height
                    .Height = side
                    .Width = side
                    
                    .Left = tgtCell.Left + (tgtCell.Width - .Width) / 2
                    .Top = tgtCell.Top
                    
                    .Placement = xlMoveAndSize
                    .Locked = True
                End With
            End If
        End If
        Set tgtCell = Nothing
    Next s
    
    MsgBox "QR codes snapped to their cells.", vbInformation
End Sub


' ============================================================
' URL ENCODER
' ============================================================
Private Function SafeUrlEncode(ByVal txt As String) As String
    On Error Resume Next
    SafeUrlEncode = Application.WorksheetFunction.EncodeURL(txt)
    If Err.Number <> 0 Or SafeUrlEncode = "" Then
        Err.Clear
        SafeUrlEncode = Replace(txt, " ", "%20")
    End If
    On Error GoTo 0
End Function


' ============================================================
' DOWNLOAD QR PNG FROM API
' ============================================================
Private Function DownloadQRToTempFile(ByVal urlEncodedData As String, _
                                      Optional ByVal size As Long = 300) As String
    Dim http As Object, stm As Object
    Dim tmpPath As String, filePath As String
    Dim url As String
    
    url = "https://api.qrserver.com/v1/create-qr-code/?size=" & size & "x" & size & _
          "&data=" & urlEncodedData
    
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "GET", url, False
    http.Send
    
    If http.Status <> 200 Then
        DownloadQRToTempFile = ""
        Exit Function
    End If
    
    tmpPath = Environ$("TEMP")
    filePath = tmpPath & "\qr_" & Format(Now, "yyyymmdd_hhnnss") & _
               "_" & Replace(CStr(Rnd), "0.", "") & ".png"
    
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 1   'binary
    stm.Open
    stm.Write http.responseBody
    stm.SaveToFile filePath, 2   'overwrite
    stm.Close
    
    DownloadQRToTempFile = filePath
End Function


