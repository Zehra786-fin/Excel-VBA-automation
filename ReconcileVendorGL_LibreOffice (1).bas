' ==========================================================================
' RECONCILIATION MACRO (LibreOffice Basic) - Vendor_Ledger vs General_Ledger
'
' Matching key   : VendorID + RefNo  (RefNo is the shared invoice/payment
'                  reference across both systems - DocNo differs between
'                  the two, since each system generates its own internal
'                  document numbers)
' Amount compared: Vendor.Amount (col H)  vs  GL.NetAmount (col J)
' Tolerance      : RECON_THRESHOLD (edit below)
'
' Update VENDOR_SHEET / GL_SHEET constants if your tab names differ.
' ==========================================================================

Sub ReconcileVendorGL
    Dim oDoc As Object, oSheets As Object
    Dim oVendor As Object, oGL As Object, oSummary As Object
    Dim oCursor As Object, oHeaderRange As Object
    Dim i As Integer, j As Integer, k As Integer
    Dim lastVendorRow As Integer, lastGLRow As Integer
    Dim glKey() As String, glRow() As Integer, glMatched() As Boolean
    Dim nGL As Integer
    Dim vendorAmount As Double, glAmount As Double, diff As Double
    Dim matchedCount As Integer, mismatchCount As Integer
    Dim missingGLCount As Integer, missingVendorCount As Integer
    Dim summaryRow As Integer
    Dim key As String, vVid As String, vDoc As String
    Dim foundIndex As Integer

    Const VENDOR_SHEET As String = "Vendor_Ledger"
    Const GL_SHEET As String = "General_Ledger"
    Const SUMMARY_SHEET As String = "Reconciliation_Summary"
    Const RECON_THRESHOLD As Double = 1

    oDoc = ThisComponent
    oSheets = oDoc.getSheets()
    oVendor = oSheets.getByName(VENDOR_SHEET)
    oGL = oSheets.getByName(GL_SHEET)

    ' Recreate summary sheet fresh each run
    If oSheets.hasByName(SUMMARY_SHEET) Then
        oSheets.removeByName(SUMMARY_SHEET)
    End If
    oSheets.insertNewByName(SUMMARY_SHEET, oSheets.getCount())
    oSummary = oSheets.getByName(SUMMARY_SHEET)

    oSummary.getCellByPosition(0,0).setString("Key (VendorID|RefNo)")
    oSummary.getCellByPosition(1,0).setString("VendorName")
    oSummary.getCellByPosition(2,0).setString("Vendor Amount")
    oSummary.getCellByPosition(3,0).setString("GL NetAmount")
    oSummary.getCellByPosition(4,0).setString("Difference")
    oSummary.getCellByPosition(5,0).setString("Status")
    oHeaderRange = oSummary.getCellRangeByPosition(0,0,5,0)
    oHeaderRange.CharWeight = com.sun.star.awt.FontWeight.BOLD

    summaryRow = 1   ' zero-indexed; row index 1 = visible row 2

    ' Helper column headers (zero-indexed columns)
    oVendor.getCellByPosition(13,0).setString("Recon_Status")        ' col N
    oVendor.getCellByPosition(14,0).setString("Difference")          ' col O
    oGL.getCellByPosition(15,0).setString("Recon_Status")            ' col P
    oGL.getCellByPosition(16,0).setString("Matched_VendorRow")       ' col Q

    ' Find last used row on each sheet
    oCursor = oVendor.createCursor()
    oCursor.gotoEndOfUsedArea(False)
    lastVendorRow = oCursor.RangeAddress.EndRow

    oCursor = oGL.createCursor()
    oCursor.gotoEndOfUsedArea(False)
    lastGLRow = oCursor.RangeAddress.EndRow

    ' Clear any leftover background color from previous runs (whole used range)
    oVendor.getCellRangeByPosition(0, 0, oVendor.createCursor().RangeAddress.EndColumn, lastVendorRow).CellBackColor = -1
    oGL.getCellRangeByPosition(0, 0, oGL.createCursor().RangeAddress.EndColumn, lastGLRow).CellBackColor = -1

    ' Build GL lookup arrays (rows 1..lastGLRow, row 0 is header)
    nGL = lastGLRow
    ReDim glKey(1 To nGL)
    ReDim glRow(1 To nGL)
    ReDim glMatched(1 To nGL)

    For j = 1 To lastGLRow
        glKey(j) = Trim(oGL.getCellByPosition(3, j).getString()) & "|" & Trim(oGL.getCellByPosition(5, j).getString())
        glRow(j) = j
        glMatched(j) = False
        oGL.getCellByPosition(15, j).setString("Unmatched")
    Next j

    matchedCount = 0 : mismatchCount = 0
    missingGLCount = 0 : missingVendorCount = 0

    ' Loop through Vendor rows and try to find a GL match
    For i = 1 To lastVendorRow
        vVid = Trim(oVendor.getCellByPosition(3, i).getString())   ' col D (VendorID)
        vDoc = Trim(oVendor.getCellByPosition(5, i).getString())   ' col F (RefNo)
        key = vVid & "|" & vDoc
        vendorAmount = oVendor.getCellByPosition(7, i).getValue()  ' col H

        foundIndex = -1
        For k = 1 To nGL
            If glKey(k) = key And Not glMatched(k) Then
                foundIndex = k
                Exit For
            End If
        Next k

        If foundIndex > -1 Then
            j = glRow(foundIndex)
            glAmount = oGL.getCellByPosition(9, j).getValue()       ' col J (NetAmount)
            diff = vendorAmount - glAmount

            If Abs(diff) <= RECON_THRESHOLD Then
                oVendor.getCellByPosition(13, i).setString("Matched")
                oVendor.getCellByPosition(14, i).setValue(diff)
                oVendor.getCellByPosition(13, i).CellBackColor = RGB(198,239,206)
                oGL.getCellByPosition(15, j).setString("Matched")
                oGL.getCellByPosition(15, j).CellBackColor = RGB(198,239,206)
                matchedCount = matchedCount + 1
            Else
                oVendor.getCellByPosition(13, i).setString("Mismatch")
                oVendor.getCellByPosition(14, i).setValue(diff)
                oVendor.getCellByPosition(13, i).CellBackColor = RGB(255,235,156)
                oGL.getCellByPosition(15, j).setString("Mismatch")
                oGL.getCellByPosition(15, j).CellBackColor = RGB(255,235,156)
                mismatchCount = mismatchCount + 1

                oSummary.getCellByPosition(0, summaryRow).setString(key)
                oSummary.getCellByPosition(1, summaryRow).setString(oVendor.getCellByPosition(4, i).getString())
                oSummary.getCellByPosition(2, summaryRow).setValue(vendorAmount)
                oSummary.getCellByPosition(3, summaryRow).setValue(glAmount)
                oSummary.getCellByPosition(4, summaryRow).setValue(diff)
                oSummary.getCellByPosition(5, summaryRow).setString("Mismatch")
                summaryRow = summaryRow + 1
            End If

            oGL.getCellByPosition(16, j).setValue(i + 1)   ' 1-based row ref for readability
            glMatched(foundIndex) = True

        Else
            oVendor.getCellByPosition(13, i).setString("Missing in GL")
            oVendor.getCellByPosition(14, i).setString("N/A")
            oVendor.getCellByPosition(13, i).CellBackColor = RGB(255,199,206)
            missingGLCount = missingGLCount + 1

            oSummary.getCellByPosition(0, summaryRow).setString(key)
            oSummary.getCellByPosition(1, summaryRow).setString(oVendor.getCellByPosition(4, i).getString())
            oSummary.getCellByPosition(2, summaryRow).setValue(vendorAmount)
            oSummary.getCellByPosition(3, summaryRow).setString("N/A")
            oSummary.getCellByPosition(4, summaryRow).setString("N/A")
            oSummary.getCellByPosition(5, summaryRow).setString("Missing in GL")
            summaryRow = summaryRow + 1
        End If
    Next i

    ' Any GL rows never matched = Missing in Vendor
    For k = 1 To nGL
        If Not glMatched(k) Then
            j = glRow(k)
            oGL.getCellByPosition(15, j).CellBackColor = RGB(255,199,206)
            missingVendorCount = missingVendorCount + 1

            oSummary.getCellByPosition(0, summaryRow).setString(glKey(k))
            oSummary.getCellByPosition(1, summaryRow).setString(oGL.getCellByPosition(4, j).getString())
            oSummary.getCellByPosition(2, summaryRow).setString("N/A")
            oSummary.getCellByPosition(3, summaryRow).setValue(oGL.getCellByPosition(9, j).getValue())
            oSummary.getCellByPosition(4, summaryRow).setString("N/A")
            oSummary.getCellByPosition(5, summaryRow).setString("Missing in Vendor")
            summaryRow = summaryRow + 1
        End If
    Next k

    MsgBox "Reconciliation Complete!" & Chr(10) & Chr(10) & _
           "Matched: " & matchedCount & Chr(10) & _
           "Mismatch: " & mismatchCount & Chr(10) & _
           "Missing in GL: " & missingGLCount & Chr(10) & _
           "Missing in Vendor: " & missingVendorCount, 0, "Reconciliation Summary"

End Sub
