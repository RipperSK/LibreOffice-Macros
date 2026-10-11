REM  *****  BASIC  *****

Sub CleanBankExport
    Dim oSheet As Object, oCur As Object, oRange As Object, oDesc As Object
    Dim i As Long, lastRow As Long, lastCol As Long

    oSheet = ThisComponent.CurrentController.ActiveSheet

    ' --- hide columns (0-based) ---
    Dim hideCols As Variant
    hideCols = Array(0,3,7,8,9,10,11,13,14,15,20,21)
    For i = 0 To UBound(hideCols)
        oSheet.Columns.getByIndex(hideCols(i)).IsVisible = False
    Next

    oCur = oSheet.createCursor()
    oCur.gotoEndOfUsedArea(False)
    lastRow = oCur.RangeAddress.EndRow
    lastCol = oCur.RangeAddress.EndColumn
    oRange = oSheet.getCellRangeByPosition(0, 0, lastCol, lastRow)

    ' --- sort ascending on column 0 ---
    Dim aSortFields(0) As New com.sun.star.table.TableSortField
    aSortFields(0).Field = 0
    aSortFields(0).IsAscending = True

    Dim aSortDesc(1) As New com.sun.star.beans.PropertyValue
    aSortDesc(0).Name = "SortFields"
    aSortDesc(0).Value = aSortFields()
    aSortDesc(1).Name = "ContainsHeader"
    aSortDesc(1).Value = True
    oRange.sort(aSortDesc())

    ' --- row filter ---
    ' --- AutoFilter on header row (create if missing) ---
    Const DB_NAME As String = "BankExport"
    Dim oDBs As Object, oDB As Object
    oDBs = ThisComponent.DatabaseRanges

    If oDBs.hasByName(DB_NAME) Then oDBs.removeByName(DB_NAME)
    oDBs.addNewByName(DB_NAME, oRange.RangeAddress)
    oDB = oDBs.getByName(DB_NAME)
    oDB.AutoFilter = True        ' dropdown arrows on row 1
    oDB.ContainsHeader = True

    oDesc = oDB.getFilterDescriptor()
    oDesc.UseRegularExpressions = True
    oDesc.IsCaseSensitive = False

    Dim aField(1) As New com.sun.star.sheet.TableFilterField

    ' condition 1: column 18 must NOT match the fee/POS patterns
    aField(0).Field = 18
    aField(0).Operator = com.sun.star.sheet.FilterOperator.NOT_EQUAL
    aField(0).IsNumeric = False
    aField(0).StringValue = "(Poplatok za balík|Poplatky za transakcie|POS|Transakčná daň).*"

    ' condition 2: column E (index 4) must be "Debet"
    aField(1).Connection = com.sun.star.sheet.FilterConnection.AND
    aField(1).Field = 4
    aField(1).Operator = com.sun.star.sheet.FilterOperator.EQUAL
    aField(1).IsNumeric = False
    aField(1).StringValue = "Debet"

    oDesc.setFilterFields(aField())
    oDB.refresh()

End Sub

Sub Main
    CleanBankExport
End Sub
