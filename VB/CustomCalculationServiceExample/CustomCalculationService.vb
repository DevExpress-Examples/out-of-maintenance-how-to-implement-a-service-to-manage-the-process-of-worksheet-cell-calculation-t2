#Region "#namespaces"
Imports DevExpress.Spreadsheet
Imports DevExpress.Spreadsheet.Formulas
Imports DevExpress.XtraTreeList.Nodes
Imports System
Imports System.Collections.Generic
#End Region ' #namespaces

Namespace CustomCalculationServiceExample
	#Region "#TestCustomCalculationService"
	Public Class TestCustomCalculationService
		Implements DevExpress.XtraSpreadsheet.Services.ICustomCalculationService

		Private loggingControl As DevExpress.XtraTreeList.TreeList
		Private rootListNode As TreeListNode = Nothing
		Public CircularReferencedCells As New HashSet(Of CellKey)()

		Public Sub New(ByVal controlForLogging As Object)
			loggingControl = TryCast(controlForLogging, DevExpress.XtraTreeList.TreeList)
		End Sub

		Public Function OnBeginCalculation() As Boolean Implements DevExpress.XtraSpreadsheet.Services.ICustomCalculationService.OnBeginCalculation
			' Set the root node for displaying info on the current calculation.
			Dim rootNode As TreeListNode = Nothing
			rootListNode = loggingControl.AppendNode(New Object() {String.Format("Calculation starts at {0}", Date.Now)}, rootNode)
			' To highlight cells with circular references, this example uses a hash set of such cells. 
			' Clear it when a new calculation starts.
			CircularReferencedCells.Clear()
			' True to perform a calculation. Return False to cancel it.
			Return True
		End Function
		Public Sub OnBeginCellCalculation(ByVal args As CellCalculationArgs) Implements DevExpress.XtraSpreadsheet.Services.ICustomCalculationService.OnBeginCellCalculation
			' Add a record about the cell being calculated.
			CreateLogEntry("Cell calculation begins", New String() {String.Format("CellKey: ({0})", args.CellKey)})
		End Sub
		Public Function OnBeginCircularReferencesCalculation() As Boolean Implements DevExpress.XtraSpreadsheet.Services.ICustomCalculationService.OnBeginCircularReferencesCalculation
			' Indicate that a circular reference calculation starts. 
			' If the SpreadsheetControl.Document.DocumentSettings.Calculation.Iterative property is false (default),
			' this method is not called.
			' If a circular reference calculation starts, the OnBeginCellCalculation method is called again.
			CreateLogEntry("Circular Reference calculation begins", New String() {})
			' True to perform a calculation. Return False to cancel it.
			Return True
		End Function
		Public Sub OnEndCalculation() Implements DevExpress.XtraSpreadsheet.Services.ICustomCalculationService.OnEndCalculation
			' Add a record that the calculation has finished.
			CreateLogEntry(String.Format("Calculation finishes at {0}", Date.Now), New String() {})
			rootListNode = Nothing
		End Sub
		Public Sub OnEndCellCalculation(ByVal cellKey As CellKey, ByVal startValue As CellValue, ByVal endValue As CellValue) Implements DevExpress.XtraSpreadsheet.Services.ICustomCalculationService.OnEndCellCalculation
			' Display cell information, a cell value before calculation and the calculated cell value.
			Dim info As String = String.Format("CellKey: ({0}) Before: {1}, After: {2}", cellKey, startValue, endValue)
			CreateLogEntry("Cell calculation ends", New String() {info})
		End Sub
		Public Sub OnEndCircularReferencesCalculation(ByVal cellKeys As IList(Of CellKey)) Implements DevExpress.XtraSpreadsheet.Services.ICustomCalculationService.OnEndCircularReferencesCalculation
			' Store cell keys of cells with circular references for further use.
			CircularReferencedCells = New HashSet(Of DevExpress.Spreadsheet.CellKey)(cellKeys)
			Dim sKeys(cellKeys.Count - 1) As String
			Dim i As Integer = 0
			For Each key As CellKey In cellKeys
				sKeys(i) = key.ToString()
				i += 1
			Next key
			' Display the information on cells with circular references.
			CreateLogEntry("Circular Reference calculation ends", sKeys)
		End Sub
		Public Function ShouldMarkupCalculateAlwaysCells() As Boolean Implements DevExpress.XtraSpreadsheet.Services.ICustomCalculationService.ShouldMarkupCalculateAlwaysCells
			' Mark as needing calculation the "calculate always" cells, 
			' such as cells containing volatile function or referencing another "calculate always" cell.
			Return True
		End Function

		Private Sub CreateLogEntry(ByVal actionName As String, ByVal info() As String)
			Dim firstLevelNode As TreeListNode = loggingControl.AppendNode(New Object() { actionName }, rootListNode)
			If info.Length <> 0 Then
				For i As Integer = 0 To info.Length - 1
					Dim secondLevelNode As TreeListNode = loggingControl.AppendNode(New Object() { info(i) }, firstLevelNode)
				Next i
			End If
			firstLevelNode.ExpandAll()
			loggingControl.MakeNodeVisible(firstLevelNode)
		End Sub
	End Class
	#End Region ' #TestCustomCalculationService
End Namespace
