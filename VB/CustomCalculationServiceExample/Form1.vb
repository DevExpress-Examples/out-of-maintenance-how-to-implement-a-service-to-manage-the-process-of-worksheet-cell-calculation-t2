Imports DevExpress.Spreadsheet
Imports DevExpress.XtraSpreadsheet.Services
Imports System
Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Data
Imports System.Drawing
Imports System.IO
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks
Imports System.Windows.Forms

Namespace CustomCalculationServiceExample
	Partial Public Class Form1
		Inherits DevExpress.XtraBars.Ribbon.RibbonForm

		Private calculationService As TestCustomCalculationService

		Public Sub New()
			InitializeComponent()
			spreadsheetControl1.BeginUpdate()
			spreadsheetControl1.Document.DocumentSettings.Calculation.Iterative = True
			calculationService = New TestCustomCalculationService(loggingControl)
			spreadsheetControl1.AddService(GetType(ICustomCalculationService), calculationService)
			AddHandler spreadsheetControl1.CustomDrawCell, AddressOf spreadsheetControl1_CustomDrawCell
			spreadsheetControl1.ActiveWorksheet.Cells("A3").Formula = "A3*B3"
			spreadsheetControl1.ActiveWorksheet.Cells("B3").Formula = "8*9"
			spreadsheetControl1.EndUpdate()
		End Sub

		Private Sub spreadsheetControl1_CustomDrawCell(ByVal sender As Object, ByVal e As DevExpress.XtraSpreadsheet.CustomDrawCellEventArgs)
			Dim cellkey As New CellKey((e.Cell.Worksheet.Index + 1), e.Cell.ColumnIndex, e.Cell.RowIndex)
			If calculationService.CircularReferencedCells.Contains(cellkey) Then
				e.Graphics.DrawRectangle(New Pen(Color.Red, 1), e.Bounds)
			End If
		End Sub

		Private Sub clearButtonItem_ItemClick(ByVal sender As Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles clearButtonItem.ItemClick
			loggingControl.ClearNodes()
		End Sub

		Private Sub expandButtonItem_ItemClick(ByVal sender As Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles expandButtonItem.ItemClick
			loggingControl.ExpandAll()
		End Sub

		Private Sub collapseButtonItem_ItemClick(ByVal sender As Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles collapseButtonItem.ItemClick
			loggingControl.CollapseAll()
		End Sub
	End Class
End Namespace
