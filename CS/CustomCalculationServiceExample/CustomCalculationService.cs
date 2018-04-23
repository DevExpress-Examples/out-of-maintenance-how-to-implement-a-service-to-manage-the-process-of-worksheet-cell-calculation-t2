#region #namespaces
using DevExpress.Spreadsheet;
using DevExpress.Spreadsheet.Formulas;
using DevExpress.XtraTreeList.Nodes;
using System;
using System.Collections.Generic;
#endregion #namespaces

namespace CustomCalculationServiceExample
{
    #region #TestCustomCalculationService
    public class TestCustomCalculationService : DevExpress.XtraSpreadsheet.Services.ICustomCalculationService {
        private DevExpress.XtraTreeList.TreeList loggingControl;
        private TreeListNode rootListNode = null;
        public HashSet<CellKey> CircularReferencedCells = new HashSet<CellKey>();

        public TestCustomCalculationService(object controlForLogging)
        {
            loggingControl = controlForLogging as DevExpress.XtraTreeList.TreeList;
        }

        public bool OnBeginCalculation() {
            // Set the root node for displaying info on the current calculation.
            rootListNode = loggingControl.AppendNode(new object[] {String.Format("Calculation starts at {0}", DateTime.Now)}, null);
            // To highlight cells with circular references, this example uses a hash set of such cells. 
            // Clear it when a new calculation starts.
            CircularReferencedCells.Clear();
            // True to perform a calculation. Return False to cancel it.
            return true;
        }
        public void OnBeginCellCalculation(CellCalculationArgs args) {
            // Add a record about the cell being calculated.
            CreateLogEntry("Cell calculation begins", new string[] {String.Format("CellKey: ({0})", args.CellKey)} );
        }
        public bool OnBeginCircularReferencesCalculation() {
            // Indicate that a circular reference calculation starts. 
            // If the SpreadsheetControl.Document.DocumentSettings.Calculation.Iterative property is false (default),
            // this method is not called.
            // If a circular reference calculation starts, the OnBeginCellCalculation method is called again.
            CreateLogEntry("Circular Reference calculation begins", new string[] {});
            // True to perform a calculation. Return False to cancel it.
            return true;
        }
        public void OnEndCalculation() {
            // Add a record that the calculation has finished.
            CreateLogEntry(String.Format("Calculation finishes at {0}", DateTime.Now), new string[] {});
            rootListNode = null;
        }
        public void OnEndCellCalculation(CellKey cellKey, CellValue startValue, CellValue endValue) {
            // Display cell information, a cell value before calculation and the calculated cell value.
            string info = String.Format("CellKey: ({0}) Before: {1}, After: {2}", cellKey, startValue, endValue);
            CreateLogEntry("Cell calculation ends", new string[] {info} );
        }
        public void OnEndCircularReferencesCalculation(IList<CellKey> cellKeys) {
            // Store cell keys of cells with circular references for further use.
            CircularReferencedCells = new HashSet<DevExpress.Spreadsheet.CellKey>(cellKeys); 
            string[] sKeys = new string[cellKeys.Count];
            int i = 0;
            foreach (CellKey key in cellKeys) {
                sKeys[i] = key.ToString();
                i++;
            } 
            // Display the information on cells with circular references.
            CreateLogEntry("Circular Reference calculation ends", sKeys);
        }
        public bool ShouldMarkupCalculateAlwaysCells() {
            // Mark as needing calculation the "calculate always" cells, 
            // such as cells containing volatile function or referencing another "calculate always" cell.
            return true;
        }

        private void CreateLogEntry(string actionName, string[] info)
        {
            TreeListNode firstLevelNode = loggingControl.AppendNode(new object[] { actionName }, rootListNode);
            if (info.Length != 0)
            {
                for (int i = 0; i < info.Length; i++)
                {
                    TreeListNode secondLevelNode = loggingControl.AppendNode(new object[] { info[i] }, firstLevelNode);
                }
            }
            firstLevelNode.ExpandAll();
            loggingControl.MakeNodeVisible(firstLevelNode);
        }
    }
    #endregion #TestCustomCalculationService
}
