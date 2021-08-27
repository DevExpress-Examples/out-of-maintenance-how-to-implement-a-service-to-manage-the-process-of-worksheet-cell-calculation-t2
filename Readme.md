<!-- default badges list -->
![](https://img.shields.io/endpoint?url=https://codecentral.devexpress.com/api/v1/VersionRange/128613660/15.1.5%2B)
[![](https://img.shields.io/badge/Open_in_DevExpress_Support_Center-FF7200?style=flat-square&logo=DevExpress&logoColor=white)](https://supportcenter.devexpress.com/ticket/details/T270403)
[![](https://img.shields.io/badge/📖_How_to_use_DevExpress_Examples-e9f6fc?style=flat-square)](https://docs.devexpress.com/GeneralInformation/403183)
<!-- default badges end -->
<!-- default file list -->
*Files to look at*:

* [CustomCalculationService.cs](./CS/CustomCalculationServiceExample/CustomCalculationService.cs) (VB: [CustomCalculationService.vb](./VB/CustomCalculationServiceExample/CustomCalculationService.vb))
* [Form1.cs](./CS/CustomCalculationServiceExample/Form1.cs) (VB: [Form1.vb](./VB/CustomCalculationServiceExample/Form1.vb))
<!-- default file list end -->
# How to implement a service to manage the process of worksheet cell calculation


This example illustrates the use of the service which implements the <a href="http://help.devexpress.com/#CoreLibraries/clsDevExpressXtraSpreadsheetServicesICustomCalculationServicetopic">DevExpress.XtraSpreadsheet.Services.ICustomCalculationService</a>  interface and allows to manage the process of worksheet calculations.<br />The application creates log entries when worksheet calculation starts and finishes, when cell calculation begins and ends, when a cell with the circular reference starts and finishes its calculation.<br />Cells containing circular reference are indicated by drawing a red box within the cell.<br /><br /><img src="https://raw.githubusercontent.com/DevExpress-Examples/how-to-implement-a-service-to-manage-the-process-of-worksheet-cell-calculation-t270403/15.1.5+/media/9fc8a8af-314d-11e5-80bf-00155d62480c.png"><br /><br />

<br/>


