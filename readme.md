# MS-Excel

<details>
<summary> First Steps </summary>

```delphi

uses
	...
  ComObj,
  ActiveX,
  Variants,
  
  Excel2010; // optional...
  
  ...
  
const
	xlCellTypeLastCell = $0000000B;
  
 var ExApp: Variant;
	 WB, WS: Variant;
 begin
 
	// CoInitialize(nil); // to activate COM
 
	ExApp := CreateOleObject('Excel.Application');
	ExApp.Visible := True;

	WB := ExApp.Workbooks.Open(ExtractFilePath(ParamStr(0)) + 'Mappe1.xlsx');
	WS := WB.WorkSheets[SheetNum];
	WS.Activate;

	WS.Cells.SpecialCells(xlCellTypeLastCell).Activate;

	// Used Range
	Rows := ExApp.ActiveCell.Row;
	Cols := ExApp.ActiveCell.Column;

	WB.Close;
	EX.Quit;
	EX := Unassigned;
	WB := Unassigned;
	WS := Unassigned;
	
	
	// CoUninitialize; // deactivate
 
 end;
	
```
</details>