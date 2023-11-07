# MS-Excel

<details>
<summary> Initialize Excel </summary>

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
  
var	ExApp: Variant;
	WB, WS: Variant;
	 
 begin
 
	// CoInitialize(nil); // to activate COM
 
	ExApp := CreateOleObject('Excel.Application');
	ExApp.Visible := True;

	WB := ExApp.Workbooks.Open(ExtractFilePath(ParamStr(0)) + 'Excel-File.xlsx');
	WS := WB.WorkSheets[SheetNum];
	
	...
	
	...

	WB.Close;
	
	WB := Unassigned;
	WS := Unassigned;	
	
	ExApp.Quit;
	ExApp := Unassigned;
	
	// CoUninitialize; // deactivate COM
 
 end;
	
```
</details>

<details>
<summary> Used Range </summary>

```delphi
	WS.Activate;

	WS.Cells.SpecialCells(xlCellTypeLastCell).Activate;

	// Used Range
	Rows := ExApp.ActiveCell.Row;
	Cols := ExApp.ActiveCell.Column;

```
</details>

<details>
<summary> Cells </summary>
```delphi
	// Write into Cell
	WS.Cells[RowNum, ColNum] := AnyVariantValue;
	
	// Read from Cell
	AnyVariantValue := WS.Cells[RowNum, ColNum];
```
</details>

