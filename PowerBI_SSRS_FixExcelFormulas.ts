// Fix Excel formulas in SSRS exports (which can be added in text with a prefix)
// Translate some Dutch Translations back to English, so they will work with the Excel Backend

class PBIFormulaFixer {
  private _workbook: ExcelScript.Workbook;
  private sheetname: string;
  private _worksheet: ExcelScript.Worksheet;
  private sheetvalues: (string | number | boolean)[][];
  private numberFormat: string = "#.### ;-#.###;-";
  private formulaColLbl: string = "FORMULA";
  private formulaPrefix: string = "FORMULA_";
  private formulaCols: number[] = [];
  private lastFormCol: number;
  private firstDataRow: number;
  private lastDataRow: number;

  public constructor(wrkbook: ExcelScript.Workbook, wrksheet: string = "Formatieberekening") {
    // Init Object
    this._workbook = wrkbook;
    this.sheetname = wrksheet;
  }
  public setSheet() {
    // Setup and check Excel Worksheet
    this._worksheet = this._workbook.getWorksheet(this.sheetname);
    this.sheetvalues = this._worksheet.getUsedRange().getValues();
  }
  private initFixer() {
    // Setup all variables for Fixer
    this.setSheet();
    if (this._worksheet != undefined) {
      this.findFormulaCols();
      console.log("formula cols: " + this.formulaCols.toString());
      this.lastFormCol = this.formulaCols[this.formulaCols.length - 1];
      console.log("Last col: " + this.lastFormCol.toString());
      this.findFirstDataRow();
      console.log("First Data row: " + this.firstDataRow.toString());
      this.findLastDataRow();
      console.log("Last Data row: " + this.lastDataRow.toString());
    } else {
      console.log("ERROR: " + this.sheetname + " does not exists!");
    }
  }
  public runFixer() {
    // Fix all formulas in Worksheet
    console.log("Setting up the Fixer Settings");
    this.initFixer();
    if (this._worksheet != undefined) {
      console.log("Fixing this shit!");
      this.fixR1C1Formulas();
      console.log("And.... Done!");
    } else {
      console.log("ERROR: Fixer could not be completed!");
    }
  }

  private findFormulaCols() {
    //
    let maxRow = 3;
    let maxCol = 30;
    for (let row = 1; row <= maxRow; row++) {
      for (let col = 1; col <= maxCol; col++) {
        let cell = this._worksheet.getCell(row, col);
        if (this.sheetvalues[row - 1][col - 1] == this.formulaColLbl) {
          cell.getEntireColumn().setNumberFormatLocal(this.numberFormat);
          this.formulaCols.push(cell.getColumnIndex());
          cell.setValue("");
        }
      }
    }
  }
  private findFirstDataRow() {
    let maxRow = 5;
    let maxCol = 5;
    for (let row = 1; row <= maxRow; row++) {
      for (let col = 1; col <= maxCol; col++) {
        if (this.sheetvalues[row - 1][col - 1] == "College") {
          this.firstDataRow = this._worksheet.getCell(row + 2, col).getRowIndex();
          return;
        }
      }
    }
  }
  private findLastDataRow() {
    let maxRow = 500;
    for (let row = this.firstDataRow; row < maxRow; row++) {
      for (let col = this.lastFormCol; col < this.lastFormCol + 1; col++) {
        let cell = this._worksheet.getCell(row, col);
        if (cell.getValueType() != ExcelScript.RangeValueType.string) {
          let val = cell.getValue().toString();
          if (val == "") {
            this.lastDataRow = this._worksheet.getCell(row, col).getRowIndex();
            return;
          }
        }
      }
    }
  }
  private translateFormula(formula: string): string {
    let translated = formula;
    const regexp = new RegExp("RK", "g");
    translated = translated.replace(RegExp("AFRONDEN", "g"), "ROUND");
    translated = translated.replace(RegExp("SOM", "g"), "SUM");
    translated = translated.replace(regexp, "RC");
    return translated;
  }

  private fixFormula(formula: string): string {
    let formula_fixed = formula;
    formula_fixed = this.translateFormula(formula_fixed);
    formula_fixed = formula_fixed.replace(",", ".");
    formula_fixed = formula_fixed.replace(";", ",");
    formula_fixed = formula_fixed.replace(this.formulaPrefix, "=");
    return formula_fixed;
  }

  private fixR1C1Formulas() {
    for (var col of this.formulaCols) {
      for (let row = this.firstDataRow; row < this.lastDataRow + 1; row++) {
        let cell = this._worksheet.getCell(row, col);
        if (cell.getValueType() == ExcelScript.RangeValueType.string) {
          //   let val = cell.getValue().toString();
          let val = this.sheetvalues[row - 1][col - 1].toString();
          if (val.substr(0, 8) == this.formulaPrefix) {
            cell.setValue(0);
            let formula = this.fixFormula(val);
            cell.setFormulaR1C1(formula);
          }
        }
      }
    }
  }
  //   End of Object
}
function main(workbook: ExcelScript.Workbook, sheet: string = "Sheet1") {
  const fixer = new PBIFormulaFixer(workbook, sheet);
  fixer.runFixer();
}
