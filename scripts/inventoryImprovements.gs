function setupInventoryEnhancements() {
  const ss = SpreadsheetApp.getActive();

  // Dashboard formulas
  const dashboard = ss.getSheetByName('Dashboard');
  if (dashboard) {
    // Equipment by Location
    dashboard.getRange('B4').setFormula("=COUNTIF('Equipment Data'!B:B,'Building 4')");
    dashboard.getRange('B5').setFormula("=COUNTIF('Equipment Data'!B:B,'Store 1')");
    dashboard.getRange('B6').setFormula("=COUNTIF('Equipment Data'!B:B,'Store 3')");
    dashboard.getRange('B7').setFormula("=COUNTIF('Equipment Data'!B:B,'Store 5')");
    dashboard.getRange('B8').setFormula("=COUNTIF('Equipment Data'!B:B,'Yard')");
    dashboard.getRange('B9').setFormula("=COUNTIF('Equipment Data'!B:B,'Central Warehouse')");
    dashboard.getRange('B10').setFormula("=COUNTIF('Equipment Data'!B:B,'Warehouse')");
    dashboard.getRange('B11').setFormula("=COUNTIF('Equipment Data'!B:B,'Store 2')");
    dashboard.getRange('B12').setFormula("=COUNTIF('Equipment Data'!B:B,'Unknown Location')");

    // Equipment by Type
    dashboard.getRange('D4').setFormula("=COUNTIF('Equipment Data'!C:C,'Vehicle Hoist')");
    dashboard.getRange('D5').setFormula("=COUNTIF('Equipment Data'!C:C,'Engine Crane')");
    dashboard.getRange('D6').setFormula("=COUNTIF('Equipment Data'!C:C,'Air Compressor')");
    dashboard.getRange('D7').setFormula("=COUNTIF('Equipment Data'!C:C,'Bendi')");
    dashboard.getRange('D8').setFormula("=COUNTIF('Equipment Data'!C:C,'Gas Forklift')");
    dashboard.getRange('D9').setFormula("=COUNTIF('Equipment Data'!C:C,'Order Picker')");
    dashboard.getRange('D10').setFormula("=COUNTIF('Equipment Data'!C:C,'Cat Cutter')");
    dashboard.getRange('D11').setFormula("=COUNTIF('Equipment Data'!C:C,'Electric Forklift')");
    dashboard.getRange('D12').setFormula("=COUNTIF('Equipment Data'!C:C,'Screw Compressor')");
    dashboard.getRange('D13').setFormula("=COUNTIF('Equipment Data'!C:C,'Loader')");
    dashboard.getRange('D14').setFormula("=COUNTIF('Equipment Data'!C:C,'Tire Machine')");
    dashboard.getRange('D15').setFormula("=COUNTIF('Equipment Data'!C:C,'Car Crusher')");
    dashboard.getRange('D16').setFormula("=COUNTIF('Equipment Data'!C:C,'Scissor Lift')");

    // Equipment by Status
    dashboard.getRange('F4').setFormula("=COUNTIF('Equipment Data'!J:J,'Active')");
    dashboard.getRange('F5').setFormula("=COUNTIF('Equipment Data'!J:J,'Needs Repair')");
    dashboard.getRange('F6').setFormula("=COUNTIF('Equipment Data'!J:J,'Broken/Cant Use')");
    dashboard.getRange('F7').setFormula("=COUNTIF('Equipment Data'!J:J,'Retired')");
    dashboard.getRange('F8').setFormula("=COUNTIF('Equipment Data'!J:J,'Storage')");

    // Key Metrics
    dashboard.getRange('B21').setFormula("=COUNTA('Equipment Data'!A:A)-1");
    dashboard.getRange('D21').setFormula("=COUNTIF('Equipment Data'!J:J,'Active')");
    dashboard.getRange('F21').setFormula("=SUM(F5:F6)");
    dashboard.getRange('B22').setFormula("=COUNTIF('Equipment Data'!I:I,'<>')");
    dashboard.getRange('D22').setFormula("=COUNTA(UNIQUE('Equipment Data'!B:B))-1");
    dashboard.getRange('F22').setFormula("=COUNTA(UNIQUE('Equipment Data'!C:C))-1");
    dashboard.getRange('B23').setFormula("=COUNTIF('Equipment Data'!N:N,'<>')");

    // QR count
    dashboard.getRange('D23').setFormula("=COUNTIF('Equipment Data'!P:P,'<>')");
  }

  // Data validation rules
  const equipmentData = ss.getSheetByName('Equipment Data');
  if (equipmentData) {
    const locations = ['Store 1','Store 2','Store 3','Store 5','Building 4','Yard','Central Warehouse','Warehouse','Unknown Location'];
    const equipmentTypes = ['Gas Forklift','Loader','Bendi','Order Picker','Electric Forklift','Scissor Lift','Electric Pallet Jack','Air Compressor','Screw Compressor','Vehicle Hoist','Engine Crane','Tire Machine','Car Crusher','Dock Truck','Box Truck','Flatbed Truck','Ramp','Cat Cutter'];
    const statuses = ['Active','Needs Repair','Broken/Cant Use','Storage','Retired'];
    const employees = ['Juan','Mike','Carlos','Troy'];

    equipmentData.getRange('B2:B1000').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(locations).setAllowInvalid(false).build());
    equipmentData.getRange('C2:C1000').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(equipmentTypes).setAllowInvalid(false).build());
    equipmentData.getRange('J2:J1000').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(statuses).setAllowInvalid(false).build());
    equipmentData.getRange('N2:N1000').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(employees).setAllowInvalid(false).build());
    equipmentData.getRange('A2:A1000').setDataValidation(SpreadsheetApp.newDataValidation().requireFormulaSatisfied("=REGEXMATCH(A2,'^EQ-[0-9]{3}$')").setAllowInvalid(false).setHelpText('Enter ID in format EQ-XXX (e.g., EQ-000)').build());

    // QR URL formula
    equipmentData.getRange('P2:P1000').setFormulaR1C1("=IF(RC[-15]<>'','https://forms.gle/abco_maintenance?equip_id='&RC[-15],'')");
  }

  // Maintenance sheet formulas
  const maintenance = ss.getSheetByName('Maintenance');
  if (maintenance) {
    maintenance.getRange('L2').setFormula("=COUNTIF(C:C,C2)");
    maintenance.getRange('M2').setFormula("=SUMIF(C:C,C2,I:I)");
  }

  // Repair summary formulas
  const repairSummary = ss.getSheetByName('Repair Summary');
  if (repairSummary) {
    const types = ['Tire Machine','Vehicle Hoist','Engine Crane','Air Compressor','Bendi','Gas Forklift','Order Picker','Cat Cutter','Electric Forklift','Screw Compressor','Loader','Car Crusher','Scissor Lift'];
    let row = 3;
    types.forEach(function(type) {
      repairSummary.getRange(row,3).setFormula("=SUMIF('Maintenance'!E:E,'"+type+"','Maintenance'!L:L)");
      repairSummary.getRange(row,4).setFormula("=SUMIF('Maintenance'!E:E,'"+type+"','Maintenance'!M:M)");
      row++;
    });
  }

  // Efficiency Audit sheet
  if (!ss.getSheetByName('Efficiency Audit')) {
    const sheet = ss.insertSheet('Efficiency Audit');
    sheet.getRange('A1:F1').setValues([[
      'Equipment ID','Equipment Type','Location','Last Used','Downtime Hours','Assigned Employee'
    ]]);
    sheet.getRange('E2:E').setFormula("=IF(F2='Active',0,NETWORKDAYS(D2,TODAY())*8)");
  }
}
