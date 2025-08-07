/**
 * Applies dashboard formulas, validation rules, and supporting sheets for the
 * master equipment inventory.
 */
function setupInventoryEnhancements() {
  const ss = SpreadsheetApp.getActive();

  // Dashboard formulas
  const dashboard = ss.getSheetByName('Dashboard');
  if (dashboard) {
    const locationCounts = {
      B4: 'Building 4',
      B5: 'Store 1',
      B6: 'Store 3',
      B7: 'Store 5',
      B8: 'Yard',
      B9: 'Central Warehouse',
      B10: 'Warehouse',
      B11: 'Store 2',
      B12: 'Unknown Location'
    };

    const typeCounts = {
      D4: 'Vehicle Hoist',
      D5: 'Engine Crane',
      D6: 'Air Compressor',
      D7: 'Bendi',
      D8: 'Gas Forklift',
      D9: 'Order Picker',
      D10: 'Cat Cutter',
      D11: 'Electric Forklift',
      D12: 'Screw Compressor',
      D13: 'Loader',
      D14: 'Tire Machine',
      D15: 'Car Crusher',
      D16: 'Scissor Lift'
    };

    const statusCounts = {
      F4: 'Active',
      F5: 'Needs Repair',
      F6: 'Broken/Cant Use',
      F7: 'Retired',
      F8: 'Storage'
    };

    Object.entries(locationCounts).forEach(([cell, location]) => {
      dashboard.getRange(cell)
        .setFormula(`=COUNTIF('Equipment Data'!B:B,"${location}")`);
    });

    Object.entries(typeCounts).forEach(([cell, type]) => {
      dashboard.getRange(cell)
        .setFormula(`=COUNTIF('Equipment Data'!C:C,"${type}")`);
    });

    Object.entries(statusCounts).forEach(([cell, status]) => {
      dashboard.getRange(cell)
        .setFormula(`=COUNTIF('Equipment Data'!J:J,"${status}")`);
    });

    const metrics = {
      B21: "=COUNTA('Equipment Data'!A:A)-1",
      D21: "=COUNTIF('Equipment Data'!J:J,'Active')",
      F21: "=SUM(F5:F6)",
      B22: "=COUNTIF('Equipment Data'!I:I,'<>')",
      D22: "=COUNTA(UNIQUE('Equipment Data'!B:B))-1",
      F22: "=COUNTA(UNIQUE('Equipment Data'!C:C))-1",
      B23: "=COUNTIF('Equipment Data'!N:N,'<>')",
      D23: "=COUNTIF('Equipment Data'!P:P,'<>')"
    };

    Object.entries(metrics).forEach(([cell, formula]) => {
      dashboard.getRange(cell).setFormula(formula);
    });

    // Conditional formatting
    const rules = dashboard.getConditionalFormatRules();
    const cfRanges = [
      { range: 'F5', color: '#f44336' }, // red
      { range: 'F6', color: '#f44336' }, // red
      { range: 'F8', color: '#fff176' }, // yellow
      { range: 'B12', color: '#fff176' } // yellow
    ];
    cfRanges.forEach(({ range, color }) => {
      rules.push(
        SpreadsheetApp.newConditionalFormatRule()
          .whenNumberGreaterThan(0)
          .setBackground(color)
          .setRanges([dashboard.getRange(range)])
          .build()
      );
    });
    dashboard.setConditionalFormatRules(rules);

    // Font and alignment settings
    dashboard.getRange('1:1').setFontFamily('Arial').setFontSize(12)
      .setFontWeight('bold').setHorizontalAlignment('center');
    dashboard.getRange('3:3').setFontFamily('Arial').setFontSize(12)
      .setFontWeight('bold').setHorizontalAlignment('center');
    const dataRange = dashboard.getRange('A4:Z');
    dataRange.setFontFamily('Arial').setFontSize(10)
      .setHorizontalAlignment('left');
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
