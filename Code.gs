// Global declarations
const maintenanceMode = false;
const debugLog = true;

// Global delarations for sheets and specific columns
const sheetAisles         = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Aisles");
const allAisles           = sheetAisles.getRange(2,1,sheetAisles.getLastRow()-1,3).getValues();
const sheetBrands         = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Brands");
const allBrands           = sheetBrands.getRange(2,1,sheetBrands.getLastRow()-1,3).getValues();

function onEdit(e){
  if(maintenanceMode != true){
    var activeCell = e.range;
    var v = activeCell.getValue();
    var r = activeCell.getRow();
    var c = activeCell.getColumn();
    var activeSheetName = activeCell.getSheet().getName();
    const activeSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(activeSheetName);

    if(debugLog === true){
      Logger.log("activeCell: "+activeCell);
      Logger.log("v: "+v);
      Logger.log("r: "+r);
      Logger.log("c: "+c);
      Logger.log("activeSheetName: "+activeSheetName);
      Logger.log("activeSheet: "+activeSheet);
    } // end of: if(debugLog === true)

    // Was it the q_ProductHistory sheet which was changed?
    if(activeSheetName === "Aisles"){
      
      // Was it the value in the "MasterAisleSelected" cell which was changed?
      if(c === 13 && r === 2){
        for(i=r+1;i<=8;i++){
          activeSheet.getRange(i,c).clearContent();
        } // end of: for(i=r+1;i<=8;i++)
      } // end of: if(c === 13 && r === 2)

      // Was it the value in the "SecondaryAisleSelected" cell which was changed?
      if(c === 13 && r === 3){
        for(i=r+1;i<=8;i++){
          activeSheet.getRange(i,c).clearContent();
        } // end of: for(i=r+1;i<=7;i++)
      } // end of: if(c === 13 && r === 3)

      // Was it the value in the "TertiaryAisleSelected" cell which was changed?
      if(c === 13 && r === 4){
        for(i=r+1;i<=8;i++){
          activeSheet.getRange(i,c).clearContent();
        } // end of: for(i=r+1;i<=6;i++)
      } // end of: if(c === 13 && r === 4)

      // Now set the BrandSelected cell value's data validation.

      //Step 1: Clear the current data validation.
      activeSheet.getRange("BrandSelected").setDataValidation(null);

      // Step 2: Check the first value of each of the filtered brands, in order tertiary to master.
      var myBrandsFilteredByTertiary = activeSheet.getRange("isemptyBrandsFilteredByTertiary").getValue();
      var myBrandsFilteredBySecondary = activeSheet.getRange("isemptyBrandsFilteredBySecondary").getValue();
      var myBrandsFilteredByMaster = activeSheet.getRange("isemptyBrandsFilteredByMaster").getValue();

      if(debugLog === true){
        Logger.log("myBrandsFilteredByTertiary: "+myBrandsFilteredByTertiary);
        Logger.log("myBrandsFilteredBySecondary: "+myBrandsFilteredBySecondary);
        Logger.log("myBrandsFilteredByMaster: "+myBrandsFilteredByMaster);
      }// end of: if(debugLog === true)

      /* POSSIBLE NEW SOLUTION, BY COPYING THE VALIDATION RULE ONLY */
      if(myBrandsFilteredByTertiary === ''){
        myValidation = activeSheet.getRange("BrandSelectedByTertiary").getDataValidation().copy();
        Logger.log("myValidation: "+myValidation);
        activeSheet.getRange("BrandSelected").setDataValidation(myValidation);
      } else if(myBrandsFilteredBySecondary === ''){
        myValidation = activeSheet.getRange("BrandSelectedBySecondary").getDataValidation().copy();
        Logger.log("myValidation: "+myValidation);
        activeSheet.getRange("BrandSelected").setDataValidation(myValidation);
      } else  if(myBrandsFilteredByMaster === ''){
        myValidation = activeSheet.getRange("BrandSelectedByMaster").getDataValidation().copy();
        Logger.log("myValidation: "+myValidation);
        activeSheet.getRange("BrandSelected").setDataValidation(myValidation);

      } // end of: if (isemptyBrandsFilteredByTertiary === "") else



      /* TEMPORARILY REMARK IT ALL OUT

      // Step 3: Populate rangeToApply using values from the filtered brands, in order tertiary to master.
      if(myBrandsFilteredByTertiary === ""){
        var rangeToApply = activeSheet.getRange("BrandsFilteredByTertiary");
      } else if(myBrandsFilteredBySecondary === ""){
        var rangeToApply = activeSheet.getRange("BrandsFilteredBySecondary");
      } else  if(myBrandsFilteredByMaster === ""){
        var rangeToApply = activeSheet.getRange("BrandsFilteredByMaster");
      } // end of: if (myBrandsFilteredByTertiary === "") else

      if(debugLog === true){
        Logger.log("rangeToApply: "+rangeToApply);
      }// end of: if(debugLog === true)

      // Step 4: Set the other required fields for the data validation call.
      var rule = activeSheet.getRange('BrandSelected').getDataValidation();
      var listToApply = [""];
      // var rangeToApply = null;
      // var listToApply = activeSheet.getRange("O2:O3").getValues();
      var thisCell = activeSheet.getRange('BrandSelected');
      var invalidsPolicy = false;

      if(debugLog === true){
        Logger.log("rule: "+rule);
        Logger.log("listToApply: "+listToApply);
        Logger.log("rangeToApply: "+rangeToApply);
        Logger.log("thisCell: "+thisCell);
        Logger.log("invalidsPolicy: "+invalidsPolicy);
      }// end of: if(debugLog === true)

      // Step 5: Call the data validation setting function.
      applyValidationToCell(listToApply,rangeToApply,thisCell,invalidsPolicy);
      */

    } // end of: if(activeSheetName === "Aisles")

  } // end of: if(maintenanceMode != true)

} // end of: function onEdit(e)

/**
 * @desc Applies a validation rule to a cell, either from a list or a range, and either strictly or permissably.
 * @author FrittRo
 * @todo Conforming code to the Google JavaScript Style Guide. https://git.io/Jcqk2
 * @todo Conforming to Google Apps Script Best Practices. https://git.io/Jcqk1
 * 
 * @param {Array} listToApply [OPTION 1] A list of items to be displayed
 * @param {Object} rangeToApply [OPTION 2] A range containing the items to be displayed 
 * @param {Object} thisCell A cell reference to contain the dropdown in a given sheet
 * @param {Boolean} invalidsPolicy Whether or not to allow invalid selections in the cell
 */
function applyValidationToCell(listToApply,rangeToApply,thisCell,invalidsPolicy) {
  if (rangeToApply === null){
    var rule = SpreadsheetApp
    .newDataValidation()
    .requireValueInList(listToApply)
    .setAllowInvalid(invalidsPolicy)
    .build();
  } else {
    var rule = SpreadsheetApp
    .newDataValidation()
    .requireValueInRange(rangeToApply,true)
    .setAllowInvalid(invalidsPolicy)
    .build();
  }
  thisCell.setDataValidation(rule);
  SpreadsheetApp.flush();
}