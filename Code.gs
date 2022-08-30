/**
 * Groceries since 1-Sep-2022
 * 
 * This spreadsheet contains our grocery purchasing data, collected from the orders and invoices
 * paperwork which we receive via email from a "supplier" such as  PAK'nSAVE, when we place our
 * weekly grocery requests, and receive our weekly purchases. The terminology used here is specific
 * and intentional. Orders relate to "requests", and invoices relate to "purchases". There are also
 * "products", which are stock items that are available for purchase from a supplier. In future
 * additions to this spreadsheet, we will also define those purchases which we have in stock at home,
 * as being "items". An item will therefore be a product which has been requested and purchased, and
 * is now available for consumption within our home. This will allow us to include these items in a
 * "recipe", which in turn can be consumed as part of a "meal", which will be part of a "mealplan".
 * Once an item has been consumed, whether as a food item it is eaten, or as a non-food item it is
 * used up, it will be removed from our items list and disposed of in the "consumed" list. In the
 * longer-term, another advance on this spreadsheet will allow for "nutrients" to be tracked, such
 * as the sugar, salt, and carbohydrate content of a food item, allowing for a measure of our
 * nutrient intake. But that is a way off in the future yet.
 * 
 * @author FrittRo on {@link https://github.com/frittro|GitHub}.
 * @version 0.1
 * @copyright Robert Frittmann 2022
 * @license CC-BY-4.0
 */

  /**
   * Global declarations
   */ 
    const maintenanceMode = false; // globally disable all scripts if true.
    const debugLog = true; // globally enable debug logging if true.
    if(maintenanceMode === true){
      Logger.log("MAINTENANCE MODE: all scripts are disabled");
    }



  /**
   * Global delarations for sheets
   */
    const sheetSuppliers        = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("t_Suppliers");
    const sheetAisles           = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("t_Aisles");
    const sheetBrands           = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("t_Brands");
    const sheetProducts         = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("t_Products");
    const sheetOrderMetadata    = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("t_OrderMetadata");
    const sheetRequests         = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("t_Requests");
    const sheetInvoiceMetadata  = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("t_InvoiceMetadata");
    const sheetPurchases        = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("t_Purchases");



/**
 * The event handler triggered when opening the spreadsheet.
 * @param {Event} e The onOpen event.
 * @see https://developers.google.com/apps-script/guides/triggers#onopene
 */
function onOpen(e) {
  // var maintenanceMode = true;  // local override to disable this script only
  if(maintenanceMode = false){
    // var debugLog = false; // local override
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
    }
  }
}



/**
 * The event handler triggered when the selection changes in the spreadsheet.
 * @param {Event} e The onSelectionChange event.
 * @see https://developers.google.com/apps-script/guides/triggers#onselectionchangee
 */
function onSelectionChange(e) {
  // var maintenanceMode = true;  // local override to disable this script only
  if(maintenanceMode = false){
    // var debugLog = false; // local override
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
  } // end of: if(maintenanceMode = false)
} // end of: function onSelectionChange(e)



/**
 * The event handler triggered when editing the spreadsheet.
 * @param {Event} e The onEdit event.
 * @see https://developers.google.com/apps-script/guides/triggers#onedite
 */
function onEdit(e){
  // var maintenanceMode = true;  // local override to disable this script only
  if(maintenanceMode = false){
    // var debugLog = false; // local override
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

    // Was it the t_Aisles sheet which was changed?
    if(activeSheetName === "t_Aisles"){
      
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

      // Now set the data validation on the BrandSelected cell.

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
    }
  }
}



/**
 * Applies a validation rule to a cell, either from a list or a range, and either strictly or permissably.
 * @author FrittRo on {@link https://github.com/frittro|GitHub}.
 * @version 0.1
 * @copyright Robert Frittmann 2022
 * @license CC-BY-4.0
 * 
 * @todo Conforming code to the Google JavaScript Style Guide. https://git.io/Jcqk2
 * @todo Conforming to Google Apps Script Best Practices. https://git.io/Jcqk1
 * 
 * @param {Array} listToApply [OPTION 1] A list of items to be displayed
 * @param {Object} rangeToApply [OPTION 2] A range containing the items to be displayed 
 * @param {Object} thisCell A cell reference to contain the dropdown in a given sheet
 * @param {Boolean} invalidsPolicy Whether or not to allow invalid selections in the cell
 */
function applyValidationToCell(listToApply = ["no items"],rangeToApply,thisCell,invalidsPolicy) {
  // var maintenanceMode = true;  // local override to disable this script only
  if(maintenanceMode = false){
    // var debugLog = false; // local override
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
}



/**
 * Function to get a column number by its header name. Assumes that all sheets follow the same pattern
 * of only the first row being the single header row.
 * 
 * @author FrittRo on {@link https://github.com/frittro|GitHub}.
 * @version 1.0
 * @copyright Robert Frittmann 2022
 * @license CC-BY-4.0
 * @see {@link https://stackoverflow.com/a/72419394/19709446|StackOverflow} for the inspiration.
 * 
 * @param {String} sheetName The name of the sheet to search in.
 * @param {String} headerName The name of the column to search for.
 * @return An integer value for the column number found by its name.
 */
function getColumnNumberByHeaderName(sheetName,headerName){
  // var maintenanceMode = true;  // local override to disable this script only
  if(maintenanceMode = false){
    // var debugLog = false; // local override
    const sheet = ss.getSheetByName(sheetName);
    return parseInt(sheet.getRange('1:1').getValues()[0].indexOf(headerName) + 1);
  }
}



/**
 * Function to copy all requested items by ProductID, for a specified date, to the purchases sheet.
 * 
 * @author FrittRo on {@link https://github.com/frittro|GitHub}.
 * @version 0.1
 * @copyright Robert Frittmann 2022
 * @license CC-BY-4.0
 * 
 * 
 */
function copyOrderedItemsToPurchases(){

}
