function calculateTargetActualizationScore() {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var salesSheet = spreadsheet.getSheetByName('sales');
    var salesTargetSheet = spreadsheet.getSheetByName('sales_target');
    var stockJanSheet = spreadsheet.getSheetByName('stock_jan');
    var stockFebSheet = spreadsheet.getSheetByName('stock_feb');
    var stockMarchSheet = spreadsheet.getSheetByName('stock_march');
    var sCJanSheet = spreadsheet.getSheetByName('stock_connection_jan');
    var sCFebSheet = spreadsheet.getSheetByName('stock_connection_feb');
    var sCMarchSheet = spreadsheet.getSheetByName('stock_connection_march');
    var scoresSheet = spreadsheet.getSheetByName('scores');
  
    if (!scoresSheet) {
      scoresSheet = spreadsheet.insertSheet('scores');
    } else {
      scoresSheet.clear();
    }
  
    var months = ['Jan', 'Feb', 'March'];
    var scoresData = [['DMSFirmCode', 'DMSFirmName', 'Jan', 'Feb', 'March' , 'Target Actualization Rate', 'Score', 'Total Stock Number']];
    var infoData = [['Month', 'TurkeyTargetActualizationRate']];
  
    var turkeyTargetActualizationRate;
    var firmInfo = {}
    var month;
  
    for (var i = 0; i < months.length; i++) {
      month = months[i];
      
      var salesRange = salesSheet.getRange(2, i + 3, salesSheet.getLastRow()-1 , 1);
      var targetRange = salesTargetSheet.getRange(2, i + 3, salesTargetSheet.getLastRow()-1, 1);
      var salesValues = salesRange.getValues().flat();
      var targetValues = targetRange.getValues().flat();
  
      var sJan = stockJanSheet.getRange(2, 6, stockJanSheet.getLastRow() - 1, 1);
      var sFeb = stockFebSheet.getRange(2, 6, stockFebSheet.getLastRow() - 1, 1);
      var sMarch = stockMarchSheet.getRange(2, 6, stockMarchSheet.getLastRow() - 1, 1);
  
      var sJanFirm = stockJanSheet.getRange(2, 2, stockJanSheet.getLastRow() - 1, 1);
      var sFebFirm = stockFebSheet.getRange(2, 2, stockFebSheet.getLastRow() - 1, 1);
      var sMarchFirm = stockMarchSheet.getRange(2, 2, stockMarchSheet.getLastRow() - 1, 1);
  
    
      var sCJanRange = sCJanSheet.getRange(2, 6, sCJanSheet.getLastRow() - 1, 2);
      var sCFebRange = sCFebSheet.getRange(2, 6, sCFebSheet.getLastRow() - 1, 2);
      var sCMarchRange = sCMarchSheet.getRange(2, 6, sCMarchSheet.getLastRow() - 1, 2);
  
      var sCJanRangeFirm = sCJanSheet.getRange(2, 2, sCJanSheet.getLastRow() - 1, 2);
      var sCFebRangeFirm = sCFebSheet.getRange(2, 2, sCFebSheet.getLastRow() - 1, 2);
      var sCMarchRangeFirm = sCMarchSheet.getRange(2, 2, sCMarchSheet.getLastRow() - 1, 2);
  
      var stockJanValues = sJan.getValues().flat();
      var stockJanFirmValues = sJanFirm.getValues().flat();
      var stockFebValues = sFeb.getValues().flat();
      var stockFebFirmValues = sFebFirm.getValues().flat();
      var stockMarchValues = sMarch.getValues().flat();
      var stockMarchFirmValues = sMarchFirm.getValues().flat();
  
      var monthlyStocks = [stockJanValues, stockFebValues, stockMarchValues]
      var monthlyStockNames = [stockJanFirmValues,stockFebFirmValues,stockMarchFirmValues]
  
  
      var sCJanValues = sCJanRange.getValues().flat();
      var sCFebValues = sCFebRange.getValues().flat();
      var sCMarchValues = sCMarchRange.getValues().flat();
  
      var sCJanFirmValues = sCJanRangeFirm.getValues().flat();
      var sCFebFirmValues = sCFebRangeFirm.getValues().flat();
      var sCMarchFirmValues = sCMarchRangeFirm.getValues().flat();
  
      var monthlySCs = [sCJanValues, sCFebValues, sCMarchValues]
      var monthlySCNames = [sCJanFirmValues,sCFebFirmValues,sCMarchFirmValues]
  
      var totalSales = salesValues.reduce((a, b) => a + b, 0);
      var totalTarget = targetValues.reduce((a, b) => a + b, 0);
      turkeyTargetActualizationRate = (totalSales / totalTarget) * 100;
      var minTurkeyTargetRate = turkeyTargetActualizationRate - 10;
      var maxTurkeyTargetRate = turkeyTargetActualizationRate + 10;
      var minScore = 0;
      var maxScore = 15;
      var targetActualizationRate
      var firmCode
      var firmName
      var totalStockNumber 
  
      for (var j = 0; j < salesValues.length; j++) {
        targetActualizationRate = (salesValues[j] / targetValues[j]) * 100;
        firmCode = salesSheet.getRange(j + 2, 2).getValue();
        firmName = salesSheet.getRange(j + 2, 1).getValue();
        totalStockNumber = calculateTotalStockNumber(firmCode,  monthlyStocks[i], monthlyStockNames[i] ,monthlySCs[i], monthlySCNames[i]);
  
        if (targetActualizationRate == turkeyTargetActualizationRate) {
          score = 7.5;
          if (totalStockNumber == 0){
            score = 15;
          }
        } else if (targetActualizationRate >= 1.10 * turkeyTargetActualizationRate || totalStockNumber == 0) {
          score = 15;
        } else if (targetActualizationRate >= (turkeyTargetActualizationRate - 10) && targetActualizationRate < turkeyTargetActualizationRate) {
          var ratio = (targetActualizationRate - minTurkeyTargetRate) / (maxTurkeyTargetRate - minTurkeyTargetRate);
          score = minScore + ratio * (maxScore - minScore);
          score = score.toFixed(2);
        } else if (targetActualizationRate <= (turkeyTargetActualizationRate + 10) && targetActualizationRate > turkeyTargetActualizationRate) {
          var ratio = (targetActualizationRate - minTurkeyTargetRate) / (maxTurkeyTargetRate - minTurkeyTargetRate);
          score = minScore + ratio * (maxScore - minScore);
          score = score.toFixed(2);
        } else {
          score = 0;
        }
        if (!firmInfo[firmCode]) {
          firmInfo[firmCode] = {};
        }
  
        if (!firmInfo[firmCode][month]) {
          firmInfo[firmCode][month] = {};
        }
  
  
        firmInfo[firmCode][month] = {"firmName": firmName, "targetActualizationRate": targetActualizationRate, "totalStockNumber": totalStockNumber, "score": score}
        scoresData.push([firmCode, firmName,...salesValues.slice(1), targetActualizationRate, totalStockNumber]);
      }
  
      infoData.push([month, turkeyTargetActualizationRate]);
    }
  
  
    // var range = scoresSheet.getRange(1, 1, scoresData.length, scoresData[0].length);
    writeToSheet(scoresSheet, firmInfo)
    var range2 = scoresSheet.getRange(1,15, infoData.length, infoData[0].length);
    // range.setValues(firmInfo);
    range2.setValues(infoData);
  }
  
  function calculateTotalStockNumber(firmCode, stockValues, stockFirmValues, stockConnectionValues, stockFirmConnectionValues) {
  
    var totalStockNumber = 0;
  
    // sifirli
    // if (stockFirmConnectionValues[i] == firmCode && (stockConnectionValues[i] == 'Bayi Stokunda') ){
    // sifirsiz
    // if (stockFirmConnectionValues[i] == firmCode && (stockConnectionValues[i] == 'Bayi Stokunda' || stockValues[j] == "Bayi'ye Sevk/Yolda" ) ){
  
    for (var i = 0; i < stockConnectionValues.length; i++) {
      if (stockFirmConnectionValues[i] == firmCode && (stockConnectionValues[i] == 'Bayi Stokunda' || stockValues[j] == "Bayi'ye Sevk/Yolda" ) ){
        totalStockNumber++;
      }
    }
  
    
    // sifirli
    // if (stockFirmValues[j] == firmCode && (stockValues[j] == 'Bayi Stokunda') ){
    // sifirsiz
    // if (stockFirmValues[j] == firmCode && (stockValues[j] == 'Bayi Stokunda' || stockValues[j] == "Bayi'ye Sevk/Yolda" ) ){
    for (var j = 0; j < stockValues.length; j++) {
      if (stockFirmValues[j] == firmCode && (stockValues[j] == 'Bayi Stokunda' || stockValues[j] == "Bayi'ye Sevk/Yolda" ) ){
        totalStockNumber++;
      }
    }
  
    return totalStockNumber;
  }
  
  function writeToSheet(sheet, data) {
    var dataArray = [];
  
    // Add header row
    dataArray.push(['DMSFirmCode', 'DMSFirmName', 'Score Jan', 'Score Feb', 'Score March', 'TAR Jan','TAR Feb','TAR March', 'Stock Jan', 'Stock Feb', 'Stock March']);
  
    // Convert object to array of arrays
    for (var firmCode in data) {
      var firmInfo = data[firmCode];
      var row = [firmCode, firmInfo.Jan.firmName];
      
      // Add score for Jan
      row.push(firmInfo.Jan.score);
      row.push(firmInfo.Feb.score);
      row.push(firmInfo.March.score);
      row.push(firmInfo.Jan.targetActualizationRate);
      row.push(firmInfo.Feb.targetActualizationRate);
      row.push(firmInfo.March.targetActualizationRate);
      row.push(firmInfo.Jan.totalStockNumber);
      row.push(firmInfo.Feb.totalStockNumber);
      row.push(firmInfo.March.totalStockNumber);
      
      dataArray.push(row);
    }
  
    // Write data to sheet
    var range = sheet.getRange(1, 1, dataArray.length, dataArray[0].length);
    range.setValues(dataArray);
  }
  