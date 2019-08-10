//Function that sets range of values in Spreadsheet
function populateSheet() {
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var lastRow = ss.getLastRow();
  var lastColumn = ss.getLastColumn();
  var data = lastRow;
  //var index = 0;
  //var cell = sheet.getRange(2, 10);
  var urlCol = ss.getRange("O2:O"+lastRow).getValues();
  getHtml(urlCol, data);
  
}

//Function used to Iterate through the Url Column
function getHtml(urlCol, data) {
  var skip = 0;
  var options = {
    "method" : "GET",
    "muteHttpExceptions" : true
  };
  
  for (var i = 2; i <= data; i++) {
    
    var url = urlCol[i-2];
    //var response = UrlFetchApp.fetch(url, options);
    //var content = response.getContentText();
    getData(url, i, skip);
    
  }
}

//Function that parses Html for values
function getData(url, i, skip) {

  var temp = i;
  var tempurl = url;
  var options = {
    method : "GET",
    muteHttpExceptions : true
  };
  
  var response = UrlFetchApp.fetch(url, options);
  var content = response.getContentText();
  
  var start1 = 'Rear head room</th><td class="px-1 px-lg-0_75 px-xl-1 py-0_5">';
  var end1 = 'in.</td>';
  var start2 = 'Rear shoulder room</th><td class="px-1 px-lg-0_75 px-xl-1 py-0_5">';
  var end2 = 'in.</td>';
  var start3 = 'Cargo capacity, all seats in place</th><td class="px-1 px-lg-0_75 px-xl-1 py-0_5">';
  var end3 = 'cu.ft.</td>';
  var start4 = 'Maximum cargo capacity</th><td class="px-1 px-lg-0_75 px-xl-1 py-0_5">';
  var end4 = 'cu.ft.</td>';
  
  //var dataa;
  var cut1 = content.indexOf(start1);
  var cut2 = content.indexOf(start2);
  var cut3 = content.indexOf(start3);
  var cut4 = content.indexOf(start4);
  Logger.log(cut1);
  Logger.log(cut2);
  
  var finish1 = content.indexOf(end1, cut1);
  var finish2 = content.indexOf(end2, cut2);
  var finish3 = content.indexOf(end3, cut3);
  var finish4 = content.indexOf(end4, cut4);
  Logger.log(finish1);
  Logger.log(finish2);
  
  var data1 = content.substring(cut1, finish1);
  var data2 = content.substring(cut2, finish2);
  var data3 = content.substring(cut3, finish3);
  var data4 = content.substring(cut4, finish4);
  Logger.log(data1);

  
  var value1 = data1.substring(62, 66);
  var value2 = data2.substring(66, 70);
  var value3 = data3.substring(82, 86);
  var value4 = data4.substring(70, 75);
  //var value = [value1, value2, value3, value4];
  
  var comp1 = isNaN(value1);
  var comp2 = isNaN(value2);
  var comp3 = isNaN(value3);
  var comp4 = isNaN(value4);
  
  var comp = new Array;
  comp.push(comp1);
  comp.push(comp2);
  comp.push(comp3);
  comp.push(comp4);
  
  var checktest1 = 'n" c';
  var checktest2 = 'lass';
  var checktest3 = '  <m';
  var checktest4 = '="">';
  
  if (skip > 2) {
    var x = 0;
    while (x <= 3) {
      if (comp[x] == true) {
        switch(x) {
          case 0:
            value1 = "No Data";
            break;
          case 1:
            value2 = "No Data";
            break;
          case 2:
            value3 = "No Data";
            break;
          case 3:
            value4 = "No Data";
        }
        x++;
      }
      else {
        x++;
      }
    }
    setData(value1, value2, value3, value4, temp);
  }
  
  else if (value1 == checktest1 || value2 == checktest2 || value3 == checktest3 || value4 == checktest4) {
    var dataa = content;
    skip++;
    delete content;
    getData(tempurl, temp, skip); 
  }
  
  else {
    setData(value1, value2, value3, value4, temp);
  }
  
}


function setData(value1, value2, value3, value4, temp) {
  
  var index = temp - 2;
  var x = 0;
  var ss1 = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss1.getActiveSheet();
  var cell = sheet.getRange(2, 10);
  
  var cell1 = value1;
  var cell2 = value2;
  var cell3 = value3;
  var cell4 = value4;
  
  var isEmp = new Array;
  isEmp.push(value1);
  isEmp.push(value2);
  isEmp.push(value3);
  isEmp.push(value4);
  var coltest = '="">';
  
  while (x <= 3) {   
    if (isEmp[x] == '' || isEmp[x] == coltest) {
      switch(x) {
        case 0:
          value1 = "No Data";
          break;
        case 1:
          value2 = "No Data";
          break;
        case 2:
          value3 = "No Data";
          break;
        case 3:
          value4 = "No Data";
        }
      x++;
    }
    else {
      x++;
    }
  }
  
  cell.offset(index, 0).setValue(cell1);
  cell.offset(index, 1).setValue(cell2);
  cell.offset(index, 2).setValue(cell3);
  cell.offset(index, 3).setValue(cell4);

}
