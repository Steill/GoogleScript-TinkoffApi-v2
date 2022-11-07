function tradesOneTimeCalculations() {
  var tradesRange = SpreadsheetApp.getActive().getRangeByName('Trades');
  var summaryRange = SpreadsheetApp.getActive().getRangeByName('Summaries');
  if (tradesRange == null || summaryRange == null) return;

  var uniqueTikers = [];
  var trades = tradesRange.getValues();
  for (var i = 1; i < trades.length; i++) {
    var tiker = trades[i][1];
    if (tiker == "") break;

    var lastPos = 0;
    for (var k = i - 1; k > 0; k--) {
      if (trades[k][1] == tiker) {
        lastPos = k;
        break;
      }
    }
    var avgPrice = trades[i][3];
    var qtyTotal = trades[i][5];
    var qtyAtStart = 0;
    var lastVolume = 0;
    var daysPassed = "";
    var currency = "";
    var delta = "";
    var deltaPercent = "";
    if (lastPos > 0) {
      qtyAtStart = trades[lastPos][17];
      qtyTotal += qtyAtStart;
      if (qtyTotal * trades[i][5] > 0) {
        var avg = avgPrice;
        var qty = trades[i][5];
        var lastavg = trades[lastPos][18];
        avgPrice = (avg * qty + lastavg * qtyAtStart) / qtyTotal;
        lastVolume = avgPrice * qtyAtStart;
      }
      else {
        avgPrice = trades[lastPos][18];
        lastVolume = avgPrice * qtyAtStart;
        delta = (avgPrice - trades[i][3]) * trades[i][5];
        deltaPercent = delta / Math.abs(lastVolume);
        currency = trades[i][4];
      }
      var deltaTime = new Date(trades[i][0]).getTime() - new Date(trades[lastPos][0]).getTime();
      daysPassed = deltaTime / mills_per_day;
    }
    trades[i][17] = qtyTotal;
    trades[i][18] = avgPrice;
    tradesRange.offset(i, 12, 1, 9).setValues([[
      delta,        //12
      currency,     //13
      deltaPercent, //14
      lastPos,      //15
      qtyAtStart,   //16
      qtyTotal,     //17
      avgPrice,     //18
      lastVolume,   //19
      daysPassed    //20
    ]]);

    var isContained = false;
    for (var k = 0; k < uniqueTikers.length; k++) {
      if (tiker == uniqueTikers[k][0]) {
        isContained = true;
        var uniq = uniqueTikers[k];
        if (trades[i][5] > 0) {
          uniq[1] += trades[i][5];
          uniq[17] += trades[i][6];
        }
        else
          uniq[2] += -trades[i][5];
        uniq[4] = avgPrice;
        uniq[8] += trades[i][8];
        uniq[15] += trades[i][10];
        if (delta != "")
          uniq[10] += delta;
        break;
      }
    }
    if (isContained) continue;

    var newEntry = [
      tiker,        //0
      0,            //1
      0,
      0,
      avgPrice,     //4
      trades[i][4], //5
      0,
      trades[i][4],
      trades[i][8], //8
      trades[i][4],
      0,            //10
      trades[i][4],
      0,
      0,
      trades[i][4],
      trades[i][10],//15
      trades[i][4],
      trades[i][6], //17
      trades[i][4]];
    if (trades[i][5] > 0) {
      newEntry[1] = trades[i][5];
      newEntry[17] = trades[i][6];
    }
    else
      newEntry[2] = -trades[i][5];
    uniqueTikers.push(newEntry);
  }

  uniqueTikers.sort(function (a, b) {
    var aString = a[0].toLowerCase();
    var bString = b[0].toLowerCase();
    if (bString < aString) return 1;
    if (aString < bString) return -1;
    return 0;
  });
  for (var k = 0; k < uniqueTikers.length; k++) {
    var uniq = uniqueTikers[k];
    uniq[3] = uniq[1] - uniq[2];
    uniq[6] = uniq[4] * uniq[3];
    uniq[13] = uniq[10] - uniq[8];
    uniq[12] = uniq[13] / uniq[17];
    summaryRange.offset(k + 1, 0, 1, uniqueTikers[k].length).setValues([
      uniqueTikers[k]
    ]);
  }
  return;
}

//-----------------------------------------------------------------------------------------------
function getLastTrades(lastOperations = null) {
  var tradesRange = SpreadsheetApp.getActive().getRangeByName('Trades');
  if (tradesRange == null) return;

  var lastTrade = 0;
  var pivotTime = null;
  var lastOperationDate = null;
  var trades = tradesRange.getValues();
  if (lastOperations == null) {
    for (var i = trades.length - 1; i >= 1; i--) {
      if (pivotTime != null) {
        var deltaTime = pivotTime - new Date(trades[i][0]).getTime();
        if (deltaTime < mills_per_day) {
          lastTrade = i;
          continue;
        }
        lastOperationDate = trades[lastTrade][0];
        lastTrade = i;
        break;
      }
      if (trades[i][0] != "") {
        lastTrade = i;
        if (lastTrade > 1) {
          pivotTime = new Date(trades[lastTrade][0]).getTime();
          continue;
        }
        break;
      }
    }
    lastOperations = getOperations(lastOperationDate);
  }
  if (lastOperations == null) return null;

  for (var i = lastOperations.length - 1; i >= 0; i--) {
    var lastOp = lastOperations[i];
    if (lastOp.quantity > 0) {
      var volume = lastOp.price * lastOp.quantity;
      if (lastOp.otype == 'Sell') lastOp.quantity = -lastOp.quantity;

      var tiker = lastOp.tiker;
      if (tiker == 'USD000UTSTOM') tiker = '_USD';
      if (tiker == 'EUR_RUB__TOM') tiker = '_EUR';

      var date = new Date(lastOp.date);
      var comiss = lastOp.comiss;
      if (lastOp.currency != lastOp.comcurr)
        comiss = -comiss / lastOp.quantity * volume;
      var summary = volume + comiss;

      if (lastOp.currency == 'RUB') lastOp.currency = '₽';
      if (lastOp.currency == 'USD') lastOp.currency = '$';
      if (lastOp.currency == 'EUR') lastOp.currency = '€';

      tradesRange.offset(++lastTrade, 0, 1, 12).setValues([[
        date,
        tiker,
        lastTrade,
        lastOp.price,
        lastOp.currency,
        lastOp.quantity,
        volume,
        lastOp.currency,
        comiss,
        lastOp.currency,
        summary,
        lastOp.currency,
      ]]);
      tradesRange.offset(lastTrade, 21, 1, 1).setValue(lastOp.hash);
    }
  }
  tradesOneTimeCalculations();
  return;
}

//-----------------------------------------------------------------------------------------------
function getLastDividends(lastOperations = null)
{
  var divsRangeRUB = SpreadsheetApp.getActive().getRangeByName('DividendsRUB');
  var divsRangeUSD = SpreadsheetApp.getActive().getRangeByName('DividendsUSD');
  if (divsRangeRUB == null || divsRangeUSD == null) return;
  var dividendsR = divsRangeRUB.getValues();
  var dividendsU = divsRangeUSD.getValues();

  for (var i = 1; i < dividendsR.length; i++)
  {
    var lastDiv = dividendsR[i];
    if (lastDiv[1] == "") continue;
    if (lastDiv[2] == "")
    {
      var valueSum = 0;
      var payoutDate = null;
      var opTimeShifted = new Date(new Date(lastDiv[1]).getTime() + mills_per_day * 75);
      var operations = getOperations(lastDiv[1], opTimeShifted, lastDiv[0]); 
      for (var k = 0; k < operations.length; k++)
      {
        var lastOp = operations[k];
        if (payoutDate == null)
        {
          if (lastOp[3] == "Dividend")
          {
            payoutDate = new Date(lastOp[0]);
            valueSum += lastOp[7];
            continue;
          }
          continue;
        }
        var nextOpDate = new Date(lastOp[0]);
        if (nextOpDate.getTime() - payoutDate.getTime() > mills_per_day) break;
        if (lastOp[3] == "Dividend") valueSum += lastOp[7];
      }
      if (valueSum > 0)
      {
        divsRange.offset(i,2,1,1).setValue(payoutDate);
        divsRange.offset(i,13,1,4).setValues([[valueSum, lastDiv[5], lastDiv[10] - valueSum, lastDiv[5]]]);
      }
    }
    else 
      if (lastDiv[13] != "" && lastDiv[14] == "")
        divsRange.offset(i,14,1,3).setValues([[lastDiv[5], lastDiv[10] - lastDiv[13], lastDiv[5]]]);
  }
  return;
}

function getFullOperationList() {
  var debugRange = SpreadsheetApp.getActive().getRangeByName('Debug');
  if (debugRange == null) return;

  var pivotIds = [];
  var opCount = 0;
  var lastOperationDate = null;
  var operationList = debugRange.getValues();
  for (var i = 1; i < operationList.length; i++) {
    if (operationList[i][0] == "") {
      opCount = i - 3;
      if (opCount < 0) opCount = 0;
      var date = operationList[opCount][0];
      if (opCount > 0 && date != "") {
        lastOperationDate = date;
        for (var k = opCount; k > 0 && opCount - k < 10; k--)
          pivotIds.push(operationList[k][10]);
      }
      break;
    }
  }
  var lastOperations = getOperations(lastOperationDate);
  if (lastOperations == null) return null;

  for (var i = lastOperations.length - 1; i >= 0; i--) {
    var lastOp = lastOperations[i];
    if (pivotIds.includes(lastOp[10])) continue;
    debugRange.offset(++opCount, 0, 1, 12).setValues([lastOp]);
  }
  return lastOperations;
}
