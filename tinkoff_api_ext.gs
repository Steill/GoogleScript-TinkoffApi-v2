/**
 * @customfunction
 */
function getFigi(tiker, forceUpdate = false) {
  if (tiker == null) return null;
  tiker = tiker.toUpperCase();
  if (cachedTFPairs.length > 0) {
    for (var i = 0; i < cachedTFPairs.length; i++) {
      if (cachedTFPairs[i].tiker == tiker) {
        if (forceUpdate) {
          var data = makeApiGetCall_(`market/search/by-ticker?ticker=${tiker}`);
          if (data == null) return null;
          cachedTFPairs[i].figi = data.instruments[0].figi;
        }
        return cachedTFPairs[i].figi;
      }
    }
  }
  var data = makeApiGetCall_(`market/search/by-ticker?ticker=${tiker}`);
  if (data == null) return null;
  var figi = data.instruments[0].figi;
  cachedTFPairs.push({ 'tiker': tiker, 'figi': figi });
  return figi;
}

/**
 * @customfunction
 */
function getTiker(figi, forceUpdate = false) {
  if (figi == null) return null;
  figi = figi.toUpperCase();
  if (cachedTFPairs.length > 0) {
    for (var i = 0; i < cachedTFPairs.length; i++) {
      if (cachedTFPairs[i].figi == figi) {
        if (forceUpdate) {
          var data = makeApiGetCall_(`market/search/by-figi?figi=${figi}`);
          if (data == null) return null;
          cachedTFPairs[i].tiker = data.ticker;
        }
        return cachedTFPairs[i].tiker;
      }
    }
  }
  var data = makeApiGetCall_(`market/search/by-figi?figi=${figi}`);
  if (data == null) return null;
  var tiker = data.ticker;
  cachedTFPairs.push({ 'tiker': tiker, 'figi': figi });
  return tiker;
}

/**
 * @customfunction
 */
function getOperations(from_date = null, to_date = null, figiOrTiker = null) {
  if (from_date == null)
    from_date = props.getProperty("firstDayString");
  from_date = new Date(from_date).toISOString();

  if (to_date == null)
    to_date = new Date();
  else
    to_date = new Date(to_date);
  to_date = new Date(to_date.getTime() - to_date.getTimezoneOffset() * 60000).toISOString();

  var apicall = `operations?from=${from_date}&to=${to_date}`
  if (figiOrTiker != null) {
    if (figiOrTiker.length < 6)
      figiOrTiker = getFigi(figiOrTiker);
    apicall += `&figi=${figiOrTiker}`;
  }
  Logger.log(apicall);
  var operations = makeApiGetCall_(apicall);
  if (operations == null) return null;
  var operations = operations.operations;

  var values = [];
  getCachedTFPairs();
  var hashShift = props.getProperty("firstDaySecs").toString();
  for (var i = 0; i < operations.length; i++) {
    var n = operations[i];
    var tiker = getTiker(n.figi);
    var commission = "";
    var comCurrency = "";
    if (n.quantityExecuted > 0) {
      commission = -n.commission.value;
      comCurrency = n.commission.currency;
    }
    var txtHash = '';
    var rawHash = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, hashShift + n.id.toString());
    for (j = 0; j < rawHash.length; j++) {
      var hashVal = rawHash[j];
      if (hashVal < 0) hashVal += 256;
      if (hashVal.toString(16).length == 1) txtHash += '0';
      txtHash += hashVal.toString(16);
    }
    values.push({'date': n.date, 'figi': n.figi, 'tiker': tiker, 'otype': n.operationType, 'price': n.price, 'currency': n.currency, 'quantity': n.quantityExecuted, 'payment': n.payment, 'comiss': commission, 'comcurr': comCurrency, 'status': n.status, 'hash': txtHash });
  }
  putCachedTFPairs();
  return values;
}
