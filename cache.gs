const cache = CacheService.getUserCache();
var cachedTFPairs = [];

function getCachedTFPairs() {
  if (cachedTFPairs.length == 0) {
    var cached = cache.get("TikerFigiPairs");
    if (cached != null) {
      var rows = cached.split("#");
      for (var i = 0; i < rows.length; i++) {
        var pair = rows[i].split(" ");
        cachedTFPairs.push({ 'tiker': pair[0], 'figi': pair[1] });
      }
    }
  }
}

function putCachedTFPairs() {
  var strcache = "";
  if (cachedTFPairs.length > 0) {
    cachedTFPairs.sort(function (a, b) {
      var aString = a[0].toLowerCase();
      var bString = b[0].toLowerCase();
      if (bString < aString) return 1;
      if (aString < bString) return -1;
      return 0;
    });
    strcache = cachedTFPairs[0].tiker + " " + cachedTFPairs[0].figi;
    for (var i = 1; i < cachedTFPairs.length; i++)
      strcache += "#" + cachedTFPairs[i].tiker + " " + cachedTFPairs[i].figi;
  }
  cache.put("TikerFigiPairs", strcache, 102000);
}
