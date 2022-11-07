/** @OnlyCurrentDoc */
const mills_per_day = 1000 * 86400;
const props = PropertiesService.getUserProperties();
const token_ = "";


function getPortfolio()
{
  props.deleteProperty("token");
  props.setProperty("token", token_);
  var plio = makeApiGetCall_(`portfolio`)
  var portfolio = plio.positions;
  if (portfolio.length == 0) 
    return 'PORTFOLIO IS EMPTY';
  var values = [];
  for (var i = 0; i < portfolio.length; i++)
  {
    var n = portfolio[i];
    values.push({'figi':n.figi, 'tiker':n.ticker, 'balance':n.balance, 'avPP':n.averagePositionPrice.value, 'name':n.name});
  }
  return values;
}

function getP2()
{
  props.deleteProperty("token");
  props.setProperty("token", token_);
  props.deleteProperty("firstDaySecs");  
  props.setProperty("firstDaySecs", (new Date("01.02.2018").getTime() / 1000 + Math.floor(Math.random() * 5001)).toString());
  console.time("finish");
  console.log(authToken);
  getOperations();
  console.log(authToken);
  console.timeEnd("finish");
  
  console.time("trades");
  var tradesRange = SpreadsheetApp.getActive().getRangeByName('Trades').getValues();
  console.timeEnd("trades");
  //console.log(getTiker(getFigi("BA")));
  //console.log(authToken);
  return 1;
}
