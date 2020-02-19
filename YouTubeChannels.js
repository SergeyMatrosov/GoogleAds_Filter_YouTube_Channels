//Oбьявление функций, которые будут использоваться в программе
//Сама программа ниже. НИЖЕ!
//По вопросам скрипта обращаться к Сергею Матросову (но он может уже не помнить скрип ко времени обращения)

//функция обновить эксель-лист на драйве
//очень важен порядок аргументов!!!
function updateSpreadSheet(sheetName, sheetForBanList, channelList, channelListRowsNumber, channelListColumnsNumber) {
 
 	var searchSheet = DriveApp.getFilesByName(sheetName);
   	var url = searchSheet.next().getUrl();
   	var mainSheet = SpreadsheetApp.openByUrl(url);
    
    var mainSheelAllData = mainSheet.getDataRange().getValues();
    var rows = mainSheelAllData.length;
    var columns = mainSheelAllData[0].length;
    
    //индексация в листе начинается не с 0, а с 1. В 1 - названия колонок.
    //2, 2 - начальная row и начальная column; rows - количество строк для выгрузки; 1 - первая колонка как раз с Url;
    var channelUrlFromSheet = mainSheet.getSheetValues(2, 1, rows, columns).toString();
    
    var verifiedUrlList = [];
  	
  	//проверяем, есть ли этот URL в нашем списке (по идее не должно быть, для страховки);
    for (var i = 1; i < channelList.length; i++) {
          var urlToCheck = channelList[i][0];
          var checker = channelUrlFromSheet.search(urlToCheck);
          
          //-1 возвращается, если значение из channelList (за вчера) не было найдено в имеющемся в эксель-файле. Любая другая цифра говорит о том, что значение уже есть;
          //с другими методами почему-то возникает проблема;
          if (checker == -1) {
               var passedUrl = verifiedUrlList.push(channelList[i]);
              continue;
          }
    
          Logger.log('Nothing');
      }
    
  	//теперь подключаем второй лист с неподходящими словами
    var secondSheet = mainSheet.getSheetByName(sheetForBanList);
    var secondSheetData = secondSheet.getDataRange().getValue();
  
  	//удаление лишних строк
  	try {
    	var numberRowsToDelete = secondSheet.getMaxRows() - secondSheet.getLastRow();
    	var deleteRows = secondSheet.deleteRows(secondSheet.getLastRow() + 1, numberRowsToDelete);
    }
  	catch(err) {
    	Logger.log('Лишние строки удалять не пришлось!');
    }
  
    rows = secondSheet.getMaxRows();
    columns = secondSheet.getMaxColumns();
    var badNamesDataFrame = secondSheet.getSheetValues(1, 1, rows, columns);

  	var arrayToReturn = [];
	
  	//проверяем, содержат ли title, description, topic слова из бан-листа, который мы сформировали после создания файла
  	//если содержит, запихиваем канал и его описание в лист
  	//начинаем с 1, чтобы заголовок не проверять
    for (var i = 1; i < badNamesDataFrame.length; i++) {
      var badWord = badNamesDataFrame[i][0].toLowerCase();
      var modifiedBadWord = ' ' + badWord + ' ';
      
      for (var j = 0; j < verifiedUrlList.length; j++)  {
        var verifiedUrlListString = verifiedUrlList[j].join('|').toLowerCase();
      	var titleCheck = verifiedUrlListString.search(modifiedBadWord);
        var descriptionCheck = verifiedUrlListString.search(modifiedBadWord);
        var topicCheck = verifiedUrlListString.search(modifiedBadWord);
        
        if (titleCheck !== -1 || descriptionCheck !== -1 || topicCheck !== -1) {
           var passedItems = mainSheet.appendRow(verifiedUrlList[j]);
           arrayToReturn.push(verifiedUrlList[j]); 
           //Logger.log('This is bad');
          
        }
      }
    }
       
    Logger.log('Sheet updated');
  	return arrayToReturn;	
}
 

//создание нового (обычно, если впервые запущен скрипт); без фильтрации! 
function newSpreadSheet(sheetName, sheetForBanList, channelList, channelListRowsNumber, channelListColumnsNumber) {
  
 	var newSpread = SpreadsheetApp.create(sheetName, channelListRowsNumber, channelListColumnsNumber);
 	var url = newSpread.getUrl();
    var sheet = SpreadsheetApp.openByUrl(url);
  	
  	//меняем название первого листа в файле
    var getSheet = sheet.getSheets();
  	var setSheetName = getSheet[0].setName('Exclude List');
   
  	var passHeaders = sheet.appendRow(channelList[0]);
    /*
  	for (var i = 0; i < channelList.length; i++) {
    	var changeSheet = sheet.appendRow(channelList[i]);
    	}
    */
   	
   	//будет использоваться как хранилище "плохих" слов
   	var secondSheet = sheet.insertSheet(sheetForBanList);
  	var passFirstBanWord = secondSheet.appendRow(['BadWords']);
  	
  	//удаление лишних строк и колонок
  	var numberColumnsToDelete = secondSheet.getMaxColumns() - secondSheet.getLastColumn();
    var numberRowsToDelete = secondSheet.getMaxRows() - secondSheet.getLastRow();
  	
  	var deleteColumns = secondSheet.deleteColumns((secondSheet.getLastColumn() + 1), numberColumnsToDelete);
    var deleteRows = secondSheet.deleteRows(secondSheet.getLastRow() + 1, numberRowsToDelete);
   	
    Logger.log('New sheet made');
    Logger.log('Зайди на драйв в файл ' + sheetName + ' и на листе ' + sheetForBanList + ' добавь слова для исключений');
   	Logger.log('Перезапусти скрипт или выстави расписание');
 	
}


//функция обновить общего на аккаунт листа с negative placements
//очень важен порядок аргументов!!!
function updateSharedNegativeList(sheetName, arrayToReturn) {
  var excludeChannelList = AdWordsApp.excludedPlacementLists().withCondition('Name CONTAINS ' + sheetName).get().next();
  for (var i = 0; i < arrayToReturn.length; i++) {
      	excludeChannelList.addExcludedPlacement("https://youtube.com/channel/" + arrayToReturn[i][0].toString());
 	  }
  
  Logger.log('NegativeList updated'); 
}

//создание нового (без фильтрации)
function newSharedNegativeList(sheetName, channelList) {
  var excludeChannelListBuilder = AdWordsApp.newExcludedPlacementListBuilder().withName(sheetName).build();
  
  /*
  for (var i = 0; i < channelList.length; i++) {
      	excludeChannelListBuilder.getResult().addExcludedPlacement("https://youtube.com/channel/" + channelList[i][0].toString());
  };
  */
  Logger.log('NegativeList made');
}


//НАЧАЛО
function main() {
  
  //названия листов (можно менять, если хотите назвать по-своему - при первом запуске все равно будет создание объектов
  var sheetName = 'VideoChannelsExcluded';
  var nameForExcludeList = 'VideoChannelsExcluded';
  var sheetForBanList = 'RestrictedNames';
  
  //здесь вставляем все, что нас не устраивает в описании;
  var titleBanList = [];
  var descriptionBanList = [];
  var topicBanList = [];
  
  var report = AdWordsApp.report(
   	"SELECT Url " +
    "FROM URL_PERFORMANCE_REPORT " +
    "WHERE CampaignName CONTAINS 'YT_' " +
    "DURING LAST_MONTH"
   );
  
   var videoIdList = [];
   var rows = report.rows();
   while (rows.hasNext()) {
      var row = rows.next();
      var videoUrl = row['Url'].toString().replace('www.youtube.com/video/', '');
	  videoIdList.push(videoUrl);
   }
   
   videoIdList.sort();


   var channelList = [[]];
   var prevChannel = 0;
   
   //заголовки колонок
   var headerList = ['URL', 'Title', 'Description', 'Topics'];
   
   //вставляем заголовки
   for (var i = 0; i < headerList.length; i++) {
   	  channelList[0][i] = headerList[i];
   }

   //чтобы вставлять ПОСЛЕ названия колонок. Колонки находятся в индексе 0
   var counter = 1;
  
   //составляем массив для таблицы/списка исключений
   for(var i = 0; i < videoIdList.length; i++) {
     var video = videoIdList[i].toString();
     var topicDetails = YouTube.Videos.list('snippet,topicDetails', {id: video});
     
     try {
          var channel = topicDetails.items[0].snippet.channelId;
          try {
              var topicCategories = topicDetails.items[0].topicDetails.topicCategories;
              var title = topicDetails.items[0].snippet.title;
              var description = topicDetails.items[0].snippet.description;
              var topicCategoriesList = [];

              for (var j = 0; j < topicCategories.length; j++) {
                  var string = topicCategories[j];
                  var indexForSlice = string.lastIndexOf('/wiki/') + 6;
                  var sliced = string.slice(indexForSlice);
                  var category = sliced.replace('_', '');
     
                  topicCategoriesList.push(category);
              }

            var toStringCategoriesList = topicCategoriesList.toString();
            channelList[counter] = [channel, title, description, toStringCategoriesList];
            //Logger.log(channelList[0]);
            counter++;
           }

           catch(err) {
               Logger.log('Error is on channel ' + channel);
           }

           finally {
              continue;
           }
   }
     
   catch(err) {
     	Logger.log('Error is on video ' + video);
   }
     
   finally {
        continue;
     }
  }
 
 var channelListColumnsNumber = channelList[0].length;
 var channelListRowsNumber = channelList.length;
 
 var checkFileOnDrive = DriveApp.getFilesByName(sheetName).hasNext();
 Logger.log(checkFileOnDrive);
 
 if (checkFileOnDrive) {
 //очень важен порядок аргументов!!!
 	updateSpreadSheet(sheetName, sheetForBanList, channelList, channelListRowsNumber, channelListColumnsNumber);
    updateSharedNegativeList(nameForExcludeList, channelList);
   
   	return;
 }
  
 newSpreadSheet(sheetName, sheetForBanList, channelList, channelListRowsNumber, channelListColumnsNumber);
 newSharedNegativeList(nameForExcludeList, channelList);
}

