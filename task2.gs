access_token = '' // VK token
ss = SpreadsheetApp.getActive()
sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0] // Лист

function Pendragon_wall() {
  var link = '1369223'
  /*
   A - ID поста
   B - Дата
   C - Подпись к картинке
   D - Картинка
   E - Подпись к картинке из репоста
  */
  var html = UrlFetchApp.fetch('https://api.vk.com/method/wall.get.json?owner_id=-'+link+'&access_token='+access_token+'&v=5.131').getContentText()
  var json = JSON.parse(html)
  var response = json['response']['items']
  for (var n = 0; n < json['response']['count']; n++){
    try {
    var id = response[n]['id']
    var date = new Date(response[n]['date'] * 1000)
    var text = response[n]['text']
    var copy = response[n]['copy_history']
    if (copy == null) {
        var attach = response[n]['attachments']
        try {
          var length = attach.length
          }
      catch(e) {
        var length = 0
        }
        for (var i = 0; i < length; i++) {
        if (attach[i]['photo'] != null){
          var photo_size = attach[i]['photo']['sizes']
          for (var j = 0; j < photo_size.length; j++){
           var photo_rez = photo_size[j]['height']
           if (photo_rez > 300){
             var photo_url = photo_size[j]['url']
             var j = photo_size.length + 1
           }
          }
          var lastRow = sheet.getLastRow()+1
          var exist = 0
          for (var l = 1; l < lastRow; l++){
            var images = sheet.getRange(l, 4).getFormula()
            if (images == '=IMAGE("'+(photo_url)+'"; 4; 320; 320)'){
              exist = 1
            }
          }
          if (exist == 0){
            sheet.getRange("A"+lastRow).setValue(id)
            sheet.getRange("B"+lastRow).setValue(date)
            sheet.getRange("C"+lastRow).setValue(text)
            sheet.getRange("D"+lastRow).setValue('=IMAGE("'+(photo_url)+'"; 4; 320; 320)')
            sheet.setRowHeight(lastRow, 320)
          }
        }
      }
    }
    else {
      var further = copy[0]['attachments']
      var copy_text = copy[0]['text']
      for (var i = 0; i < further.length; i++) {
        if (further[i]['photo'] != null){
          var photo_size = further[i]['photo']['sizes']
          for (var j = 0; j < photo_size.length; j++){
           var photo_rez = photo_size[j]['height']
           if (photo_rez > 300){
             var photo_url = photo_size[j]['url']
             var j = photo_size.length + 1
           }
          }
          var lastRow = sheet.getLastRow()+1
          var exist = 0
          for (var l = 1; l < lastRow; l++){
            var images = sheet.getRange(l, 4).getFormula()
            if (images == '=IMAGE("'+(photo_url)+'"; 4; 320; 320)'){
              exist = 1
            }
          }
          if (exist == 0){
            sheet.getRange("A"+lastRow).setValue(id)
            sheet.getRange("B"+lastRow).setValue(date)
            sheet.getRange("C"+lastRow).setValue(text)
            sheet.getRange("D"+lastRow).setValue('=IMAGE("'+(photo_url)+'"; 4; 320; 320)')
            sheet.getRange("E"+lastRow).setValue(copy_text)
            sheet.setRowHeight(lastRow, 320)
          }
        }
      }
    }
  }
    catch (e){
    }
  }
  sheet.setColumnWidth(4, 320)
  var fullRange = sheet.getRange("A1:Z1001");
  fullRange.setVerticalAlignment("top")
  fullRange.setHorizontalAlignment("left")
  fullRange.setWrap(true)

  ScriptApp.newTrigger('Pendragon_wall').forSpreadsheet(ss).onOpen().create()

}


