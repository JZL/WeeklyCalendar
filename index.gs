//random font for days
var msInDay = 24*60*60*1000
var newLines = "\n\n\n\n\n\n\n\n\n\n"
var hourTable = [[8,""],[9,""],[10,""],[11,""],[12,""],[1,""],[2,""],[3,""],[4,""],[5,""],[6,""],[7,""],[8,""]]
var longestLength=16

  //    [[start, end], "spaces"]
var boldArr = [[/A/g, "𝐀"],[/B/g, "𝐁"],[/C/g, "𝐂"],[/D/g, "𝐃"],[/E/g, "𝐄"],[/F/g, "𝐅"],[/G/g, "𝐆"],[/H/g, "𝐇"],[/I/g, "𝐈"],[/J/g, "𝐉"],[/K/g, "𝐊"],[/L/g, "𝐋"],[/M/g, "𝐌"],[/N/g, "𝐍"],[/O/g, "𝐎"],[/P/g, "𝐏"],[/Q/g, "𝐐"],[/R/g, "𝐑"],[/S/g, "𝐒"],[/T/g, "𝐓"],[/U/g, "𝐔"],[/V/g, "𝐕"],[/W/g, "𝐖"],[/X/g, "𝐗"],[/Y/g, "𝐘"],[/Z/g, "𝐙"],[/a/g, "𝐚"],[/b/g, "𝐛"],[/c/g, "𝐜"],[/d/g, "𝐝"],[/e/g, "𝐞"],[/f/g, "𝐟"],[/g/g, "𝐠"],[/h/g, "𝐡"],[/i/g, "𝐢"],[/j/g, "𝐣"],[/k/g, "𝐤"],[/l/g, "𝐥"],[/m/g, "𝐦"],[/n/g, "𝐧"],[/o/g, "𝐨"],[/p/g, "𝐩"],[/q/g, "𝐪"],[/r/g, "𝐫"],[/s/g, "𝐬"],[/t/g, "𝐭"],[/u/g, "𝐮"],[/v/g, "𝐯"],[/w/g, "𝐰"],[/x/g, "𝐱"],[/y/g, "𝐲"],[/z/g, "𝐳"]]
function doAll(emailNow){
  var document = DocumentApp.openById("1t5SK5nrz4DrpN0Ryjv3fpzLwjgUL62jExbV3RUM6acU")
  var body = document.getBody().clear()
  body.setMarginTop(0)
  body.setMarginBottom(28.8)
  //  body.setMarginLeft(28.8)
  body.setMarginLeft(0)
  body.setMarginRight(0)
  //resets font size  
  body.getChild(0).asText().setFontSize(11)
//  var thisMonday = getMonday(new Date())
  var thisMonday = getMonday(new Date("10/9/2016"))
  //true means want to NOT minimize size
  makePaper(thisMonday, body, false)
  body.appendPageBreak()
  body.appendParagraph("")
  body.getChild(0).asText().setFontSize(0)
//  makePaper(getMonday(new Date(thisMonday.getTime()+6*msInDay)), body, true)
  
  Utilities.sleep(1000)
  if(emailNow == false){
    MailApp.sendEmail({
      to: "jlangli1@swarthmore.edu",
      subject: "New Weekly Cal",
      htmlBody: "<a href='https://docs.google.com/document/d/1t5SK5nrz4DrpN0Ryjv3fpzLwjgUL62jExbV3RUM6acU/edit'>gDoc</a><br><a href='https://docs.google.com/a/swarthmore.edu/document/export?format=pdf&id=1t5SK5nrz4DrpN0Ryjv3fpzLwjgUL62jExbV3RUM6acU'>PDF</a>",
      //wasn't updating
      //      attachments: [DocumentApp.openById("1t5SK5nrz4DrpN0Ryjv3fpzLwjgUL62jExbV3RUM6acU").getAs('application/pdf')]
    });
  }
  //   MailApp.sendEmail("jlangli1@swarthmore.edu", "New Weekly Calendar", "https://docs.google.com/document/d/1t5SK5nrz4DrpN0Ryjv3fpzLwjgUL62jExbV3RUM6acU/edit \n https://docs.google.com/a/swarthmore.edu/document/export?format=pdf&id=1t5SK5nrz4DrpN0Ryjv3fpzLwjgUL62jExbV3RUM6acU")
}

function runWeekly(){
  //don't email now, give 1 hour or so to let update
  doAll(false)
}

function runWeeklyLater(){
  MailApp.sendEmail({
    to: "jlangli1@swarthmore.edu",
    subject: "New Weekly Cal",
    htmlBody: "<a href='https://docs.google.com/document/d/1t5SK5nrz4DrpN0Ryjv3fpzLwjgUL62jExbV3RUM6acU/edit'>gDoc</a><br><a href='https://docs.google.com/a/swarthmore.edu/document/export?format=pdf&id=1t5SK5nrz4DrpN0Ryjv3fpzLwjgUL62jExbV3RUM6acU'>PDF</a>",
    //wasn't updating
    attachments: [DocumentApp.openById("1t5SK5nrz4DrpN0Ryjv3fpzLwjgUL62jExbV3RUM6acU").getAs('application/pdf')]
  });
}
function makePaper(startDate, body, minimizeSize) {
  var cells = []
  var startDateTime = startDate.getTime()
  var usedTimes = {};
  for(var w = 0; w<=4; w++){
    var todayTime = startDateTime+msInDay*w
    usedTimes[todayTime] = []
    //60*12
    for(var i = 0; i<=720;i++){
      usedTimes[todayTime].push(0)
    }    
  }
  
  a = body.appendTable([["Mon "+getDateString(startDate), "Tue "+getDateString(new Date(startDateTime+msInDay))], ["Wed "+getDateString(new Date(startDateTime+2*msInDay)), "Thur "+getDateString(new Date(startDateTime+3*msInDay))], ["Fri "+getDateString(new Date(startDateTime+4*msInDay)), "Sat "+getDateString(new Date(startDateTime+5*msInDay))]])
  .setBorderColor("#D3D3D3")
 
  if(minimizeSize){
//    a.setBorderColor("#000000");
    a.setBorderColor("#D3D3D3");
  }else{
    a.setBorderColor("#D3D3D3")
  }
  
  
  
  //  for(var i=0, i<=2; i++){
  for(var i=0;i<=2; i++){
    a.getCell(i, 0).getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.CENTER)
    a.getCell(i, 1).getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.CENTER)
  }
  //  
  
  //unusual, TODO on, DUEDates, schedule, jlangli1, days off
  //keep order of first 3
  var allCals = ["swarthmore.edu_dop0bh53409lheq23ell6ignqk@group.calendar.google.com", "swarthmore.edu_d4fd5qnh3r5a7aqdc2hk9fk5ag@group.calendar.google.com", "swarthmore.edu_ji5sie4fh0ddijtuithbqdse28@group.calendar.google.com", "swarthmore.edu_g7nk3sf5s5ttg27r4cdjgpdtto@group.calendar.google.com","jlangli1@swarthmore.edu","fjl75not0nhq75hkhm2phkj67o@group.calendar.google.com"]
  //  var allCals = [CalendarApp.getCalendarById("jlangli1@swarthmore.edu"), CalendarApp.getCalendarById("swarthmore.edu_g7nk3sf5s5ttg27r4cdjgpdtto@group.calendar.google.com")]
  var weeksObj = {}
  var spacesObj = {}
  
  for(var i = 0; i<=13;i++){
    weeksObj[startDateTime+msInDay*i] = {"DUE": [], "TODO":[],"allDay": [], "0":[],"1":[],"2":[],"3":[],"4":[],"5":[],"6":[],"7":[],"8":[],"9":[],"10":[],"11":[],"12":[],"13":[],"14":[],"15":[],"16":[],"17":[],"18":[],"19":[],"20":[],"21":[],"22":[],"23":[]}
//    spacesObj[startDateTime+msInDay*i] = {"0":"","1":"","2":"","3":"","4":"","5":"","6":"","7":"","8":"","9":"","10":"","11":"","12":"","13":"","14":"","15":"","16":"","17":"","18":"","19":"","20":"","21":"","22":"","23":""}
  }
  for(var z = 0; z<allCals.length;z++){
    allCals[z] = CalendarApp.getCalendarById(allCals[z])
    //    var allDays = CalendarApp.getCalendarById("jlangli1@swarthmore.edu").getEvents(startDate, new Date(startDateTime+13*msInDay))
    var allDays = allCals[z].getEvents(startDate, new Date(startDateTime+6*msInDay))
    Logger.log(allDays)
    for(var i in allDays){
      if(allDays[i].isAllDayEvent()){
        Logger.log(allDays[i].getTitle())
        var eventStartDate = allDays[i].getAllDayStartDate().getTime()
        var eventEndDate = allDays[i].getAllDayEndDate().getTime()
        if(eventStartDate == eventEndDate){
          //needed to have multiday all day not go over and single day showup
          eventEndDate+=msInDay
        }
        for(var start = eventStartDate; start <eventEndDate; start+=msInDay){
          if(z==0){
            weeksObj[start].allDay.push(bold(allDays[i].getTitle()))
          }else if(z == 1){
            weeksObj[start].TODO.push(truncate(allDays[i].getTitle(), 48))
          }else if(z == 2){
            weeksObj[start].DUE.push(truncate(allDays[i].getTitle(), 48))
          }else{
            //TDD
            weeksObj[start].allDay.push(allDays[i].getTitle())
          }
        }
        
      }else{
//        Logger.log(allDays[i].getTitle()) 
        var eventStartTime = allDays[i].getStartTime()
        var eventEndTime = allDays[i].getEndTime()
        var dayStart = new Date(eventStartTime.getTime()).setHours(0,0,0,0)
        var dayEnd = new Date(eventEndTime.getTime()).setHours(0,0,0,0)
        
        
        var neededNumSpaces = 0
        var neededSpaces = ""
        var eventStartTimeTime = eventStartTime.getTime()
        var eventEndTimeTime = eventEndTime.getTime()
        Logger.log(allDays[i].getTitle())
//        Logger.log(usedTimes[dayStart])
        if(eventStartTime.getDay()>=1 && eventStartTime.getDay()<6){
          for(var q = eventStartTimeTime; q<eventEndTimeTime;q+=60*1000){
            var mins = ((q-dayStart-(8*60*60*1000))/(60*1000))
            if(mins > 720){
              continue;
            }
            if(usedTimes[dayStart][mins] > neededNumSpaces){
              neededNumSpaces = usedTimes[dayStart][mins]
            }
            usedTimes[dayStart][mins]++
          }
        }
        for(var u = 1; u<=neededNumSpaces;u++){
          neededSpaces+="  "
        }
        
        if(z == 0){
          weeksObj[dayStart][eventStartTime.getHours()].push(((eventStartTime.getMinutes() == 0)? ':00 ' : ':'+eventStartTime.getMinutes()+" ")+neededSpaces+"▿"+bold(allDays[i].getTitle()))
          weeksObj[dayEnd][(eventEndTime.getHours()).toString()].push((eventEndTime.getMinutes() == 0? ':00 ' : ':'+eventEndTime.getMinutes()+" ")+"$"+neededSpaces+"▵"+bold(allDays[i].getTitle()))
        }else if(z==1){
          weeksObj[dayStart].TODO.push(eventStartTime.getHours()+((eventStartTime.getMinutes() == 0)? ':00 ' : ':'+eventStartTime.getMinutes()+" ")+truncate(allDays[i].getTitle(), longestLength))
        }else if(z == 2){
          weeksObj[dayStart].DUE.push(eventStartTime.getHours()+((eventStartTime.getMinutes() == 0)? ':00 ' : ':'+eventStartTime.getMinutes()+" ")+truncate(allDays[i].getTitle(), longestLength))
        }else{
        weeksObj[dayStart][eventStartTime.getHours()].push(((eventStartTime.getMinutes() == 0)? ':00 ' : ':'+eventStartTime.getMinutes()+" ")+neededSpaces+"▿"+truncate(allDays[i].getTitle(), longestLength))
        weeksObj[dayEnd][(eventEndTime.getHours()).toString()].push((eventEndTime.getMinutes() == 0? ':00 ' : ':'+eventEndTime.getMinutes()+" ")+"$"+neededSpaces+"▵"+truncate(allDays[i].getTitle(), longestLength))
        }
      }
    }
    
  }
  
  
  
  var style = {};
  //  {FONT_SIZE=null, ITALIC=null, STRIKETHROUGH=null, BORDER_COLOR=#000000, FOREGROUND_COLOR=null, BOLD=null, LINK_URL=null, UNDERLINE=null, FONT_FAMILY=null, BACKGROUND_COLOR=null, BORDER_WIDTH=1.0}
  
  style[DocumentApp.Attribute.FONT_SIZE] = 6;
//  style[DocumentApp.Attribute.FONT_FAMILY] = 'Courier New';
  style[DocumentApp.Attribute.FONT_FAMILY] = 'Roboto Mono';
//  style[DocumentApp.Attribute.BORDER_COLOR] = '#FFFFFF';
  style[DocumentApp.Attribute.BORDER_COLOR] = '#000000';
  
  for(var i=0; i<=4;i++){

    var boldedEvents = []
    var thisDayTable = []
    var thisDayObj = weeksObj[startDateTime+msInDay*i]
    
    if(!minimizeSize){
      if(thisDayObj.DUE.length !== 0 ){
        thisDayTable.push(["", "                    DUE: □ "+thisDayObj.DUE.join("\n                         □ ")])
      }
      if(thisDayObj.allDay.length !== 0){
        thisDayTable.push(["∞", thisDayObj.allDay.join("\n")])      
      }
      
////      thisDayTable.push(["←", ""]) 
////      //don't want 0 bc will be at 12:00
////      for(var q = 1; q<8; q++){
////        if(thisDayObj[q]!=""){
////          thisDayTable[1][1]+=q+thisDayObj[q].sort().join("\n"+q).replace(/\$/g, "$")+"\n"
////        }
////      }
//      thisDayTable[thisDayTable.length-1][1] = thisDayTable[thisDayTable.length-1][1].replace(/\n$/, "$")
      if(thisDayTable[thisDayTable.length-1][1] == ""){
        thisDayTable.pop()
      }
      for(var q = 1; q<24;q++){
        
//        thisDayTable.push([armyToNormalTime(q), thisDayObj[q].sort().join("\n").replace(/\$/g, "$")])
        if((q >8 && q<=20)||thisDayObj[q].join("")!=""){
          thisDayTable.push([armyToNormalTime(q), thisDayObj[q].sort(compareToSpace)
                             .join("\n").replace(/\$/g, "")])
        }
       
      }
//      thisDayTable.push(["→", ""])
      Logger.log(thisDayTable.length)
//      for(var q = 21; q<24; q++){
//        if(thisDayObj[q]!=""){
//          thisDayTable[thisDayTable.length-1][1]+=armyToNormalTime(q)+thisDayObj[q].sort().join("\n"+armyToNormalTime(q)).replace(/\$/g, "$")+"\n"
//        }
//      }
      
      thisDayTable[thisDayTable.length-1][1] = thisDayTable[thisDayTable.length-1][1].replace(/\n$/, "")
      if(thisDayTable[thisDayTable.length-1][1] == ""){
        thisDayTable.pop()
      }
      if(thisDayObj.TODO.length !== 0 ){
        thisDayTable.push(["", "                   TODO: □ "+thisDayObj.TODO.join("\n                         □ ")])
      }
      var timeTable = a.getCell(Math.floor(i/2),i%2).appendTable(thisDayTable).setAttributes(style).setColumnWidth(0, 0)
//      timeTable.getCell(0, 0).editAsText().setBold(true)
      for(var z = 0; z<thisDayTable.length;z++){
        timeTable.getCell(z, 0).setPaddingBottom(0).setPaddingTop(0).setPaddingRight(0).setPaddingLeft(0).getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.RIGHT)
        timeTable.getCell(z, 1).setPaddingBottom(0).setPaddingTop(0).setPaddingRight(0).setPaddingLeft(0)
        //      timeTable.getCell(z, 2).setPaddingBottom(0).setPaddingTop(0).setPaddingRight(0).setPaddingLeft(0).setWidth(185)
      }
//      timeTable.getCell(0, 1).getChild(1).asText().setBold(true)
    }else{
      var todaysTimes = ""
      var numLines = 0
      for(var q in thisDayObj){
        if(thisDayObj[q]!=""){
          numLines++
            todaysTimes+=armyToNormalTime(q)+thisDayObj[q].sort(compareToSpace)
            .join("\n"+armyToNormalTime(q)).replace(/\$/g, "")+"\n"
        }
      }
      for(var z =numLines; z<8; z++){
        todaysTimes+="\n"
      }
      a.getCell(Math.floor(i/2),i%2).appendParagraph(todaysTimes).setAttributes(style)
    }
  }
  //sat and sun
  var satText = ""
  var thisDayObj = weeksObj[startDateTime+msInDay*5]
  var numLines = 0
  for(var q in thisDayObj){
    if(thisDayObj[q]!=""){
      numLines++
        satText+=armyToNormalTime(q)+thisDayObj[q].sort(compareToSpace)
        .join("\n"+armyToNormalTime(q)).replace(/\$/g, "")+"\n"
    }
  }
  if(minimizeSize){
    for(var i =numLines; i<4; i++){
      satText+="\n"
    }
  }else{
    for(var i =numLines; i<8; i++){
      satText+="\n"
    }
  }
  var sunText = ""
  var thisDayObj = weeksObj[startDateTime+msInDay*6]
  for(var q in thisDayObj){
    if(thisDayObj[q]!=""){
      sunText+=armyToNormalTime(q)+thisDayObj[q].sort(compareToSpace)
      .join("\n"+armyToNormalTime(q)).replace(/\$/g, "")+"\n"
    }
  }
  
  a.getCell(2,1).appendParagraph(satText).setAttributes(style)
  .appendHorizontalRule()
  
  a.getCell(2,1).appendParagraph("Sun "+getDateString(new Date(startDateTime+6*msInDay))).setAlignment(DocumentApp.HorizontalAlignment.CENTER).editAsText().setFontSize(11).setFontFamily("Arial")
  a.getCell(2,1).appendParagraph(sunText).setAttributes(style)
  
  //  a.getCell(2, 1).appendHorizontalRule()
  //  a.getCell(2,1).appendParagraph("Sun "+getDateString(new Date(startDateTime+6*msInDay))).setAlignment(DocumentApp.HorizontalAlignment.CENTER)
  
  //  timeTable = a.getCell(0,0).appendTable(hourTable).setColumnWidth(0, 65).setAttributes(style).setColumnWidth(0, 0)
  
  
  //  
  //  
  //  
  //  var newTable = timeTable.copy()
  //  a.getCell(0,1).appendTable(newTable)
  //  var newTable = timeTable.copy()
  //  a.getCell(1,0).appendTable(newTable)
  //   var newTable = timeTable.copy()
  //  a.getCell(1,1).appendTable(newTable)
  //  var newTable = timeTable.copy()
  //  a.getCell(2,0).appendTable(newTable)
  
  
  //  Logger.log(a.getAttributes())
  //  body.appendTable([["ToDo"]])
  //  .setBorderColor("#FFFFFF")
  //  .getRow(0).getCell(0).setWidth(553).getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.CENTER)
  
  body.appendParagraph("ToDo").setAlignment(DocumentApp.HorizontalAlignment.CENTER).setSpacingBefore(0)
  body.appendParagraph("Add Dates").setAlignment(DocumentApp.HorizontalAlignment.RIGHT)
  
  
}


//function getMonday(d) {
//  d = new Date(d);
//  var day = d.getDay(),
//      diff = d.getDate() + day + (day == 0 ? -6:1); // adjust when day is sunday
//  return new Date(d.setDate(diff));
//}

function getMonday(d){
  d = new Date(d.setHours(0,0,0,0))
  var day = d.getDay()
  if(day == 0){
    return new Date(d.getTime()+msInDay)
  }else if(day == 6){
    return new Date(d.getTime()+2*msInDay)
  }else{
    return new Date(d.getTime()+(8-day)*msInDay)
  }
}

function getDateString(d){
  return d.getMonth()+1+"/"+d.getDate()
}

function armyToNormalTime(hour){
  if(hour == "allDay"){
    return ""
  }else if(isNaN(hour) == true){
    return hour+": "
  }
  if(hour == 12){
    return 12
  }else{
    return hour%12
  }
}
//var hourTable = []
//for(var i = 8; i<=12; i++){
//hourTable.push([i, ""])
//}
//for(var i = 1; i<=8; i++){
//hourTable.push([i, ""])
//}
//JSON.stringify(hourTable)

function truncate(str, length){
//  if(z == 0 && isStart === true){
//    return bold(str.replace(/^\!/, ""))
//  }
  if(str.indexOf("!")==0){
    return str.replace(/^\!/, "")
  }
  if(str.length > length){
    return str.substring(0, length-2)+".."
  }else{
    return str
  }
}

function bold(str){
  for(i in boldArr){
    str = str.replace(boldArr[i][0], boldArr[i][1])
  }
  return str
}

function test(){
allCals = CalendarApp.getCalendarById("fjl75not0nhq75hkhm2phkj67o@group.calendar.google.com")
    //    var allDays = CalendarApp.getCalendarById("jlangli1@swarthmore.edu").getEvents(startDate, new Date(startDateTime+13*msInDay))
    var allDays = allCals.getEvents(new Date("11/28/2016"), new Date("11/29/2016"))
    for(var event in allDays){
    Logger.log(allDays[event].getTitle())
    }
}

function addSpaces(startTime, endTime, write){
  

  //if false, is end and then want to change usedTimes arr, else, don't change usedtimes
    //    [[start, end], "spaces"]
//    var usedTimes = []

  for(var p in usedTimes){
//    if()
  }
}

function compareToSpace(a,b){
  return (a.substring(0, a.indexOf(" "))).localeCompare(b.substring(0, b.indexOf(" ")));
}