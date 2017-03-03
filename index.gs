//observatory series unfortunate
//yarmulka
//sunday add spaces for nested dates
//!! doesn't truncate
//next days add bolded/etc (3d printing)s
//random font for days
//https://www.googleapis.com/webfonts/v1/webfonts?key=AIzaSyCU6bm3mMmZQCn-cbuHa1wzZAu7oh7OlT0
//https://developers.google.com/fonts/docs/developer_api
//DON"T USE FOR DST var msInDay = 24*60*60*1000
var newLines = "\n\n\n\n\n\n\n\n\n\n"
var hourTable = [[8,""],[9,""],[10,""],[11,""],[12,""],[1,""],[2,""],[3,""],[4,""],[5,""],[6,""],[7,""],[8,""]]
var longestLength=16

//    [[start, end], "spaces"]
//var boldArr = [[/A/g, "𝐀"],[/B/g, "𝐁"],[/C/g, "𝐂"],[/D/g, "𝐃"],[/E/g, "𝐄"],[/F/g, "𝐅"],[/G/g, "𝐆"],[/H/g, "𝐇"],[/I/g, "𝐈"],[/J/g, "𝐉"],[/K/g, "𝐊"],[/L/g, "𝐋"],[/M/g, "𝐌"],[/N/g, "𝐍"],[/O/g, "𝐎"],[/P/g, "𝐏"],[/Q/g, "𝐐"],[/R/g, "𝐑"],[/S/g, "𝐒"],[/T/g, "𝐓"],[/U/g, "𝐔"],[/V/g, "𝐕"],[/W/g, "𝐖"],[/X/g, "𝐗"],[/Y/g, "𝐘"],[/Z/g, "𝐙"],[/a/g, "𝐚"],[/b/g, "𝐛"],[/c/g, "𝐜"],[/d/g, "𝐝"],[/e/g, "𝐞"],[/f/g, "𝐟"],[/g/g, "𝐠"],[/h/g, "𝐡"],[/i/g, "𝐢"],[/j/g, "𝐣"],[/k/g, "𝐤"],[/l/g, "𝐥"],[/m/g, "𝐦"],[/n/g, "𝐧"],[/o/g, "𝐨"],[/p/g, "𝐩"],[/q/g, "𝐪"],[/r/g, "𝐫"],[/s/g, "𝐬"],[/t/g, "𝐭"],[/u/g, "𝐮"],[/v/g, "𝐯"],[/w/g, "𝐰"],[/x/g, "𝐱"],[/y/g, "𝐲"],[/z/g, "𝐳"]]
function doAll(emailNow){
  //  try{
  var document = DocumentApp.openById("1t5SK5nrz4DrpN0Ryjv3fpzLwjgUL62jExbV3RUM6acU")
  var body = document.getBody().clear()
  body.setMarginTop(0)
//  body.setMarginBottom(28.8)
  body.setMarginBottom(28.8)
  //  body.setMarginLeft(28.8)
  body.setMarginLeft(15)
  body.setMarginRight(0)
  body.setAttributes({FOREGROUND_COLOR:"#000000"})
  //resets font size  
  body.getChild(0).asText().setFontSize(11)//.setFontFamily(getRandomFont())
  var thisMonday = getMonday(new Date())
  Logger.log(thisMonday.getTime())
//  var thisMonday = getMonday(new Date("1/20/2017"))
  //true means want to NOT minimize size
  
  makePaper(thisMonday, body, false)
  body.appendPageBreak()
  body.appendParagraph("")
  body.getChild(0).asText().setFontSize(0)
  Logger.log("--------------")
  //    Logger.log(getMonday(addDaysDate(thisMonday, 6*msInDay)))
  makePaper(getMonday(addDaysDate(thisMonday, 6)), body, true)
//  makePaper(thisMonday, body, true)
  
  Utilities.sleep(1000)
  if(emailNow == false){
//    MailApp.sendEmail({
//      to: "jlangli1@swarthmore.edu",
//      subject: "New Weekly Cal",
//      htmlBody: "<a href='https://docs.google.com/document/d/1t5SK5nrz4DrpN0Ryjv3fpzLwjgUL62jExbV3RUM6acU/edit'>gDoc</a><br><a href='https://docs.google.com/a/swarthmore.edu/document/export?format=pdf&id=1t5SK5nrz4DrpN0Ryjv3fpzLwjgUL62jExbV3RUM6acU'>PDF</a>",
//      //wasn't updating
//      //      attachments: [DocumentApp.openById("1t5SK5nrz4DrpN0Ryjv3fpzLwjgUL62jExbV3RUM6acU").getAs('application/pdf')]
//    });
  }
  //   MailApp.sendEmail("jlangli1@swarthmore.edu", "New Weekly Calendar", "https://docs.google.com/document/d/1t5SK5nrz4DrpN0Ryjv3fpzLwjgUL62jExbV3RUM6acU/edit \n https://docs.google.com/a/swarthmore.edu/document/export?format=pdf&id=1t5SK5nrz4DrpN0Ryjv3fpzLwjgUL62jExbV3RUM6acU")
  //  }catch(e){
  //    MailApp.sendEmail("jonahmail1@gmail.com", "Problem with weeklyCal", e)
  //  }
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
  
  //unusual, TODO on, DUEDates, schedule, jlangli1, days off, SCCS Schedule
  //keep order of first 3
  var allCals = ["swarthmore.edu_dop0bh53409lheq23ell6ignqk@group.calendar.google.com", "swarthmore.edu_d4fd5qnh3r5a7aqdc2hk9fk5ag@group.calendar.google.com", "swarthmore.edu_ji5sie4fh0ddijtuithbqdse28@group.calendar.google.com", "swarthmore.edu_g7nk3sf5s5ttg27r4cdjgpdtto@group.calendar.google.com","jlangli1@swarthmore.edu","fjl75not0nhq75hkhm2phkj67o@group.calendar.google.com", "en.judaism#holiday@group.v.calendar.google.com","swarthmore.edu_0ha19taudgvpckfmel7okbq6ic@group.calendar.google.com"]
  //  var allCals = [CalendarApp.getCalendarById("jlangli1@swarthmore.edu"), CalendarApp.getCalendarById("swarthmore.edu_g7nk3sf5s5ttg27r4cdjgpdtto@group.calendar.google.com")]
  var noSpacesweeksObj = {}
  var weeksObj = {}
  var spacesObj = {}
  var hoursTaken = {}
  var goalsArr = []
  var hasGotGoals = false
  var todoTitle = []
  for(var w = 0; w<=6; w++){
    var todayTime = addDaysTime(startDate, w)
    usedTimes[todayTime] = []
    //60*12
    for(var i = 0; i<=720;i++){
      usedTimes[todayTime].push(0)
    }    
    //    var time = new Date(new Date(startDateTime+msInDay*i).setHours(0,0,0,0)).getTime()
    noSpacesweeksObj[todayTime] = {"DUE": [], "TODO":[],"allDay": [], "0":[],"1":[],"2":[],"3":[],"4":[],"5":[],"6":[],"7":[],"8":[],"9":[],"10":[],"11":[],"12":[],"13":[],"14":[],"15":[],"16":[],"17":[],"18":[],"19":[],"20":[],"21":[],"22":[],"23":[]}
    weeksObj[todayTime] = {"GOALS": [], "BOLD": [], "DUE": [], "TODO":[],"allDay": [], "0":[],"1":[],"2":[],"3":[],"4":[],"5":[],"6":[],"7":[],"8":[],"9":[],"10":[],"11":[],"12":[],"13":[],"14":[],"15":[],"16":[],"17":[],"18":[],"19":[],"20":[],"21":[],"22":[],"23":[]}
    hoursTaken[todayTime] = {"0":[],"1":[],"2":[],"3":[],"4":[],"5":[],"6":[],"7":[],"8":[],"9":[],"10":[],"11":[],"12":[],"13":[],"14":[],"15":[],"16":[],"17":[],"18":[],"19":[],"20":[],"21":[],"22":[],"23":[]}
  }
  
  a = body.appendTable([["Mon "+getDateString(startDate), "Tue "+getDateString(addDaysDate(startDate, 1))], ["Wed "+getDateString(addDaysDate(startDate, 2)), "Thur "+getDateString(addDaysDate(startDate, 3))], ["Fri "+getDateString(addDaysDate(startDate, 4)), "Sat "+getDateString(addDaysDate(startDate, 5))]])
  .setBorderColor("#000000")//.setColumnWidth(1, 270)
  
  if(minimizeSize){
    a.setBorderColor("#000000");
    //    a.setBorderColor("#D3D3D3");
  }else{
    a.setBorderColor("#D3D3D3")
  }
  
  //  for(var i=0, i<=2; i++){
  for(var i=0;i<=2; i++){
    a.getCell(i, 0).getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.CENTER)
    a.getCell(i, 1).getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.CENTER)
  }
  //  
  
  
  
  //  for(var i = 0; i<=7;i++){
  //    var time = new Date(new Date(startDateTime+msInDay*i).setHours(0,0,0,0)).getTime()
  //    noSpacesweeksObj[time] = {"DUE": [], "TODO":[],"allDay": [], "0":[],"1":[],"2":[],"3":[],"4":[],"5":[],"6":[],"7":[],"8":[],"9":[],"10":[],"11":[],"12":[],"13":[],"14":[],"15":[],"16":[],"17":[],"18":[],"19":[],"20":[],"21":[],"22":[],"23":[]}
  //    weeksObj[time] = {"GOALS": [], "BOLD": [], "DUE": [], "TODO":[],"allDay": [], "0":[],"1":[],"2":[],"3":[],"4":[],"5":[],"6":[],"7":[],"8":[],"9":[],"10":[],"11":[],"12":[],"13":[],"14":[],"15":[],"16":[],"17":[],"18":[],"19":[],"20":[],"21":[],"22":[],"23":[]}
  //    hoursTaken[time] = {"0":[],"1":[],"2":[],"3":[],"4":[],"5":[],"6":[],"7":[],"8":[],"9":[],"10":[],"11":[],"12":[],"13":[],"14":[],"15":[],"16":[],"17":[],"18":[],"19":[],"20":[],"21":[],"22":[],"23":[]}
  //    spacesObj[startDateTime+msInDay*i] = {"0":"","1":"","2":"","3":"","4":"","5":"","6":"","7":"","8":"","9":"","10":"","11":"","12":"","13":"","14":"","15":"","16":"","17":"","18":"","19":"","20":"","21":"","22":"","23":""}
  //  }
  for(var z = 0; z<allCals.length;z++){
    allCals[z] = CalendarApp.getCalendarById(allCals[z])
    //    var allDays = CalendarApp.getCalendarById("jlangli1@swarthmore.edu").getEvents(startDate, new Date(startDateTime+13*msInDay))
    var allDays = allCals[z].getEvents(startDate, addDaysDate(startDateTime, 7))
    Logger.log(allDays)
    for(var i in allDays){
      Logger.log(allDays[i].getTitle())
      if(allDays[i].isAllDayEvent()){
        var eventStartDate = allDays[i].getAllDayStartDate()
        var eventEndDate = allDays[i].getAllDayEndDate()
        Logger.log(eventStartDate.getTime())
        Logger.log(eventEndDate.getTime())
        if(eventStartDate.getTime() == eventEndDate.getTime()){
          //needed to have multiday all day not go over and single day showup
          eventEndDate  = addDaysDate(eventEndDate, 1)
        }
        var title = allDays[i].getTitle()
        //        var truncateTitle
        //        if(title.indexOf("!")==0){
        //          truncateTitle = title.replace(/^\!/, "")
        //        }else{
        //          truncateTitle = truncate(title, 48)
        //        }
        var numDaysBetween = daysBetween(eventStartDate, eventEndDate)
        //        Logger.log("days: "+daysBetween(eventStartDate, eventEndDate))
        for(var day = 0; day<numDaysBetween;day++){
          var start = addDaysTime(eventStartDate, day)
          if(start >= addDaysTime(startDate, 7) || start < startDateTime){
            continue;
          }
          //        for(var startTime = eventStartDate; startTime <eventEndDate; startTime+=msInDay){
          //          var start = new Date(new Date(startTime).setHours(0,0,0,0)).getTime()
          //          if(start >= startDateTime+7*msInDay){
          //            continue
          //          }
          if(z == 1){
            if(title.indexOf("::") == 0){
              var thisGoalArr = title.split("::")
              for(var goal in thisGoalArr){
                if(thisGoalArr[goal] == ""){
                  continue
                }
//                if(hasGotGoals==false){
                  goalsArr.push(thisGoalArr[goal].split(":"))
//                }
                todoTitle.push(thisGoalArr[goal].split(":")[0])
              }
              //unshift with reverse later makes repeared goals last
              if(weeksObj[start] == null){
//                Logger.log("start: "+start+"  "+weeksObj[start])
              }
              weeksObj[start].TODO.unshift(todoTitle.join(", □ "))
              for(var w in weeksObj[start].BOLD){
                if(weeksObj[start].BOLD[w][0] == "TODO"){
                  weeksObj[start].BOLD[w][1]++
                }
              }
              hasGotGoals = true
              
            }else{
              weeksObj[start].TODO.push(title)
              if(title.indexOf("!!")==0){
                weeksObj[start].BOLD.push(["TODO",  weeksObj[start].TODO.length-1, "TODO: "+title])
              }
            }
          }else if(z == 2){
            weeksObj[start].DUE.push(title)
            if(title.indexOf("!!")==0){
              weeksObj[start].BOLD.push(["DUE",  weeksObj[start].DUE.length-1])
            }
          }else{
            weeksObj[start].allDay.push(title)
            if(z == 0 || title.indexOf("!!")==0){
              weeksObj[start].BOLD.push(["allDay",  weeksObj[start].allDay.length-1, "allDay: "+title])
            }
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
//        Logger.log(allDays[i].getTitle())
        //        Logger.log(usedTimes[dayStart])
        //        if(eventStartTime.getDay()>=1 && eventStartTime.getDay()<6){
        //          for(var q = eventStartTimeTime; q<eventEndTimeTime;q+=60*1000){
        //            var mins = ((q-dayStart-(8*60*60*1000))/(60*1000))
        //            if(mins > 720){
        //              continue;
        //            }
        //            if(usedTimes[dayStart][mins] > neededNumSpaces){
        //              neededNumSpaces = usedTimes[dayStart][mins]
        //            }
        //            usedTimes[dayStart][mins]++
        //          }
        //        }
        //        for(var u = 1; u<=neededNumSpaces;u++){
        //          neededSpaces+="  "
        //        }
        
        /*if(z == 0){
        weeksObj[dayStart][eventStartTime.getHours()].push(((eventStartTime.getMinutes() == 0)? ':00 ' : ':'+eventStartTime.getMinutes()+" ")+neededSpaces+"▿"+bold(allDays[i].getTitle()))
        weeksObj[dayEnd][(eventEndTime.getHours()).toString()].push((eventEndTime.getMinutes() == 0? ':00 ' : ':'+eventEndTime.getMinutes()+" ")+"$"+neededSpaces+"▵"+bold(allDays[i].getTitle()))
        }else 
        */
        //last is true if want to be bold
        var title = allDays[i].getTitle()
        var truncateTitle
        if(title.indexOf("!")==0 ||z==0){
          truncateTitle = title.replace(/^\!/, "")
        }else{
          truncateTitle = truncate(title, 48)
        }
        
        
        
        if(noSpacesweeksObj[dayStart] == null){
//          Logger.log("@@@@@@@"+allDays[i].getTitle()+" "+new Date(eventStartTime))
        }
        
        noSpacesweeksObj[dayStart][eventStartTime.getHours()].push([[eventStartTime, eventEndTime], truncateTitle, (z==0 || title.indexOf("!!")==0)])
//        noSpacesweeksObj[dayStart][eventEndTime.getHours()].push([[eventStartTime, eventEndTime], truncateTitle, (z==0 || title.indexOf("!!")==0)])
        
        
        //for use to decide if empty hour for " |"
        hoursTaken[dayStart][eventStartTime.getHours()].push("")
        hoursTaken[dayEnd][eventEndTime.getHours()].push("")
        if(z==1){
          weeksObj[dayStart].TODO.push(eventStartTime.getHours()+((eventStartTime.getMinutes() == 0)? ':00 ' : ':'+eventStartTime.getMinutes()+" ")+truncateTitle)
        }else if(z == 2){
          weeksObj[dayStart].DUE.push(eventStartTime.getHours()+((eventStartTime.getMinutes() == 0)? ':00 ' : ':'+eventStartTime.getMinutes()+" ")+truncateTitle)
        }
        
        /*else if(z==0){
        weeksObj[dayStart].BOLD.push([eventStartTime.getHours(), noSpacesweeksObj[dayStart][eventStartTime.getHours()].length-1])
        }*/
        
        //        weeksObj[dayStart][eventStartTime.getHours()].push([((eventStartTime.getMinutes() == 0)? ':00 ' : ':'+eventStartTime.getMinutes()+" ")+neededSpaces+"▿"+truncate(allDays[i].getTitle(), longestLength),
        //          [(eventEndTime.getHours()).toString(), (eventEndTime.getMinutes() == 0? ':00 ' : ':'+eventEndTime.getMinutes()+" ")+"$"+neededSpaces+"▵"+truncate(allDays[i].getTitle(), longestLength)]
        //          ])
        //        weeksObj[dayEnd][(eventEndTime.getHours()).toString()].push((eventEndTime.getMinutes() == 0? ':00 ' : ':'+eventEndTime.getMinutes()+" ")+"$"+neededSpaces+"▵"+truncate(allDays[i].getTitle(), longestLength))
        
        
        
      }
    }
    
  }
  // Logger.log(JSON.stringify(noSpacesweeksObj))
  for(var i in noSpacesweeksObj){
    var dayStart = i
    for(var z = 0; z<=23;z++){
//      Logger.log("sort"+JSON.stringify(noSpacesweeksObj[i][z]))
      noSpacesweeksObj[i][z] = noSpacesweeksObj[i][z].sort(compareSecondElem)
      for(var p =0; p<noSpacesweeksObj[i][z].length;p++){
//        Logger.log( noSpacesweeksObj[i][z][p][0][0])
//        Logger.log( noSpacesweeksObj[i][z][p][0][1])
//        Logger.log("------")
        eventStartTime = noSpacesweeksObj[i][z][p][0][0]
        eventEndTime = noSpacesweeksObj[i][z][p][0][1]   
        var eventStartTimeTime = eventStartTime.getTime()
        //        eventStartTimeTime = new Date(new Date(eventStartTimeTime).setHours(0,0,0,0)).getTime()
        var eventEndTimeTime = eventEndTime.getTime()
        //        eventEndTimeTime = new Date(new Date(eventEndTimeTime).setHours(0,0,0,0)).getTime()
        neededSpaces = ""
        neededNumSpaces = 0
        
//        Logger.log("running "+(eventStartTimeTime))
        //        if(eventStartTime.getDay()>=1 && eventStartTime.getDay()<6){
        for(var q = eventStartTimeTime; q<eventEndTimeTime;q+=60*1000){
          //            var q = new Date(new Date(b).setHours(0,0,0,0)).getTime()
          var mins = ((q-dayStart-(8*60*60*1000))/(60*1000))
          //            Logger.log("mins: "+mins)
          if(mins > 720){
            continue;
          }
          if(usedTimes[dayStart][mins] > neededNumSpaces){
            neededNumSpaces = usedTimes[dayStart][mins]
          }
          usedTimes[dayStart][mins]++
        }
        //        }
        for(var u = 1; u<=neededNumSpaces;u++){
          //          neededSpaces+="  "
          neededSpaces+="  "
        }
        weeksObj[i][z].push([((eventStartTime.getMinutes() == 0)? ':00 ' : ':'+eventStartTime.getMinutes()+" ")+neededSpaces+"▽"+noSpacesweeksObj[i][z][p][1],noSpacesweeksObj[i][z][p][2]])
        weeksObj[i][eventEndTime.getHours()].push([(eventEndTime.getMinutes() == 0? ':00 ' : ':'+eventEndTime.getMinutes()+" ")+neededSpaces+"△"+noSpacesweeksObj[i][z][p][1], noSpacesweeksObj[i][z][p][2]])
        
//        weeksObj[i][z] = weeksObj[i][z].sort()
//        weeksObj[i][eventEndTime.getHours()] = weeksObj[i][eventEndTime.getHours()].sort()
        
//        Logger.log("hours"+(eventStartTime.getHours()+1)+", "+eventEndTime.getHours())
        for(var e = eventStartTime.getHours()+1; e<eventEndTime.getHours();e++){
//          Logger.log("hour "+e)
//          Logger.log("length "+hoursTaken[i][e].length)
//          if(hoursTaken[i][e].length == 0){
//          Logger.log(weeksObj[i][e])
          weeksObj[i][e].push(["  |"])
//            Logger.log(weeksObj[i][e])
            hoursTaken[i][e].push("|")
//          }
        }
      }
    }
    //    for(var z = 0; z<=23;z++){
    //      if(weeksObj[i][z].length >1  && weeksObj[i][z][0] == " |" ){
    //        weeksObj[i][z].splice(0,1)
    //      }
    //    }
  }
  
  
//          if(noSpacesweeksObj[i][z][p][2] == true){
//          if(z == eventEndTime.getHours()){
//            weeksObj[i].BOLD.push([z, weeksObj[i][z].length-2, "normal: "+noSpacesweeksObj[i][z][p][1]])
//          }else{
//            weeksObj[i].BOLD.push([z, weeksObj[i][z].length-1, , "normal: "+noSpacesweeksObj[i][z][p][1]])
//          }
//          
//          //if modify this, modify the push to eventEndTime.getHours() :(
//          Logger.log("looking for "+(eventEndTime.getMinutes() == 0? ':00 ' : ':'+eventEndTime.getMinutes()+" ")+neededSpaces+"▵"+noSpacesweeksObj[i][z][p][1]+" in "+JSON.stringify(weeksObj[i][eventEndTime.getHours()]))
//          //weeksObj[i].BOLD.push([eventEndTime.getHours(), weeksObj[i][eventEndTime.getHours()].length-1, , "normal: "+noSpacesweeksObj[i][z][p][1]])
//          weeksObj[i].BOLD.push([eventEndTime.getHours(), weeksObj[i][eventEndTime.getHours()].indexOf(endTitle) , "normal: "+noSpacesweeksObj[i][z][p][1]])
//        }
  for(var i in weeksObj){
    for(var z=0; z<=23;z++){
//      Logger.log(weeksObj[i][z])
      weeksObj[i][z] = weeksObj[i][z].sort()
      for(var q in weeksObj[i][z]){
        if(weeksObj[i][z][q][1] == true){
          weeksObj[i].BOLD.push([z, q])
        }
        weeksObj[i][z][q] = weeksObj[i][z][q][0]
      }
    }
  }
  //  Logger.log("-----------")
  //  Logger.log("-----------")
  //  Logger.log(JSON.stringify(weeksObj))
  
  
  
  var style = {};
  //  {FONT_SIZE=null, ITALIC=null, STRIKETHROUGH=null, BORDER_COLOR=#000000, FOREGROUND_COLOR=null, BOLD=null, LINK_URL=null, UNDERLINE=null, FONT_FAMILY=null, BACKGROUND_COLOR=null, BORDER_WIDTH=1.0}
  
  style[DocumentApp.Attribute.FONT_SIZE] = 6;
  //  style[DocumentApp.Attribute.FONT_FAMILY] = 'Courier New';
  //  style[DocumentApp.Attribute.FONT_FAMILY] = 'Roboto Mono';
  style[DocumentApp.Attribute.FONT_FAMILY] = 'Droid Sans Mono';
  style[DocumentApp.Attribute.BORDER_COLOR] = '#FFFFFF';
  //  style[DocumentApp.Attribute.BORDER_COLOR] = '#000000';
  
  for(var i=0; i<=4;i++){
    
    var boldedEvents = []
    var thisDayTable = []
    var thisDayObj = weeksObj[addDaysTime(startDate, i)]
    thisDayObj.TODO.push(todoTitle.join(", □ "))
//    var thisDayObj = []
//    for(var q =0; q<weeksObjPlusBold.length;q++){
////      thisDayObj.push(weeksObjPlusBold[q][0])
//    }
    
    
    var hoursToIndices = {}
    
    if(!minimizeSize){
      if(thisDayObj.DUE.length !== 0 ){
        thisDayTable.push(["", "                    DUE: □ "+thisDayObj.DUE.join("\n                         □ ")])
        hoursToIndices["DUE"] = thisDayTable.length
      }
      if(thisDayObj.allDay.length !== 0){
        thisDayTable.push(["∞", thisDayObj.allDay.join("\n")])   
        hoursToIndices["allDay"] = thisDayTable.length
      }
      
      ////      thisDayTable.push(["←", ""]) 
      ////      //don't want 0 bc will be at 12:00
      ////      for(var q = 1; q<8; q++){
      ////        if(thisDayObj[q]!=""){
      ////          thisDayTable[1][1]+=q+thisDayObj[q].sort().join("\n"+q).replace(/\$/g, "$")+"\n"
      ////        }
      ////      }
      //      thisDayTable[thisDayTable.length-1][1] = thisDayTable[thisDayTable.length-1][1].replace(/\n$/, "$")
      //      if(thisDayTable[thisDayTable.length-1][1] == ""){
      //        thisDayTable.pop()
      //      }
      for(var q = 1; q<24;q++){
        
        //        thisDayTable.push([armyToNormalTime(q), thisDayObj[q].sort().join("\n").replace(/\$/g, "$")])
        if((q >8 && q<=20)||thisDayObj[q].join("")!=""){
          thisDayTable.push([armyToNormalTime(q), thisDayObj[q]//.sort(compareToSpace)
                             .join("\n").replace(/\$/g, "")])
          hoursToIndices[q] = thisDayTable.length
        }
        
      }
      //      thisDayTable.push(["→", ""])
//      Logger.log(thisDayTable.length)
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
        //reverse makes repeared goals last
        thisDayTable.push(["", "                                   TODO: □ "+thisDayObj.TODO.reverse().join("\n                                         □ ")])
        hoursToIndices["TODO"] = thisDayTable.length
        for(var w in thisDayObj.BOLD){
          if(thisDayObj.BOLD[w][0] == "TODO"){
            //bc reverse array, need to reverse bold indexes
            thisDayObj.BOLD[w][1] = thisDayObj.TODO.length-thisDayObj.BOLD[w][1]-1
          }
        }
      }
      if(thisDayObj.DUE.length !== 0){
        hoursToIndices["DUE"] = 1
      }
      var timeTable = a.getCell(Math.floor(i/2),i%2).appendTable(thisDayTable).setAttributes(style).setColumnWidth(0, 8)
      //      timeTable.getCell(0, 0).editAsText().setBold(true)
      for(var z = 0; z<thisDayTable.length;z++){
        timeTable.getCell(z, 0).setPaddingBottom(0).setPaddingTop(0).setPaddingRight(0).setPaddingLeft(0).getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.RIGHT)
        timeTable.getCell(z, 1).setPaddingBottom(0).setPaddingTop(0).setPaddingRight(0).setPaddingLeft(0)
        //      timeTable.getCell(z, 2).setPaddingBottom(0).setPaddingTop(0).setPaddingRight(0).setPaddingLeft(0).setWidth(185)
      }
      Logger.log(thisDayObj.BOLD)
      for(var q in thisDayObj.BOLD){
        if(thisDayObj.BOLD[q][0] == "DUE"){
          Logger.log("")
        }
        Logger.log("BOLD "+thisDayObj.BOLD[q][0]+" , "+hoursToIndices[thisDayObj.BOLD[q][0]]+" , "+thisDayObj.BOLD[q][1])
        var cell = timeTable.getCell(hoursToIndices[thisDayObj.BOLD[q][0]]-1, 1).getChild(thisDayObj.BOLD[q][1]).asText().setBold(true).setFontSize(7)
        if(thisDayObj.BOLD[q][0] == "TODO"){
          cell.setFontSize(6)
        }
      }
      //      if(i==0){
      //        Logger.log("BOLD" +timeTable.getCell(5, 1).getChild(4).asText().setBold(true))
      //      }
      //      timeTable.getCell(0, 1).getChild(1).asText().setBold(true)
      
    }else{
      var todaysTimes = ""
      var numLines = 0
      for(var q in thisDayObj){
        if(q == "BOLD"){
          continue;
        }
        var spaces = ""
        if(i%2 == 0){
          spaces = "    " 
        }else{
          spaces = ""
        }
        if(thisDayObj[q]!=""){
          numLines++
            //2 spaces so can have 0 left margin
//            delete thisDayObj[q].BOLD
            todaysTimes+=spaces+armyToNormalTime(q, true)+thisDayObj[q]//.sort(compareToSpace)
            .join("\n"+spaces+armyToNormalTime(q, true)).replace(/\$/g, "")+"\n"
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
  var thisDayObj = weeksObj[addDaysTime(startDate, 5)]
  var numLines = 0
  for(var q in thisDayObj){
    if(q == "BOLD"){
      continue
    }
    if(thisDayObj[q]!=""){
      numLines++
        
//        if(thisDayObj[q].toString().substr(-1) == "|" && (armyToNormalTime(q) == "11"||armyToNormalTime(q) == "12")){
//          thisDayObj[q] = thisDayObj[q].toString().replace("__", "_")
//        }
//        satText+=(armyToNormalTime(q)+thisDayObj[q]//.sort(compareToSpace)
//        .join("\n"+armyToNormalTime(q)).replace(/\$/g, "")).replace("12  |", "12 |").replace("11  |", "11 |")+"\n"
        
        satText+=(" "+armyToNormalTime(q)+thisDayObj[q]//.sort(compareToSpace)
        .join("\n"+" "+armyToNormalTime(q)).replace(/\$/g, "")).replace(" 12", "12").replace(" 11", "11").replace(" 10", "10")+"\n"
        
    }
  }
  satText+="□ "+todoTitle.join(", □ ")
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
  var thisDayObj = weeksObj[addDaysTime(startDate, 6)]
  for(var q in thisDayObj){
    if(q == "BOLD"){
      continue
    }
    if(thisDayObj[q]!=""){
      sunText+=(armyToNormalTime(q)+thisDayObj[q].join("\n"+armyToNormalTime(q)).replace(/\$/g, "")).replace("12  |", "12 |").replace("11  |", "11 |").replace("10  |", "10 |")+"\n"
    }
    
  }
  
  var randFontsArr = getRandomFonts(8)
  Logger.log("FONTS "+ JSON.stringify(randFontsArr))
  a.getCell(0, 0).getChild(0).asText().setFontFamily(randFontsArr[0])
  a.getCell(0, 1).getChild(0).asText().setFontFamily(randFontsArr[1])
  a.getCell(1, 0).getChild(0).asText().setFontFamily(randFontsArr[2])
  a.getCell(1, 1).getChild(0).asText().setFontFamily(randFontsArr[3])
  a.getCell(2, 0).getChild(0).asText().setFontFamily(randFontsArr[4])
  a.getCell(2, 1).getChild(0).asText().setFontFamily(randFontsArr[5])
  a.getCell(0, 0).getChild(0).asText().setFontFamily(randFontsArr[6])
  a.getCell(2,1).appendParagraph(satText).setAttributes(style)
  .appendHorizontalRule()
  
  var addParagraph = ""
  var subscriptArr = ["","₁","₂","₃","₄","₅","₆","₇","₈","₉"]
  var superscriptArr = ["", "¹", "²", "³", "⁴","⁵","⁶","⁷","⁸","⁹"]
  var alternateArr = ["", "¹","₂","³","₄","⁵","₆","⁷"]
  for(var o in goalsArr){
    addParagraph+=goalsArr[o][0]+" __ |"
    for(var i = 1; i<=7;i++){
      if(i == parseInt(goalsArr[o][1])){
        addParagraph+=superscriptArr[i]+"   ❚"
      }else{
        addParagraph+=superscriptArr[i]+" |"
      }
    }
    addParagraph+="\n"
  }
  
  a.getCell(2,1).appendParagraph("Sun "+getDateString(addDaysDate(startDate, 6))).setAlignment(DocumentApp.HorizontalAlignment.CENTER).editAsText().setFontSize(11).setFontFamily(randFontsArr[7])
  a.getCell(2,1).appendParagraph(sunText).setAttributes(style)
  if(!minimizeSize){
    a.getCell(2,1).appendParagraph("").setAlignment(DocumentApp.HorizontalAlignment.RIGHT).appendText(addParagraph).setFontSize(9).setFontFamily("Droid Sans Mono")
  }
  //   a.getCell(2,1)
  //   .setAlignment(DocumentApp.HorizontalAlignment.RIGHT).appendParagraph(addParagraph).editAsText().setFontSize(11).setFontFamily("Droid Sans Mono")
  
  //  a.getCell(0, 0).getChild(0).asText().setFontFamily(getRandomFont())
  
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
  if(!minimizeSize){
    
    body.appendParagraph("Add Dates\t").setAlignment(DocumentApp.HorizontalAlignment.RIGHT).setSpacingBefore(0).editAsText().setFontFamily(randFontsArr[8])
    //    body.appendParagraph("Books Out\t\t\tAdd Dates").setAlignment(DocumentApp.HorizontalAlignment.RIGHT).editAsText().setFontFamily(randFontsArr[9])
  }else{
    var monad = ".  "
    var poly = ""
    var metaPoly = ""
    for(var i=0; i<57;i++){
      poly+=monad
    }
    poly = poly.replace(/\s*$/, "")
    for(var i = 0; i<=38;i++){
      metaPoly+=poly+"\n"
    }
//    body.appendParagraph(metaPoly).setSpacingBefore(0).editAsText().setForegroundColor("#551a8b")
    var a = body.getChild(body.getNumChildren()-1).asText().setFontSize(0)
    body.appendParagraph(metaPoly).setSpacingBefore(0).getChild(0).asText().setAttributes({FONT_SIZE:11, FOREGROUND_COLOR:"#808080", FONT_FAMILY:"Courier New"})
//    body.getChild(body.getNumChildren()-1).asText()
//    .setAttributes({FONT_SIZE:11, FOREGROUND_COLOR:"#808080", FONT_FAMILY:"Courier New"})
    //{FONT_SIZE=11, ITALIC=null, HORIZONTAL_ALIGNMENT=Left, INDENT_END=0.0, INDENT_START=0.0, LINE_SPACING=1.15, LINK_URL=null, UNDERLINE=null, BACKGROUND_COLOR=null, INDENT_FIRST_LINE=0.0, LEFT_TO_RIGHT=true, SPACING_BEFORE=0.0, HEADING=Normal, SPACING_AFTER=0.0, STRIKETHROUGH=null, FOREGROUND_COLOR=#000000, BOLD=null, FONT_FAMILY=Courier New}

//    .setAttributes({FONT_SIZE:11, LINE_SPACING:1.15, SPACING_BEFORE:0.0, FOREGROUND_COLOR:"#808080"})
//    a.editAsText().setAttributes({FONT_SIZE:11, LINE_SPACING:1.15, SPACING_BEFORE:0.0, FOREGROUND_COLOR:"#551a8b"})
//    .setAttributes({FONT_SIZE=11, ITALIC=null, STRIKETHROUGH=null, BORDER_COLOR=#000000, FOREGROUND_COLOR=null, BOLD=null, LINK_URL=null, UNDERLINE=null, FONT_FAMILY=null, BACKGROUND_COLOR=null, BORDER_WIDTH=1.0})
  }
  
  
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
    return addDaysDate(d.getTime(), 1)
  }else if(day == 6){
    return addDaysDate(d, 2)
  }else{
    return addDaysDate(d, (8-day))
  }
}

function getDateString(d){
  return d.getMonth()+1+"/"+d.getDate()
}

function armyToNormalTime(hour, padNums){
  if(hour == "allDay"){
    return ""
  }else if(hour == "TODO"){
    return "□ "
  }else if(isNaN(hour) == true){
    return hour+": "
  }
  if(hour == 12){
    return 12
  }else{
    if(padNums){
      return pad(hour%12, 2, " ")
    }else{
      return hour%12
    }
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

function compareSecondElem(a, b){
  return (a[1]).localeCompare(b[1]);
}
function testt(){
  Logger.log(getRandomFont())
}
function getRandomFonts(numFonts){
  var resp = UrlFetchApp.fetch("https://www.googleapis.com/webfonts/v1/webfonts?key=AIzaSyCU6bm3mMmZQCn-cbuHa1wzZAu7oh7OlT0").getContentText()
  var allFonts = JSON.parse(resp).items
  var retFonts = []
  for(var i = 0; i<=numFonts;i++){
    retFonts.push(allFonts[Math.floor(Math.random()*allFonts.length)].family)
  }
  return retFonts
  
}

function pad(n, width, z) {
  z = z || '0';
  n = n + '';
  return n.length >= width ? n : new Array(width - n.length + 1).join(z) + n;
}

function convertDateToUTC(date) { 
  return new Date(date.getUTCFullYear(), date.getUTCMonth(), date.getUTCDate(), date.getUTCHours(), date.getUTCMinutes(), date.getUTCSeconds()); 
}

function normalizeDate(d){
  return new Date(new Date(d).setHours(0,0,0,0))
}

function addDaysTime(date, days){
  return addDaysDate(date, days).getTime()
}

function addDaysDate(date, days){
  var dat = new Date(date.valueOf());
  dat.setDate(dat.getDate() + days);
  return dat
}

function treatAsUTC(date) {
  var result = new Date(date);
  result.setMinutes(result.getMinutes() - result.getTimezoneOffset());
  return result;
}

function daysBetween(startDate, endDate) {
  var millisecondsPerDay = 24 * 60 * 60 * 1000;
  return (treatAsUTC(endDate) - treatAsUTC(startDate)) / millisecondsPerDay;
}

function getSharples(){

}
