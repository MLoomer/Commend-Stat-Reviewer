function single(e) {
  let folderID = 'FOLDER_ID_HERE';
  let parentFolder = DriveApp.getFolderById(folderID);
  let name = "QLIK"
  let blob = e.upFile.setContentType(MimeType.MICROSOFT_EXCEL).copyBlob()
  let file = DriveApp.createFile(blob);
  file.moveTo(parentFolder);

  let id = file.getId();
  blob = file.getBlob();
  let newFile = {
    title: name,
    "parents": [{ 'id': folderID }],
    key: id
  }
  filetwo = Drive.Files.insert(newFile, blob, {
    convert: true
  });

  id = filetwo.getId();
  let ss = SpreadsheetApp.openById(id);
  let url = ss.getUrl();
  file.setTrashed(true);

  file = newFile;

  let dataSheet = ss.getActiveSheet().setName("data");
  let titles = dataSheet.getRange("A1:M1").getDisplayValues();
  let range = dataSheet.getRange(2, 1, dataSheet.getLastRow() - 1, 13);
  let data = range.getDisplayValues();

  //sort if needed
  console.log(data[data.length - 1][0])
  let firstDate = new Date(data[0][0])
  let lastDate = new Date(data[data.length - 1][0])
  console.log(firstDate, lastDate)

  sortCells()

  function sortCells() {
    if (firstDate > lastDate) {
      data.sort((a, b) => new Date(a[0]) - new Date(b[0]))
      range.setValues(data)
      flush()
      firstDate = new Date(data[0][0])
      lastDate = new Date(data[data.length - 1][0])
    }
  }
  //set file name
  DriveApp.getFileById(id).setName(firstDate.toISOString().split('T')[0] + " - " + secondDate.toISOString().split('T')[0] + " - Single Agent Report")

  //TODO need to change this to earliest time in the array, and latest time in the array, ignoring dates

  let firstTime = returnTime(firstDate)
  let lastTime = returnTime(lastDate)
  let firstHour = firstDate.getHours()
  let lastHour = lastDate.getHours()

  let firstSeconds = secondsToSubtract(firstHour)
  let endSeconds = getLast(lastHour)
  let arraySize = getSize(firstSeconds, endSeconds)

  console.log(`size is: ${arraySize}, first time is: ${firstSeconds}, last time is: ${endSeconds}`)

  function secondsToSubtract(hour) {
    return ((hour * 3600) - 1)
  }

  function getSize(firstSeconds, endSeconds) {
    return (endSeconds - firstSeconds)

  }

  function getLast(last) {
    return ((last + 1) * (3600) - 1)

  }

  function returnTime(dateTime) {
    let time = [...dateTime.toLocaleTimeString()]
    time.splice(4, 3)
    return time.join("");
  }

  //qlikData = data

  //summary related stuff

  let summarySheet = ss.insertSheet("Summary");

  setTimeColumn()

  function setTimeColumn() {
    summarySheet.insertRowsAfter(1, (arraySize - 1000));
    SpreadsheetApp.flush()
    summarySheet.getRange("A1").setValue("Time");
    let range = summarySheet.getRange(2, 1, arraySize, 1);
    let arr = [];
    //
    let time = new Date(new Date().setHours(firstHour, 0, 0, 0))
    for (let i = firstSeconds; i < endSeconds; i++) {
      time.setSeconds(time.getSeconds() + 1)
      arr.push([time.toLocaleTimeString('en-US')])
    }
    range.setValues(arr);
  }

  //no queue stuff, remove
  
  let waitCharts = [];

  //TODO change to dates

  let nameArr = sortUnique(data, 5);

  let twoArr = [];
  let sixMinArr = [];
  let fiveFiveTwoArr = [];
  let size = arraySize
  console.log("size is " + size)
  let allUserArr = [];
  let deviceArr = [];
  let totalRow = []
  let initiatedArr = Array(size).fill('');
  initiatedArr.unshift("Initiated")

  for (let activeDate of nameArr) {
    if (activeDate.length > 1) {
      ss.insertSheet(activeDate)
      let sheet = ss.getSheetByName(activeDate);

      let agentArr = data.filter((row) => { return (row[5] == activeDate) })
      let onlyWaits = agentArr.filter((row) => (!row[10].includes("-")))
      let team = agentArr[0][4]

      let prevEndTime;
      let callArr = Array(size).fill('');
      callArr.unshift(activeDate)


      agentArr = agentArr.map((row, i) => {
        let start = getStart(row)
        let end = getEnd(start, row)
        let startTime = start.toLocaleString()
        let endTime = end.toLocaleString()
        let timeBetween = getBetween(start, prevEndTime)
        let callResponse = getResponse(timeBetween, row);

        row.push(startTime, endTime, timeBetween, callResponse)

        checkDuration(row[9], activeDate)

        function checkDuration(time, activeDate) {
          let roundedNum = Math.floor(+time)
          if (roundedNum < 3 && (!row[10].includes("-"))) { twoArr.push([row[0], row[12], roundedNum, activeDate]) }
          if ((roundedNum == 351) || (roundedNum == 352)) { fiveFiveTwoArr.push([activeDate, row[0]]) }
          if (roundedNum > 359) { sixMinArr.push([row[0], row[12], roundedNum, activeDate]) }
        }

        function getBetween(start, prevEndTime) {
          if (prevEndTime && prevEndTime != "Dropped") {
            let num = ((start - prevEndTime) / 1000)
            if (num > 0) {
              return num
            } else {
              return null
            }
          } else {
            return null;
          }
        }

        function getResponse(timeBetween, row) {
          let num = Math.round(row[10])
          //start on 2nd row
          if (prevEndTime && (!row[10].includes("-"))) {
            if (timeBetween < Math.round(row[10])) {
              return timeBetween;
            } else {
              return Math.round(row[10]);
            }
          } else {
            return null;
          }
        }

        function getStart(row) {
          let time;
          let roundedNum = Math.floor(+row[10])
          if (roundedNum > 0) {
            time = new Date(row[0]);
            time.setSeconds(time.getSeconds() + roundedNum);
          } else {
            time = new Date(row[0])
          }
          return time;
        }

        function getEnd(start, row) {
          let time = new Date(start);
          time.setSeconds(time.getSeconds() + Math.floor(+row[9]));
          if (agentArr.length > i + 1) {
            let lastTime = new Date(getStart(agentArr[i + 1]))
            if (time <= lastTime) {
              return time;
            } else {
              return "Dropped"
            }
          } else {
            return time;
          }
        }

        plotter(row)

        function plotter(row) {
          let tt = (new Date(new Date(start).setHours(0, 0, 0, 0)).getTime() / 1000)
          let curtime = (start.getTime() / 1000) - tt
          let startSpot = ((curtime - firstSeconds))
          let endSpot = startSpot + Math.round(row[9])
          if (endSpot > size) {
            endSpot = size
          }
          if ((startSpot < endSpot) && ((endSpot - startSpot) < 1200)) {
            callArr.fill(1, startSpot, endSpot)
          }
          //if has a wait time, plot the start on the init array
          let waitNum = Math.round(row[10])
          if (typeof waitNum === "number") {
            let ringTime = ((new Date(row[0]).getTime() / 1000) - tt - firstSeconds)
            initiatedArr.fill(1, ringTime, ringTime + 1)
          }
        }

        //each call run
        prevEndTime = end
        return row;
      })
      allUserArr.push(callArr)

      //single run
      //TODO: Failing due to NaN for duration, if duration is ridiculously long. Easiest solution is probably to change the originsl data, to strt of next call time, or by just removing that duration/row completely

      let durTotal;
      function totalDur() {
        let durationFilter = agentArr.filter((r) => {
          return ((+r[9] < 900) && (!r[10].includes("-")))
        })
        durationF = durationFilter.map((rw) => Math.round(rw[9])).flat();
        durTotal = durationF.reduce((acc, num) => acc + num, 0)
        let length = durationFilter.length;

        let res = (length > 1) ? Math.round(durTotal / length) : 0
        return res;
      }

      function avgResponse() {
        if (averageArr.length > 1) {
          return Math.round(averageArr.reduce((a, b) => +a + +b, 0) / averageArr.length)
        } else {
          return 0;
        }
      }

      let averageDuration = totalDur()
      let averageArr = agentArr.filter((row) => (!(row[16] === '') && +row[16] > 0.01)).map((row) => row[16]).flat()
      let averageResponse = avgResponse()

      let title = ["DateTime", "Desk", "Under 30", "Inbound", "Team", "Agent", "Access", "PARCS", "PARCS Access", "Duration", "Wait", "Facility", "Device",
        "Start Time", "End Time", "Time Between", "Call Response"]
      agentArr.unshift(title);
      sheet.getRange(1, 1, agentArr.length, agentArr[1].length).setValues(agentArr)

      removeEmptyRows(sheet)

      let dataRow = [activeDate, team];

      //fordevice %s
      devicePercents()

      function devicePercents() {
        let devices = onlyWaits.map((row) => row[12])
        let flatDevices = devices.flat()

        deviceChecker(flatDevices)

        function deviceChecker(arr) {
          let types = ["ENT", "POF", "PIL", "DRO"]
          deviceCheck(arr, types)

          function deviceCheck(arr, types) {
            let counts = [];

            for (let type of types) {
              let filtered = arr.filter((text) => {
                let beginning = text.split(" ")[0];
                return (beginning === type)
              })
              counts.push(filtered.length)
            }
            let total = counts.reduce((acc, num) => acc + num, 0)
            let countPercents = counts.map((type) => { return Math.round(((type / total) * 100)).toFixed(0) + '%' })

            dataRow.push(total, ...countPercents)
            deviceArr.push([activeDate, ...counts, total, ...countPercents])

          }
        }
      }


      function percentAccess() {
        let access = onlyWaits.filter((row) => {
          return (row[8].includes("Yes"))
        })
        let num = (Math.round((access.length / onlyWaits.length) * 100).toFixed(0) + '%');
        if (access.length > 0) {
          return num;
        } else {
          return 0
        }
      }
      let accessPercent = percentAccess()

      function getCherry() {
        return cherryArr.filter((row) => row[1].includes(activeDate)).length
      }
      let cherryCalls = getCherry()

      function getOnCallTime() {
        if (agentArr.length > 3) {
          //total seconds
          let end = new Date(agentArr[(agentArr.length - 1)][0]).getTime()
          let start = new Date(agentArr[1][0]).getTime()
          let startEndTime = (end - start) / 1000
          let percentOn = Math.round((durTotal / startEndTime) * 100).toFixed(0) + '%';
          console.log(end, start, startEndTime, percentOn)
          return percentOn
        } else {
          return 0;
        }
      }
      let percentOnCalls = getOnCallTime()

      //r
      dataRow.push(
        accessPercent,
        percentOnCalls,
        averageDuration,
        averageResponse,
        twoArr.filter((row) => row.includes(activeDate)).length,
        sixMinArr.filter((row) => row.includes(activeDate)).length,
        fiveFiveTwoArr.filter((row) => row.includes(activeDate)).length,
        cherryCalls)

      totalRow.push(dataRow)
      //agent ends here
    }
  }
  allUserArr.unshift(initiatedArr)
  console.log(allUserArr.length, allUserArr[0].length, allUserArr[1].length)

  let output = allUserArr[0].map((a, colIndex) => allUserArr.map(row => row[colIndex]));

  if (output[0].length > 25) {
    summarySheet.insertColumnsAfter(1, (output[0].length - 25))
  }
  let sumRange = summarySheet.getRange(1, 2, output.length, output[0].length)
  sumRange.setNumberFormat("0");
  sumRange.setValues(output);

  insertMultipleSheets()
  function insertMultipleSheets() {
    let i = 3;

    twoArr.unshift(["Time", "Device", "Length", "Name"])
    i++
    insertDataSheet(twoArr, "Small Calls", i);
    sixMinArr.unshift(["Time", "Device", "time", "name"])
    i++
    insertDataSheet(sixMinArr, "Six Min+", i);
    fiveFiveTwoArr.unshift(["name", "time"])
    i++
    insertDataSheet(fiveFiveTwoArr, "5:52 Calls", i);
    cherryArr.unshift(["Time", "Name", "Skipped Device", "Skipped System", "Skipped Access", "For Device", "For System", "For Access"])
    i++
    insertDataSheet(cherryArr, "Cherry Picked", i)
    deviceArr.unshift(["Name", "Entry", "POF", "PIL", "Door+OH", "Total", "Entry", "POF", "PIL", "Door+OH"])
    i++
    insertDataSheet(deviceArr, "Device Types", i)
  }


  function insertDataSheet(array, name, i) {
    if (array.length > 0) {
      let sheet = ss.insertSheet(name, i);
      sheet.getRange(1, 1, array.length, array[0].length).setValues(array)
    }
  }

  let theArray = [];
  theArray.push(Utilities.formatDate(new Date(firstDate), 'America/Mexico_City', 'yyyy-MM-DD'))
  theArray.push(firstTime);
  theArray.push(lastTime);
  theArray.push(url)

  totalRow.unshift(["Name", "Team", "# of Calls", "Entry", "POF", "PIL", "Door", "% Access", "% on Calls", "Avg Duration", "Avg Response", "<2 Sec Calls", ">6 Min Calls", "Idled", "Cherry Calls"])
  theArray.push(totalRow); // data stuff goes here

  let datatableSheet = ss.insertSheet("DataTable")
  datatableSheet.getRange(1, 1, totalRow.length, totalRow[0].length).setValues(totalRow)
  removeEmptyRows(datatableSheet)
  console.log(theArray)

  theArray.push(waitCharts)
  return theArray;


}