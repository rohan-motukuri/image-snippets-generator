var noteQueueFile = SpreadsheetApp.openById('1ME1MAHlg9W7fNOB3FZiVYsR_6dJUUXXtLKnFBQSnEkY');
const noteQueueSize = noteQueueFile.getLastRow();
const noteQueueValues = noteQueueFile.getSheetValues(1, 1, noteQueueSize, noteQueueFile.getLastColumn()).toString();

var infoSheetMap = SpreadsheetApp.openById('1SFeCw14pPvjwU4Bawdu79oZ3RFG0OM90TsasXBtmJOc');
const infoSheetMapVlaues = infoSheetMap.getSheetValues(2, 1, infoSheetMap.getLastRow() - 1, infoSheetMap.getLastColumn())

const midProcessFolder = DriveApp.getFolderById('134PmFwV0ngREkIX2-JzgfpabUynEqWtC');

var errList = ""
var pdfLinks = [], rawFileLinks = [], sysCalls = [], arrOfSnipInst = [];
var blobStack = {}, fileStack = {}
var callBack = {
  "FlowsForSnips": false
}

/* -------------Begin Notes Checking---------------- */
/*
  List of compromises - 
  - For now only one folder can be accessed && only link or id can be passed
  - Revision comments of files hasn't been coded yet
  - 
  - 
 */

function triggerCallNoteBuild() {
  interconnectionDecipher('Form/NoteBuild');
}

function triggerCalendarFlows() {
  var syncToken = PropertiesService.getUserProperties().getProperty('NB_FLOWS_SYNC_TOKEN');
  var page = Calendar.Events.list('dj8hjknsq82lim8l5v8jjfdvoo@group.calendar.google.com', { syncToken: syncToken });
  if (page.items && page.items.length > 0) {
    for (var i = 0, len = page.items.length; i < len; ++i) {
      var item = page.items[i];
    }
    syncToken = page.nextSyncToken;
  }
  PropertiesService.getUserProperties().setProperty('NB_FLOWS_SYNC_TOKEN', syncToken);
  if (item.status != 'cancelled' && item.summary == 'RunScripts') {
    interconnectionDecipher('Flows/NoteBuild', item.description)
  }
}

function interconnectionDecipher(funcCaller) {

  switch (funcCaller) {

    case 'Form/NoteBuild': {
      let responseSheet = SpreadsheetApp.openById('13TbIv-cE5LE_WDm88nOvcoqv_vgk2WeHXO9qyxgSyBI')
      let responseBody = responseSheet.getSheetValues(responseSheet.getLastRow(), 2, 1, 1)[0][0].toString()
      let formResponse = responseBody.split('||')
      let links = formResponse[0].split(',')
      let foldering, uniBlacklist, uniWhitelist

      if (formResponse[1])
        foldering = formResponse[1].split(',') // For now only one folder can be accessed && only link or id can be passed
      if (formResponse[2])
        uniBlacklist = formResponse[2].trim().split(',') // For now these apply to the one folder mentioned earlier
      if (formResponse[3])
        uniWhitelist = formResponse[3].trim().split(',') // || --

      let linkListsForCheck = [], folderLinks = []
      try {
        let linker
        if (links && links[0])
          for (let i = 0, len = links.length; i < len; ++i) {
            if (links[i].length == 1)
              linker = link[i]
            else
              linker = links[i].split('/')[5];
            if (linker)
              linkListsForCheck.push(linker)
            else
              errList += "The Link Is Broken: " + links[i] + "\n"
          }
        if (foldering && foldering[0])
          for (let i = 0, len = foldering.length; i < len; ++i) {
            if (foldering[i].split('/').length == 1)
              linker = foldering[i]
            else
              linker = foldering[i].split('/')[7];
            if (linker)
              folderLinks.push(linker)
            else
              errList += "The Link Is Broken: " + links[i] + "\n"
          }
      }
      catch (err) {
        errList += "The Error: " + err + "\n"
      }
      checkForRawInfoGiven(linkListsForCheck)
      if (foldering) {
        if (foldering.length > 1)
          errList += "Can't process more than one folder yet :(\n";
        else
          for (let i = 0, len = folderLinks.length; i < len; ++i)
            checkForRawInfoManualy(folderLinks[i], uniBlacklist + linkListsForCheck.toString(), uniWhitelist)
      }
      let snipDeterminer
      if (errList != "") // && listFromSearch != null
        GmailApp.sendEmail('20071a10a7@vnrvjiet.in',
          'Errors While Checking Note Building Files', errList +
          "\nRe-Check: https://docs.google.com/forms/d/e/1FAIpQLSeFcC_gUAVXUydg__n0LiJDhYlhpdzira_C5KILG0iVx61xXA/viewform?usp=pp_url&entry.1799046854=" + responseBody.replace(/\s/g, "+").replace('|', "%7C"))
      else try {
        if (pdfLinks.length) {
          addPDFsToRawArr();
        }
        snipDeterminer = getSnipDataFromPDF(getSnipsFromRaw(rawFileLinks))
        if (callBack.FlowsForSnips == true && !snipDeterminer) {
          throw ("No links provided for further processing\n")
        }
        else if (callBack.FlowsForSnips == false && arrOfSnipInst.length) {
          for (let i, len1 = arrOfSnipInst.length; i < len1; ++i) {

          }
          let instsList = []
          instsList = sortOnInstructions(filterInstArr(instsList))
          docCreater(instsList)
        }

        GmailApp.sendEmail('20071a10a7@vnrvjiet.in', 'No Errors While Checking Note Building Files', (snipDeterminer) ? 'Snips program called.' : null)
      }
      catch (err) {
        GmailApp.sendEmail('20071a10a7@vnrvjiet.in', 'Error While Calling Snipper', err + "\nImageLinks: " + rawFileLinks + "\nPdfLinks: " + pdfLinks)
      }
    } break;

    case 'Flows/NoteBuild': {

      try {

        var callBackLog = DriveApp.getFileById(arguments[1].replace(/\s/g, '')).getBlob().getDataAsString().split("|||")
        let implBuildInst = JSON.parse(callBackLog[0]),
          labelXID = JSON.parse("{" + callBackLog[1].substr(1) + "}")
        for (let i = 0, len = implBuildInst.length; i < len; ++i) {
          if (implBuildInst[i].rawInst.type == 'ImageFileName') {
            let id = labelXID[implBuildInst[i].rawInst.anchor]
            if (blobStack[id]) {
              implBuildInst[i].rawInst.reference = id
              implBuildInst[i].rawInst.anchor = blobStack[id]
            }
            else if (id) {
              while (true) {
                try {
                  implBuildInst[i].rawInst.reference = id
                  implBuildInst[i].rawInst.anchor = blobStack[id] = DriveApp.getFileById(id).getBlob()
                  break;
                }
                catch (err) {
                  Logger.log(err)
                }
              }
            }

            implBuildInst[i].rawInst.type = 'ImageFile'
          }
        }
        Logger.log(implBuildInst)
        docCreater(implBuildInst);
      }
      catch (err) {
        GmailApp.sendEmail('20071a10a7@vnrvjiet.in', 'Error while creating the document', err)
      }

    } break;

    default: Logger.log('Invalid Function Caller')

  }

  //checkForRawInfoManualy('1jKHvd1V0vo3Gd3o5Q8GLljIg4kcpvAjJ', "Recordings") // --> Instead of calling the function try to invoke this as an api so that total run time doesn't excede 30 sec hopefully.
}

function checkForRawInfoManualy(obj, blacklist, whitelist) {
  try {
    var subFolders = DriveApp.getFolderById(obj).getFolders()
    var subFiles = DriveApp.getFolderById(obj).getFiles()
  }
  catch (err) {
    errList += "Error While Processing 'checkForRawInfoManualy': " + err + "\n"
  }
  while (subFolders.hasNext()) {
    let currentFolder = subFolders.next()
    let fileProps = [currentFolder.getId(), currentFolder.getName()]
    if (blacklist == null || !((blacklist.indexOf(fileProps[1]) > -1 || blacklist.indexOf(fileProps[0]) > -1)))
      checkForRawInfoManualy(fileProps[0], blacklist, whitelist)
  }

  while (subFiles.hasNext()) {
    let currentFile = subFiles.next()
    let fileProps = [currentFile.getId(), currentFile.getName()]
    if (blacklist == null || !((blacklist.indexOf(fileProps[1]) > -1 || blacklist.indexOf(fileProps[0]) > -1))) {
      let commentsList = []
      let pageToken = null
      while (true) {
        let comments = Drive.Comments.list(fileProps[0], { "maxResults": 100, "pageToken": pageToken })
        for (let k = commentsList.length, j = 0, commentLen = comments.items.length; j < commentLen; ++k, ++j)
          commentsList[k] = comments.items[j]
        if (!(pageToken = comments.nextPageToken))
          break;
      }
      if (!commentsList.length)
        continue;
      if (!noteQueueValues.includes(fileProps[0]))

        if (currentFile.getMimeType().includes('pdf')) {
          let qualifier = processInsructs(commentsList, 'PDF/' + fileProps[1], 'DriveComments');
          if (qualifier != -1) {
            pdfLinks.push([fileProps[0], qualifier]);
            blacklist += ',' + fileProps[0];
            if (!fileStack[fileProps[0]]) {
              fileStack[fileProps[0]] = currentFile
            }
          }
          else {
            errList += "||Done isn't mentioned for the file : " + fileProps[1] + "\n"
          }
        }
        else if (currentFile.getMimeType().includes('image')) {
          let qualifier = processInsructs(commentsList, 'IMG/' + fileProps[1], 'DriveComments');
          if (qualifier != -1) {
            rawFileLinks.push([fileProps[0], qualifier]);
            blacklist += ',' + fileProps[0];
            if (!fileStack[fileProps[0]]) {
              fileStack[fileProps[0]] = currentFile
            }
          }
          else {
            errList += "||Done isn't mentioned for the file : " + fileProps[1] + "\n"
          }
        }

    }
  }
}

function checkForRawInfoGiven(list) {
  for (let i = 0, len = list.length; i < len; ++i) {
    try {
      let commentsList = []
      let pageToken = null
      while (true) {
        let comments = Drive.Comments.list(list[i], { "maxResults": 100, "pageToken": pageToken })
        for (let k = commentsList.length, j = 0; j < comments.items.length; k++ , j++)
          commentsList[k] = comments.items[j]
        if (!(pageToken = comments.nextPageToken))
          break;
      }
      if (noteQueueValues.includes(list[i])) {
        errList += "Processing Was Already Done For The File: " + list[i] + "\n"
      }
      else {

        let file = fileStack[list[i]]

        if (!fileStack[list[i]]) {
          fileStack[list[i]] = file = DriveApp.getFileById(list[i])
        }

        if (!blobStack[list[i]])
          blobStack[list[i]] = file.getBlob()

        let fileType = file.getMimeType()
        if (commentsList.length != 0) {
          if (fileType.includes('pdf')) {
            pdfLinks.push([list[i], processInsructs(commentsList, 'PDF/' + file.getName(), 'DriveComments')])
            if (pdfLinks[pdfLinks.length - 1][1] == -1)
              errList += "||Done isn't mentioned for the file : " + list[i] + "\n"
          }
          else if (fileType.includes('image')) {
            rawFileLinks.push([list[i], processInsructs(commentsList, 'IMG/' + file.getName(), 'DriveComments')])
            if (rawFileLinks[rawFileLinks.length - 1][1] == -1)
              errList += "||Done isn't mentioned for the file : " + list[i] + "\n"
          }
          else {
            errList += "File Type Not Supported For: " + list[i] + "\n" //Later allow for different types of raw info
          }
        }
      }
    }
    catch (err) {
      errList += "Error While Checking The File Of Id '" + list[i] + "': " + err + "\n"
      continue
    }
  }
}

/* -------------End Notes Checking---------------- */

/* -------------Begin Notes Snippets Generator---------------- */
/*
  Compromises: 
  - Labels don't mention the class timings and what not. Fetch data from other systems such as the timetable, number of units though tracker, learning tracker, events system (particularly exam event)
  x Multiple pdfs to be be processed together by merging them and sending them to the image extractor
  x Somehow remove comments and pass the pdf/image if possible 
  - Inter file communication -- For snippets to reference eachother even across diffrent files/images
  - Making each info block referenceable in the word doc
  - Make notes more flexible than just linear text file as in doc
*/

function getSnipsFromRaw(links) {
  if (!links.length)
    return null
  let pageCounter = 0
  const slidesId = '1eBEP1k6B0MZJHJCtk25vdwOnlr3B3QxrmbCuZ4tLmQk',
    snipsHolder = SlidesApp.openById(slidesId),
    pageDim = [450, 720]

  for (let i = 0, len = links.length; i < len; ++i) {

    let mainSlide = snipsHolder.appendSlide()
    let mainImg, imgDim, arrFinder = links[i][1]

    while (true) {
      try {
        if (typeof (links[i][0]) == 'string') {
          let file = fileStack[links[i][0]]
          if (!file)
            file = fileStack[links[i][0]] = DriveApp.getFileById(links[i][0])
          mainImg = mainSlide.insertImage(file.getThumbnail())
        }
        else
          mainImg = mainSlide.insertImage(links[i][0]);
        imgDim = [mainImg.getHeight(), mainImg.getWidth()];
        break;
      }
      catch (err) {
        Logger.log(err)
      }
    }

    for (let j = 0, lenn = arrOfSnipInst[arrFinder].length; j < lenn; ++j) {
      if (arrOfSnipInst[arrFinder][j].implInst.infoType != 'REGION') { // For when the context is text only
        let subSlide = snipsHolder.appendSlide(mainSlide);
        let anchr = JSON.parse(arrOfSnipInst[arrFinder][j].rawInst.anchor);
        let imgInst = null;
        let pageProp = (arrOfSnipInst[arrFinder][j].cnxtInst) ? arrOfSnipInst[arrFinder][j].cnxtInst.pageProp : null;
        if (arrOfSnipInst[arrFinder][j].cnxtInst && arrOfSnipInst[arrFinder][j].cnxtInst.cnxtType == 'IMGregion') {
          imgInst = arrOfSnipInst[arrFinder][j].cnxtInst.cnxtParams.whiteout
        }

        const c1 = (imgDim[1] / pageDim[1]), c2 = (imgDim[0] / pageDim[0])
        anchr = calcImgDims(anchr[1][1], 'mainAnchr');

        function calcImgDims(base, mode) {

          if (pageProp == true && mode == 'mainAnchr') {
            base[0] = 0;
            base[2] = c1;
          }
          else {
            base[0] = c1 * base[0];
            base[2] = c1 * base[2];
          }
          base[1] *= c2;
          base[3] *= c2;

          return base
        }

        if (imgInst) {
          whiteout(imgInst.unions, imgInst.shape, imgInst.color, subSlide, pageDim);
        }
        arrOfSnipInst[arrFinder][j].rawInst.anchor = [++pageCounter, [anchr[0], anchr[1], anchr[2], anchr[3]]]

        //  Insert Shapes For Better View
        /*
        let placeHolderShape = subSlide.insertShape(SlidesApp.ShapeType.RECTANGLE, anchr[0] * pageDim[1], anchr[1] * pageDim[0], (anchr[2] - anchr[0]) * pageDim[1], (anchr[3] - anchr[1]) * pageDim[0])
        placeHolderShape.getFill().setTransparent()
        placeHolderShape.getBorder().setWeight(4).getLineFill().setSolidFill('#ff0000')
        */
      }
    }
    mainSlide.remove()
  }

  snipsHolder.saveAndClose();

  let pptFile = DriveApp.getFileById(slidesId)

  let returnValue = [midProcessFolder.createFile(pptFile.getBlob().getAs('application/pdf')).setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW).getId(), pageDim]

  SlidesApp.openById(slidesId).getSlides().forEach((z) => { z.remove() });

  return returnValue;

  function whiteout(unions, shape, color, slide, pD) {
    shape = SlidesApp.ShapeType[shape]
    switch (color) {
      case 'WHITE': {
        color = "#ffffff"
      }
    }
    for (let i = 0, len = unions.length; i < len; ++i) {
      // nchr[1][1][0] * pageDim[1], anchr[1][1][1] * pageDim[0], (anchr[1][1][2] - anchr[1][1][0]) * pageDim[1], (anchr[1][1][3] - anchr[1][1][1]) * pageDim[0]
      unions[i] = calcImgDims(unions[i])
      let
        l = unions[i][0] * pD[1],
        t = unions[i][1] * pD[0],
        w = (unions[i][2] - unions[i][0]) * pD[1],
        h = (unions[i][3] - unions[i][1]) * pD[0];
      let whiteoutShape = slide.insertShape(shape, l, t, w, h)
      whiteoutShape.getFill().setSolidFill(color)
      whiteoutShape.getBorder().setTransparent()
    }
  }

}

function getSnipDataFromPDF(midFileData, mode) {
  class Calc {
    constructor(p, x1, y1, x2, y2, fD) {
      this.p = p
      this.x1 = x1 * fD[1]
      this.y1 = fD[0] * (1 - y2)
      this.x2 = x2 * fD[1]
      this.y2 = fD[0] * (1 - y1)
    }// fD = File Data
    log(label) {
      return ('{"Name": ' + '"' + label + '"' + ', "LowerLeftXCoordinate":  ' + this.x1 + ', "LowerLeftYCoordinate":  ' + this.y1 + ',"UpperRightXCoordinate": ' + this.x2 + ', "UpperRightYCoordinate": ' + this.y2 + ', "PageNumber": ' + this.p + ', "ImageType": "PNG", "ExtractEntirePage": false, "Resolution": 300}')
    }
  }

  if (!midFileData)
    return null
  let logTxt = midFileData[0], scanFileData = midFileData[1]
  const currentDate = new Date().getUTCDate().toString() + '_' + new Date().getUTCMonth().toString() + '_' + new Date().getUTCFullYear().toString() + '_' + new Date().getMinutes().toString()
  //name is to be decided by the anchoring and what not?

  for (let i = 0, arrSize1 = arrOfSnipInst.length; i < arrSize1; ++i) {
    if (arrOfSnipInst[i] != 'null') {
      for (let j = 0, arrSize2 = arrOfSnipInst[i].length; j < arrSize2; ++j) {
        if (arrOfSnipInst[i][j].implInst.infoType != 'REGION' && arrOfSnipInst[i][j].rawInst.type == 'ImageRegion') {
          let anchArr = arrOfSnipInst[i][j].rawInst.anchor,
            calcBase = new Calc(anchArr[0],
              anchArr[1][0],
              anchArr[1][1],
              anchArr[1][2],
              anchArr[1][3],
              scanFileData),
            label = currentDate + '_' + anchArr[0];
          logTxt += '|' + calcBase.log(label);
          arrOfSnipInst[i][j].label = label;
        }
      }
    }
  }

  let flowsInstructions = midProcessFolder.createFile('flowsInstructions.txt', logTxt, 'text/plain').getId()
  let sortedArr = []

  sortedArr = sortOnInstructions(filterInstArr(sortedArr));

  let scriptCallbackInst = midProcessFolder.createFile('sriptCallbackInst.txt', JSON.stringify(sortedArr), 'text/plain').getId()
  CalendarApp.getCalendarById('eqhtn51codv14m2hvvo1fioop0@group.calendar.google.com').createEvent('RunFlows', new Date('January 1, 2120'), new Date('January 1, 2120')).setDescription(flowsInstructions + " | " + scriptCallbackInst)

  return 1

}

function filterInstArr(arr) {
  for (let i = 0, arrSize1 = arrOfSnipInst.length; i < arrSize1; ++i) {
    for (let j = 0, arrSize2 = arrOfSnipInst[i].length; j < arrSize2; ++j) {
      if (arrOfSnipInst[i] != 'null' && arrOfSnipInst[i][j].implInst.infoType != 'REGION')
        arr.push({
          "rawInst": {
            "anchor": (arrOfSnipInst[i][j].rawInst.type == 'ImageRegion') ? arrOfSnipInst[i][j].label : arrOfSnipInst[i][j].rawInst.anchor,
            "type": (arrOfSnipInst[i][j].rawInst.type == 'ImageRegion') ? 'ImageFileName' : arrOfSnipInst[i][j].rawInst.type,
            "reference": arrOfSnipInst[i][j].rawInst.reference
          },
          "implInst": arrOfSnipInst[i][j].implInst
        })
    }
  }
  return arr
}

function sortOnInstructions(sortedArr) {
  let implStack = [], sourceStack = [] // Later add the ability to pre establish the priorities by pushing stuff into these stacks
  // Change the name to ISA+ or something ISA_p ("ISA_Plus") and the sorting has to be done by considering all the ISA+ and ISA- levels

  sortedArr.sort(function (a, b) {
    if (a.implInst.infoAbstract.source == b.implInst.infoAbstract.source) {
      if (a.implInst.infoAbstract.impl == b.implInst.infoAbstract.impl) {
        if (a.implInst.infoAbstract.hierarchy.ISA == b.implInst.infoAbstract.hierarchy.ISA) {
          let level = 0,
            aCurrentObjName = Object.keys(a.implInst.infoAbstract.hierarchy)[1],
            bCurrentObjName = Object.keys(b.implInst.infoAbstract.hierarchy)[1],
            aCurrentObj = a.implInst.infoAbstract.hierarchy[aCurrentObjName],
            bCurrentObj = b.implInst.infoAbstract.hierarchy[bCurrentObjName],
            prevA, prevB

          while (true) {
            if (level != 0) {
              aCurrentObjName = Object.keys(prevA)[0], bCurrentObjName = Object.keys(prevB)[0]
            }
            let aLvl = aCurrentObjName.replace(/\D/g, ''),
              bLvl = bCurrentObjName.replace(/\D/g, '');
            aCurrentObjName = aCurrentObjName.toUpperCase().replace(/[0-9]/g, '')
            bCurrentObjName = bCurrentObjName.toUpperCase().replace(/[0-9]/g, '')
            //Logger.log(aCurrentObjName + " " + bCurrentObjName + " " + aLvl + " " + bLvl)
            if (aCurrentObjName == bCurrentObjName) {
              if (aLvl == bLvl) {
                let aChild = Object.keys(aCurrentObj)[0], bChild = Object.keys(bCurrentObj)[0]
                //Logger.log(aChild + " " + bChild)
                if (!aChild && bChild)
                  return -1
                else if (aChild && !bChild)
                  return 1
                else if (!aChild && !bChild) {
                  if (aCurrentObjName == 'H') {
                    errList += "Heading with id : " + aLvl + " was declared 2 times or more\n"
                    return 0
                  }
                  else
                    return 0
                }
                else {
                  prevA = aCurrentObj
                  prevB = bCurrentObj
                  aCurrentObj = aCurrentObj[aChild]
                  bCurrentObj = bCurrentObj[bChild]
                  ++level;
                }
              }
              else if (aLvl > bLvl) {
                return 1
              }
              else {
                return -1
              }
            }
            else {
              //  Temporary measures
              if (aCurrentObjName == 'C' && bCurrentObjName == 'SH')
                return -1;
              else if (bCurrentObjName == 'C' && aCurrentObjName == 'SH')
                return 1;
              return 0;
              //  Write a system after thinking and gaining more ideas by exploring diff notes
              switch ('C') {
                case aCurrentObjName: if (level == 0) return -1
                case bCurrentObjName: if (level == 0) return 1
                default: switch ('D') {

                }
              }
            }
          }
        }
        else {
          if (a.implInst.infoAbstract.hierarchy.ISA.toUpperCase().includes('UNIT')
            && b.implInst.infoAbstract.hierarchy.ISA.toUpperCase().includes('UNIT')) {
            let aUnit = a.implInst.infoAbstract.hierarchy.ISA.replace(/\s/g, ''),
              bUnit = b.implInst.infoAbstract.hierarchy.ISA.replace(/\s/g, '')
            if (aUnit > bUnit)
              return 1;
            else if (aUnit < bUnit)
              return -1;
            else
              return 0;
          }
          else
            return 0
        }
      }
      else {
        if (!implStack.includes(a.implInst.infoAbstract.impl)) {
          implStack.push(a.implInst.infoAbstract.impl)
        }
        if (!implStack.includes(b.implInst.infoAbstract.impl)) {
          implStack.push(b.implInst.infoAbstract.impl)
        }
        if (implStack.indexOf(a.implInst.infoAbstract.impl) > implStack.indexOf(b.implInst.infoAbstract.impl))
          return 1
        else
          return -1
      }
    }
    else {
      if (!implStack.includes(a.implInst.infoAbstract.source)) {
        sourceStack.push(a.implInst.infoAbstract.source)
      }
      if (!implStack.includes(b.implInst.infoAbstract.source)) {
        sourceStack.push(b.implInst.infoAbstract.source)
      }
      if (sourceStack.indexOf(a.implInst.infoAbstract.source) > sourceStack.indexOf(b.implInst.infoAbstract.source))
        return 1
      else
        return -1
    }

  })

  return sortedArr
}

function processInsructs(arr, rawName, instSource) {
  let bound = [], returner = -1
  const listOfAvailableFileTypes = ['PDF', 'IMG']
  if (!instSource) instSource = 'DriveComments' // Safe gaurd !!!!!! Remove later
  if (!listOfAvailableFileTypes.includes(rawName.split('/')[0])) {
    returner = "FileTypeNotAvailable for : " + rawName.split('/')[1]
  }

  class ImplInst {
    constructor(implType, infoType, infoAbstract, implParams) {
      this.implType = 'GDoc'
      this.infoType = null
      this.infoAbstract = {
        'source': '',
        'impl': '',
        'hierarchy': {
          'ISA': ''
        },
        'date': ''
      } // [source, implFile, ISA, extra(date)] & for now put source = clg as default
      this.implParams = { //For Document type implementation
        'docName': null,
        'text': null,
        'textStyle': null,
        'textFormatting': null
        //,'inlinePrev' : true
      }
    }
  }

  class CnxtInst {
    constructor(cnxtType) {
      if (cnxtType)
        this.cnxtType = cnxtType
      if (cnxtType == 'IMGregion') {
        this.pageProp = false;
        this.cnxtParams = new IMGtype();
      }
    }
  }

  class IMGtype {
    constructor(rotation, recolor, whiteout) {
      this.rotation = null
      this.recolor = null
      this.whiteout = {
        'unions': [],
        'shape': 'RECTANGLE',
        'color': 'WHITE' // For now! Later, infer instruct form comments and/or infer from the image itself
      };
    }
  }

  for (let i = 0, len = arr.length; i < len; ++i) {
    arr[i] = {
      'rawInst': arr[i],
      'cnxtInst': null, // For the manipulation of images
      'implInst': new ImplInst()  // For manipulation of the created document 
    }
  }
  switch (instSource) {

    case 'DriveComments': {

      for (let i = 0, len = arr.length; i < len; ++i) {
        if (arr[i].rawInst.status == 'resolved') {
          arr.splice(i--, 1);
          --len;
        }
        else { // arr[i]
          arr[i].rawInst = {
            'anchor': arr[i].rawInst.anchor,
            'type': 'ImageRegion',
            'reference': arr[i].rawInst.commentId,
            'inst': arr[i].rawInst.content,
          }
        }
      } // Add only the required properties

      arr.sort(function (a, b) {
        let anchrA = JSON.parse(a.rawInst.anchor), anchrB = JSON.parse(b.rawInst.anchor)
        if (a.rawInst.anchor != null || b.rawInst.anchor != null) //  && (a.status != 'resolved' || b.status != 'resolved') && !a.content.startsWith('||')
        {
          if (anchrA[1][0] != null && anchrB[1][0] != null) {
            if (anchrA[1][0] < anchrB[1][0])
              return -1
            else if (anchrA[1][0] > anchrB[1][0])
              return 1
          }
          if (anchrA[1][1][1] > anchrB[1][1][1])
            return 1
          else
            return -1
        }
      }) // Sort the array based on physical order

      for (let i = 0, arrLen = arr.length; i < arrLen; ++i) {
        let content = arr[i].rawInst.inst.toUpperCase().replace(/\s/g, ''), instList = [];
        let pageProp = false
        if (content.startsWith('||')) {
          if (content.startsWith('||BOUNDOPEN')) {
            bound.push({
              'content': arr[i].rawInst.inst.substr(arr[i].rawInst.inst.indexOf(":") + 1),
              'anchor': arr[i].rawInst.anchor
            })
          }
          else if (content.startsWith('||BOUNDCLOSE')) {
            for (let j = 0, len = bound.length; j < len; ++j)
              if (bound[j].content == content.substr(content.indexOf(':') + 1)) { // Later have smarter code? Like, able to clode the bound by simply putting ISA or something instead of typing the whole thing?
                bound.splice(j--, 1);
                --len;
              }
          }
          else if (content.startsWith('||DONE')) {
            if (returner == -1)
              returner = 1
          }
          else if (content.startsWith('||REVISION')) {
            // Write code later
          }
          arr.splice(i--, 1)
          --arrLen;
        }
        else {
          // Apply bound parameters if applicable
          if (!arr[i].rawInst.inst.toString().includes('ISA'))
            for (let j = 0, len = bound.length; j < len; ++j) {
              //if(arr[i].content.includes(bound[j].content))
              arr[i].rawInst.inst += '\n' + bound[j].content
            }

          // Decipher and structurize the instructions from comments content
          instList = arr[i].rawInst.inst.split('\n')
          console.log(arr[i])
          for (let j = 0, instListLen = instList.length; j < instListLen; ++j) {
            
            let currentInst = instList[j].replace(/\s/g, '').toUpperCase().split(':'),
              currentInstGran = currentInst[0].split('')

            //Check for impltype and set it to the info block instructions and the temp var implTypeHolder

            let char = [], int = []
            for (let c = 0, granLen = currentInstGran.length; c < granLen; ++c) {
              if (isNaN(Number(currentInstGran[c])))
                char.push(currentInstGran[c])
              else
                int.push(currentInstGran[c])
            }
            switch (currentInst[0]) {

              case ('ISA'): {
                if (!arr[i].implInst.infoAbstract.hierarchy.ISA && currentInst[1].indexOf('UNIT') != -1 && (!isNaN(Number(currentInst[1].substr(currentInst[1].indexOf('UNIT') + 4, currentInst[1].length - currentInst[1].indexOf('UNIT') + 4))))) {
                  arr[i].implInst.infoAbstract.hierarchy.ISA = currentInst[1]
                }
                else {
                  errList += "ISA syntax error for comment Id:" + arr[i].rawInst.reference + "\n"
                }
                break
              }

              case ('SOURCE'): {
                arr[i].implInst.infoAbstract.source = currentInst[1]
                break
              }

              case ('IMPL'): { //Which Subject
                arr[i].implInst.infoAbstract.impl = currentInst[1]
                break
              }

              case ('DATE'): {
                arr[i].implInst.infoAbstract.date = currentInst[1]
                break
              }

              case ('PAGEPROP'): {
                if (currentInst[1]) {
                  let propedAnchor = JSON.parse(arr[i].rawInst.anchor);
                  if (currentInst[1] == 'FALSE') {
                    if (!arr[i].cnxtInst)
                      arr[i].cnxtInst = new CnxtInst('IMGregion');
                    arr[i].cnxtInst.pageProp = false
                  }
                  else if (currentInst[1] == 'TRUE') {
                    if (!arr[i].cnxtInst)
                      arr[i].cnxtInst = new CnxtInst('IMGregion');
                    arr[i].cnxtInst.pageProp = true
                  }
                  else {
                    errList += "PageProp parameter isn't valid for the comment :" + arr[i].rawInst.reference + "\n"
                  }
                }
                else {
                  errList += "PageProp parameter is empty for the comment :" + arr[i].rawInst.reference + "\n"
                }
              }

              case ('IMPLTYPE'): {
                arr[i].implInst.implType = currentInst[1];
                break;
              }

              default: {
                let postEval = [], currentEval, tObj = arr[i].implInst.infoAbstract.hierarchy
                let staticUnion, intersections = []
                function findUnion(key, mainUnion, mode) {
                  for (let l = 1, len = currentInst.length; l < len; ++l) {
                    if (arr[key].rawInst.inst.startsWith(currentInst[l])) {
                      if (!staticUnion) {
                        var staticAnchr = JSON.parse(mainUnion.rawInst.anchor)
                        staticAnchr = staticAnchr[1][1];
                      }
                      else
                        var staticAnchr = staticUnion;

                      let varAnchr = JSON.parse(arr[key].rawInst.anchor);
                      let
                        x1 = (staticAnchr[0] > varAnchr[1][1][0]) ? staticAnchr[0] : varAnchr[1][1][0],
                        x2 = (staticAnchr[2] < varAnchr[1][1][2]) ? staticAnchr[2] : varAnchr[1][1][2],
                        y1 = (staticAnchr[1] > varAnchr[1][1][1]) ? staticAnchr[1] : varAnchr[1][1][1],
                        y2 = (staticAnchr[3] < varAnchr[1][1][3]) ? staticAnchr[3] : varAnchr[1][1][3];
                      if (x1 > staticAnchr[2]
                        || x2 < staticAnchr[0]
                        || y1 > staticAnchr[3]
                        || y2 < staticAnchr[1]) {
                        errList += "The comment : " + arr[key].rawInst.reference + " doesn't form proper intersection with comment : " + mainUnion.rawInst.reference + "\n"
                      }
                      else {
                        if (mode == 'U')
                          staticUnion = [x1, y1, x2, y2];
                        else
                          intersections.push([x1, y1, x2, y2]);
                      }
                      continue;
                    }
                  } // Check all the current inst
                  if (staticUnion)
                    return staticUnion
                  return intersections
                }
                if (char[0] == '+' || char[0] == '-')
                  postEval.push(char[0])
                for (let c = (postEval.length) ? 1 : 0; c < char.length; c++) {
                  if (char[c] == 'H' && !currentEval) {
                    let hi = int.splice(0, 1)
                    if (!tObj['H' + hi]) {
                      currentEval = 'H'
                      tObj = tObj['H' + hi] = {};
                      if (char.length == 1)
                        arr[i].implInst.infoType = 'HEADING'
                    }
                    else
                      errList += 'Hierarchy syntax error for comment Id : '
                        + arr[i].rawInst.reference
                        + ' Hierarchy, 2 or more headings\n'
                  }
                  else if (currentEval == 'H') {

                    if (char[c] == 'S') {
                      if (char[++c] == 'H') {
                        let si = int.splice(0, 1)
                        tObj = tObj['SH' + si] = {}
                        if (c == char.length - 1)
                          arr[i].implInst.infoType = 'SUBHEADING'
                      }
                      else if (char[c] == 'Q' && char[c - 2] == 'Q') {

                      }
                      else if (char[c] == 'S') {

                      }
                      else {
                        errList += "Hierarchy syntax error for comment Id: " + arr[i].rawInst.reference + "\n"
                      }
                    }
                    else if (char[c] == 'C') {
                      let ci = int.splice(0, 1)
                      tObj = tObj['C' + ci] = {}
                      arr[i].implInst.infoType = 'CONCEPT'
                    }
                    else if (char[c] == 'E') {
                      let ei = int.splice(0, 1)
                      tObj = tObj['E' + ei] = {}
                      arr[i].implInst.infoType = 'EXPLAINATION'
                    }
                    else if (char[c] == 'D') {
                      let di = int.splice(0, 1)
                      tObj = tObj['D' + di] = {}
                      arr[i].implInst.infoType = 'DIAGRAM'
                    }
                    else if (char[c] == 'Q') {
                      let qi = int.splice(0, 1)
                      tObj = tObj['Q' + qi] = {}
                      arr[i].implInst.infoType = 'QUESTION'
                    }
                    else if (char[c] == 'A') {
                      let ai = int.splice(0, 1)
                      tObj = tObj['A' + ai] = {}
                      arr[i].implInst.infoType = 'ANSWER'
                    }

                  }
                  else if (!currentEval && char[c] == 'U') {
                    currentEval = 'I';

                    if (postEval[postEval.length - 1] == '-') {
                      for (let k = i; k < arrLen; ++k) {
                        findUnion(k, arr[i], 'U');
                      }

                      for (let k = i; k >= 0; --k) {
                        findUnion(k, arr[i], 'U');
                      }
                      if (!arr[i].cnxtInst)
                        arr[i].cnxtInst = new CnxtInst('IMGregion');
                      arr[i].cnxtInst.cnxtParams.whiteout.unions.push(staticUnion);
                    }
                  }
                  else if (!currentEval && char[c] == 'O') {
                    currentEval = 'I';
                    if (postEval[postEval.length - 1] == '-') {
                      if (!arr[i].cnxtInst)
                        arr[i].cnxtInst = new CnxtInst('IMGregion');
                      for (let k = i; k < arrLen; ++k) {
                        findUnion(k, arr[i], 'O');
                      }

                      for (let k = i; k >= 0; --k) {
                        findUnion(k, arr[i], 'O');
                      }
                      intersections.forEach((a) => {
                        arr[i].cnxtInst.cnxtParams.whiteout.unions.push(a);
                      })
                    }
                  }
                  else if (!currentEval && char[c] == 'R') {
                    let ri = int.splice(0, 1)

                    if (!tObj['R' + ri]) {
                      currentEval = 'R'
                      tObj = tObj['R' + ri] = {};
                      arr[i].implInst.infoType = 'REGION'
                    }
                    else
                      errList += 'Hierarchy syntax error for comment Id : '
                        + arr[i].rawInst.reference
                        + ' Hierarchy, 2 or more regions share same label\n'
                  }
                  else {
                    errList += 'Instructions Syntax Error for comment Id: ' + arr[i].rawInst.reference + "\n"
                  }
                  if (currentInst[1]) {
                    if (currentInst[1].indexOf('"') == 0) {
                      arr[i].rawInst.anchor = instList[j].split(":")[1].substring(2, instList[j].split(":")[1].lastIndexOf('"'))
                      arr[i].rawInst.type = "Text";
                    }
                    else if (currentInst[1].indexOf('//') == 0) {
                      arr[i].rawInst.type = "Extern_IMG";
                    }
                    else {
                      //errList += "Substitution type not available for comment : " + arr[i].rawInst.reference + "\n" // - Idk this things initial purpouse, But whilst debugging on 10th Feb 2023, I realized it's causing errors for the use case of regions. So I commented it
                      
                    }
                  }
                  if (currentEval != 'I' && currentInst[1]) {
                    if (currentInst[1].split('"')[1]) {
                      if (currentInst[1].startsWith('"+') || currentInst[1].startsWith('"-')) {
                        // Code for appending to the given image
                      }
                      else {
                        arr[i].rawInst.anchor = null;
                        arr[i].rawInst.inst = instList[j].replace('"', '')
                      }
                    }
                    else if (currentInst[1].split('/"')[1]) {
                      // Fetcht the url and append or replace the image
                    }
                  }
                }
              }
            }
          }

          // Check impl instructions
          if (!arr[i].implInst.infoAbstract.source)
            arr[i].implInst.infoAbstract.source = 'CLG' //Temp simple prediction
          if (!arr[i].implInst.infoAbstract.impl)
            errList += "Implimentation not mentioned for the comment id : " + arr[i].rawInst.reference + "\n"   //Temperorary, Later
          if (!arr[i].implInst.infoAbstract.hierarchy.ISA)                                                                 //send these through
            errList += "ISA not mentioned for the comment id : " + arr[i].rawInst.reference + "\n"              //predictorSys first
          for (let i = 0, lenT = infoSheetMapVlaues.length; i < lenT; ++i) {
            if (infoSheetMapVlaues[i][0].toString() == arr[i].implInst.infoAbstract.source
              && infoSheetMapVlaues[i][1].toString() == arr[i].implInst.infoAbstract.impl) {
              break;
            }

            if (i == lenT)
              errList += "The mentioned Implementation isn't available for comment : " + arr[i].rawInst.reference + "\n"
          }

          // Check the cnxt settings
          if (pageProp == true) {
            let propedAnchor = JSON.parse(arr[i].rawInst.anchor);

            if (!arr[i].cnxtInst)
              arr[i].cnxtInst = new CnxtInst('IMGregion');

            arr[i].cnxtInst.cnxtParams.whiteout.unions.push([0, 0, 1, propedAnchor[1][1][1]])
            arr[i].cnxtInst.cnxtParams.whiteout.unions.push([0, propedAnchor[1][1][3], 1, 1])
            arr[i].cnxtInst.cnxtParams.whiteout.unions.push([0, 0, propedAnchor[1][1][0], 1])
            arr[i].cnxtInst.cnxtParams.whiteout.unions.push([propedAnchor[1][1][2], 0, 1, 1])
          }

          if (arr[i].rawInst.type == 'ImageRegion' && !callBack.FlowsForSnips)
            callBack.FlowsForSnips = true;
        }
      }

    } break;

  }
  if (returner == -1) {
    return -1;
  }
  else if (returner != 1)
    return returner
  if (errList)
    return null;

  return (arrOfSnipInst.push(arr) - 1)

}

function addPDFsToRawArr() {
  //Logger.log(arrOfSnipInst)
  let counter = -1, thumbnailStack = [];
  for (let i = 0, len1 = pdfLinks.length; i < len1; ++i) {
    let cp, tempInstList = [];
    let rawFile = fileStack[pdfLinks[i][0]]

    if (!rawFile) {
      rawFile = fileStack[pdfLinks[i][0]] = DriveApp.getFileById(pdfLinks[i][0])
    }

    for (let j = 0, currentFileSnipInsts = arrOfSnipInst.splice(pdfLinks[i][1], 1, 'null')[0], len2 = currentFileSnipInsts.length; j < len2; ++j) {
      if (currentFileSnipInsts[j].rawInst.type == 'ImageRegion') {
        let anchr = JSON.parse(currentFileSnipInsts[j].rawInst.anchor)
        //Logger.log(tempInstList)
        if (cp != anchr[1][0] || j == len2 - 1) {
          if (tempInstList.length > 0 || len2 == 1) {
            if (j == len2 - 1) {
              cp = anchr[1][0]
              tempInstList.push(currentFileSnipInsts[j]);
            }
            rawFileLinks.push(['{' + ++counter + '}', arrOfSnipInst.push(tempInstList) - 1]);
            tempInstList = [];
            let tempFile = splitPdf(rawFile, (cp + 1) + '-' + (cp + 1));
            tempFile.setName(tempFile.getName().replace('.pdf', '') + counter);
            tempFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
            thumbnailStack.push(['{' + counter + '}', tempFile.getId()]);
          }
          if (j != len2 - 1) {
            cp = anchr[1][0];
            --j;
            continue;
          }
        }
        tempInstList.push(currentFileSnipInsts[j]);
      }
      else if (currentFileSnipInsts[j].rawInst.type == 'Text') {
        arrOfSnipInst.push([currentFileSnipInsts[j]])
      }
    }
  }
  thumbnailStack.forEach((a) => {
    rawFileLinks.forEach((c) => {
      if (c[0] == a[0]) {
        while (true) {
          a[1] = DriveApp.getFileById(a[1])
          c[0] = a[1].getThumbnail();
          if (c[0] != null)
            break;
        }
        a[1].setTrashed(true);
      }
    });
  })
}

function imageEditor() {

  class DocStructElement {
    constructor(reference, refType, refRange, boundRange, children = {}) {
      this.reference = reference
      this.refType = refType
      this.refRange = refRange
      this.boundRange = boundRange
      this.children = children
    }
  }

  //let thumbnail = splitPdf(DriveApp.getFileById('15Gxvxy6N9oZvQ-BjYua7fwE0fjvhASc3'), '1-1').getThumbnail(
  let file = Docs.Documents.get("1C6GdxKR0SneM6kPHXbyJCU1G2d5qU0VAVTuudxBvqpc").body.content
  //let doc = DocumentApp.openById("1DnLcp5Ny-dIo5lsrJh1xALsLIkmxGqpCvpxOmysiPPs").getBody().appendParagraph(file.toString())
  //Logger.log(file[2].table.tableRows[0].tableCells[0].content[3].paragraph)
  
  let currentLevel = 0, currentEval, struct = {}, boundEndIndex, parentObj = struct, levelObjSheet = {
    "1" : struct 
  }
  for(let i = 0, len = file.length; i < len; ++i){
    let childObj = new DocStructElement ();
    if(file[i].paragraph){
      switch(file[i].paragraph.paragraphStyle.namedStyleType){
        case "TITLE" : {
          childObj = getElements (file[i].paragraph.elements)
          currentEval = "ISA"
          break;
        }
        case "HEADING_1" : {
          currentEval = "HEADING"
          break;
        }
        default : {
          if(file[i].paragraph.paragraphStyle.namedStyleType.startsWith('HEADING')){
            currentEval = 'SUBHEADING'
          }
          else{
            let inLineCondition = false,
            textRunS = file[i].paragraph.elements;
            for(let j = 0, len2 = file[i].paragraph.elements.length; j < len2; ++j) {
              if(!currentEval) {
                if(file[i].paragraph.elements[j].inlineObjectElement){
                  if (inLineCondition){
                    currentEval = 'CONTENT';
                    break;
                  }
                  inLineCondition = true;
                  continue;
                }
                if (file[i].paragraph.paragraphStyle.alignment == 'CENTER' &&
                    textRunS[j].textRun.textStyle.bold &&
                    textRunS[j].textRun.textStyle.underline &&
                    (inLineCondition || (textRunS[j].textRun.textStyle.fontSize && textRunS[j].textRun.textStyle.fontSize.magnitude == 16))){
                  currentEval = 'ISA'
                }
                else if(textRunS[j].textRun.textStyle.underline &&
                        (inLineCondition || (textRunS[j].textRun.textStyle.fontSize && textRunS[j].textRun.textStyle.fontSize.magnitude == 11))){
                  currentEval = 'SUBHEADING'
                }
                else if(textRunS[j].textRun.textStyle.bold &&
                        (inLineCondition || (textRunS[j].textRun.textStyle.fontSize && textRunS[j].textRun.textStyle.fontSize.magnitude == 11))){
                  currentEval = 'HEADING'
                }
                else if(!(textRunS[j].textRun.content == "\n" ||
                        textRunS[j].textRun.content.replace(" ", "") == "")){
                  currentEval = 'CONTENT'
                  break;
                }
              }
              else {
                if (textRunS[j].textRun.content == "\n" ||
                    textRunS[j].textRun.content.replace(" ", "") == ""){

                }
                else{
                  currentEval = "CONTENT";
                  break;
                }
              }
            }
            if(inLineCondition && !currentEval)
              currentEval = "CONTENT"
          }
        }
      }
    }
    else if(file[i].table){
      let caption
      if(file[i].table.rows == 1 && file[i].table.columns == 1){
        let collapseH = file[i].table.tableRows[0].tableCells[0]
        for(let j = 0, len2 = collapseH.content.length; j < len2; ++j){
          for(let k = 0, len3 = collapseH.content[j].paragraph.elements.length; k < len3; ++k){
            if(collapseH.content[j].paragraph.elements[k].inlineObjectElement){
              currentEval = "DIAGRAM"
            }
            if(collapseH.content[j].paragraph.elements[k].textRun.content != '\n' && 
              collapseH.content[j].paragraph.elements[k].textRun.content.replace(" ", "") != ''){
              if(currentEval == 'DIAGRAM'){
                if(caption)
                  currentEval = 'CONTENT'
                else
                  caption = collapseH.content[j].paragraph.elements[k].textRun.content
              }
              else{
                currentEval = "CONTENT"
              }
            }
          }
        }
      }
    }
    if(currentEval){
      let parentObjKeys = Object.getOwnPropertyNames(parentObj), l = 0
      for(let j = 0, len2 = parentObjKeys.length; j < len2; ++j){
        if(parentObjKeys[j].startsWith(currentEval)){
          ++l;
        }
      }
      parentObj[currentEval + l] = new DocStructElement ()
      if(currentEval != 'CONTENT'){
        increase(currentLevel);
      }
    }

    Logger.log(currentEval)
    currentEval = null;
  }
  
  function increase () {

  }

  function decrease () {

  }

  function getElements (obj) {
    obj
  }
}

function docCreater(implInsts) {

  class DocStructElement {
    constructor() {
      this.content = ""
      this.type = ""
      this.range = ""
      this.children = {}
    }
  }
  let docStack = {}, currentDoc, docStructs = {}, currentDocStruct = {}
  /*
    Compromises: 
    - For now assume the file to append the info to already exists, Creation of one through instructions can be handled later.
  */
  for (let i = 0, lenM = implInsts.length; i < lenM; ++i) {

    if (implInsts[i].implInst.implType == 'GDoc') {

      // Fetch the required impl doc from "file archiving sheet" // Later, merge this with the concept of information map, etc
      for (let j = 0, len = infoSheetMapVlaues.length; j < len; ++j) {

        if (infoSheetMapVlaues[j][0].toString() == implInsts[i].implInst.infoAbstract.source
          && infoSheetMapVlaues[j][1].toString() == implInsts[i].implInst.infoAbstract.impl) {
          implInsts[i].rawInst.reference = implInsts[i].rawInst.anchor;
          implInsts[i].rawInst.type = 'DocumentId';
          implInsts[i].rawInst.anchor = docStack[infoSheetMapVlaues[j][2]]
          break;
        }

      }

      // Work with the current doc from here
      if(!docStack[implInsts[i].rawInst.anchor]){
        docStack[implInsts[i].rawInst.anchor] = getDocStruct (Docs.Documents.get(implInsts[i].rawInst.anchor));
      }

      // Make the resource required to be pushed into batchUpdate

    }

  }


  function getDocStruct (docFile = Docs.Documents.get('dasfa')) {
    // Remove later
    let elementStacks = [], struct = {}, currLevel = null;
    for(let i = 0, len = docFile.body.content.length; i < len; ++i){
      if(docFile.body.content[i].paragraph){
        let para = docFile.body.content[i].paragraph
        let currentEval; 

        for(let j = 0, len = para.elements.length; j < len; ++j){
          if(para.elements[j].inlineObjectElement){

          }
          else if(para.elements[j].textRun){
            if(!currentEval){
              
              if( para.paragraphStyle.alignment == "CENTER" && 
                  para.elements[j].textRun.textStyle.bold == true &&
                  para.elements[j].textRun.textStyle.underline == true &&
                  para.elements[j].textRun.textStyle.fontSize.magnitude == 16){
                  currentEval = 'ISA';
              }

            }
            else {
              switch (para.elements[j].textRun.content){
                case "\n" : break;
                case " " : break;
                case "" : break;
                default : currentEval = null;
              }
            }
          }
        }
        if (para.paragraphStyle.namedStyleType == "TITLE"){
          currentEval = "ISA"
        }
        else if (para.paragraphStyle.namedStyleType.startsWith('HEADING')) {
          currentEval = 'HEADING'
        }
      
      }
    }
  }

}

/* -------------End Notes Snippets Generator---------------- */


/* -------------Start External Tools------------------------- */


/**
 * Splits a PDF file and returns the requested pages.
 *
 * @param {File} pdf the file to split
 * @param {string} pageRange1 a range of pages to include, e.g., '2-4'
 * @param {string} opt_pageRange2 another range of pages to include; add as many as you like
 *
 * return {File} the requested pages in a new file
 */

function splitPdf(pdf, pageRange1, opt_pageRange2) {

  var bytes = pdf.getBlob().getBytes();
  var name = pdf.getBlob().getName();
  var xrefByteOffset = '';
  var byteIndex = bytes.length - 1;
  while (!/\sstartxref\s/.test(xrefByteOffset)) {

    xrefByteOffset = String.fromCharCode(bytes[byteIndex]) + xrefByteOffset;
    byteIndex--;

  }
  xrefByteOffset = +(/\s\d+\s/.exec(xrefByteOffset)[0]);
  var objectByteOffsets = [];
  var trailerDictionary = '';
  var rootAddress = '';
  do {

    var xrefTable = '';
    var trailerEndByteOffset = byteIndex;
    byteIndex = xrefByteOffset;
    for (byteIndex; byteIndex <= trailerEndByteOffset; ++byteIndex) {

      xrefTable = xrefTable + String.fromCharCode(bytes[byteIndex]);

    }
    xrefTable = xrefTable.split(/\s*trailer\s*/);
    trailerDictionary = xrefTable[1];
    if (objectByteOffsets.length < 1) {

      rootAddress = /\d+\s+\d+\s+R/.exec(/\/Root\s*\d+\s+\d+\s+R/.exec(trailerDictionary)[0])[0].replace('R', 'obj');

    }
    xrefTable = xrefTable[0].split('\n');
    xrefTable.shift();
    while (xrefTable.length > 0) {

      var xrefSectionHeader = xrefTable.shift().split(/\s+/);
      var objectNumber = +xrefSectionHeader[0];
      var numberObjects = +xrefSectionHeader[1];
      for (var entryIndex = 0; entryIndex < numberObjects; entryIndex++) {

        var entry = xrefTable.shift().split(/\s+/);
        objectByteOffsets.push([[objectNumber, +entry[1], 'obj'], +entry[0]]);
        objectNumber++;

      }

    }
    if (/\s*\/Prev/.test(trailerDictionary)) {

      xrefByteOffset = +(/\s*\d+\s*/.exec(/\s*\/Prev\s*\d+\s*/.exec(trailerDictionary)[0])[0]);

    }

  } while (/\s*\/Prev/.test(trailerDictionary));
  var rootObject = getObject(rootAddress, objectByteOffsets, bytes);
  var pagesAddress = /\d+\s+\d+\s+R/.exec(/\/Pages\s*\d+\s+\d+\s+R/.exec(rootObject)[0])[0].replace('R', 'obj');
  var pagesObject = getObject(pagesAddress, objectByteOffsets, bytes);
  var pagesString = pagesObject.slice(pagesObject.indexOf(/\/Kids\s*\[/.exec(pagesObject)[0]));
  pagesString = pagesString.slice(pagesString.indexOf('[') + 1, pagesString.indexOf(']'));
  var pageAddresses = [];
  while (/\d+\s+\d+\s+R/.test(pagesString)) {

    var pageAddress = /\d+\s+\d+\s+R/.exec(pagesString)[0];
    pageAddresses.push(pageAddress.replace('R', 'obj'));
    pagesString = pagesString.slice(pageAddress.length + 1);

  }
  var newObjects = ['1 0 obj\r\n<</Type/Catalog/Pages 2 0 R >>\r\nendobj'];
  var objects = [];
  for (var argumentIndex = 1; argumentIndex < arguments.length; argumentIndex++) {

    var range = arguments[argumentIndex];
    if (range.indexOf('-') > -1) {

      var begin = +range.slice(0, range.indexOf('-'));
      var end = range.slice(range.indexOf('-') + 1, range.length);
      if (end == 'end') {

        for (var pageNumber = begin; pageNumber <= pageAddresses.length; pageNumber++) {

          var pageAddress = pageAddresses[pageNumber - 1];
          var page = getObject(pageAddress, objectByteOffsets, bytes);
          objects.push(page);
          objects = objects.concat(getDependencies(page, objectByteOffsets, bytes));

        }

      } else {

        end = +end;
        for (var pageNumber = begin; pageNumber <= end; pageNumber++) {

          var pageAddress = pageAddresses[pageNumber - 1];
          var page = getObject(pageAddress, objectByteOffsets, bytes);
          objects.push(page);
          objects = objects.concat(getDependencies(page, objectByteOffsets, bytes));

        }

      }

    } else {

      range = +range;
      var pageAddress = pageAddresses[range - 1];
      var page = getObject(pageAddress, objectByteOffsets, bytes);
      objects.push(page);
      objects = objects.concat(getDependencies(page, objectByteOffsets, bytes));

    }

  }
  pageAddresses = [];
  for (var objectIndex = 0; objectIndex < objects.length; objectIndex++) {

    var newObjectAddress = [(newObjects.length + 3) + '', 0 + '', 'obj'];
    if (!Array.isArray(objects[objectIndex])) {

      objects[objectIndex] = [objects[objectIndex]];

    }
    objects[objectIndex].unshift(newObjectAddress);
    var objectAddress = objects[objectIndex][1].match(/\d+\s+\d+\s+obj/)[0].split(/\s+/);
    objects[objectIndex].splice(1, 0, objectAddress);//here
    if (/\/Type\s*\/Page[^s]/.test(objects[objectIndex][2])) {

      objects[objectIndex][2] = objects[objectIndex][2].replace(/\/Parent\s*\d+\s+\d+\s+R/.exec(objects[objectIndex][2])[0], '/Parent 2 0 R');
      pageAddresses.push(newObjectAddress.join(' ').replace('obj', 'R'));

    }
    var addressRegExp = new RegExp(objectAddress[0] + '\\s+' + objectAddress[1] + '\\s+' + 'obj');
    objects[objectIndex][2] = objects[objectIndex][2].replace(addressRegExp.exec(objects[objectIndex][2])[0], newObjectAddress.join(' '));
    newObjects.push(objects[objectIndex]);

  }
  for (var referencingObjectIndex = 0; referencingObjectIndex < newObjects.length; referencingObjectIndex++) {

    var references = newObjects[referencingObjectIndex][2].match(/\d+\s+\d+\s+R/g);
    if (references != null) {

      var string = newObjects[referencingObjectIndex][2];
      var referenceIndices = [];
      var currentIndex = 0;
      for (var referenceIndex = 0; referenceIndex < references.length; referenceIndex++) {

        referenceIndices.push([]);
        referenceIndices[referenceIndex].push(string.slice(currentIndex).indexOf(references[referenceIndex]) + currentIndex);
        referenceIndices[referenceIndex].push(references[referenceIndex].length);
        currentIndex += string.slice(currentIndex).indexOf(references[referenceIndex]);

      }
      for (var referenceIndex = 0; referenceIndex < references.length; referenceIndex++) {

        var objectAddress = references[referenceIndex].replace('R', 'obj').split(/\s+/);
        for (var objectIndex = 0; objectIndex < newObjects.length; objectIndex++) {

          if (arrayEquals(objectAddress, newObjects[objectIndex][1])) {

            var length = string.length;
            newObjects[referencingObjectIndex][2] = string.slice(0, referenceIndices[referenceIndex][0]) + newObjects[objectIndex][0].join(' ').replace('obj', 'R') +
              string.slice(referenceIndices[referenceIndex][0] + referenceIndices[referenceIndex][1]);
            string = newObjects[referencingObjectIndex][2];
            var newLength = string.length;
            if (!(length == newLength)) {

              for (var subsequentReferenceIndex = referenceIndex + 1; subsequentReferenceIndex < references.length; subsequentReferenceIndex++) {

                referenceIndices[subsequentReferenceIndex][0] += (newLength - length);

              }

            }
            break;

          }

        }

      }

    }

  }
  for (var objectIndex = 0; objectIndex < newObjects.length; objectIndex++) {

    if (Array.isArray(newObjects[objectIndex])) {

      if (newObjects[objectIndex][3] != undefined) {

        newObjects[objectIndex] = newObjects[objectIndex].slice(2);

      } else {

        newObjects[objectIndex] = newObjects[objectIndex][2];

      }

    }

  }
  newObjects.splice(1, 0, '2 0 obj\r\n<</Type/Pages/Count ' + pageAddresses.length + ' /Kids [' + pageAddresses.join(' ') + ' ]>>\r\nendobj');
  newObjects.splice(2, 0, '3 0 obj\r\n<</Title (' + name + ') /Producer (PdfManipulation.splitPdf\\(\\), a Google Apps Script project by Jarom Luker \\(pricebook@hbboys.com\\)) /CreationDate (D' +
    Utilities.formatDate(new Date(), CalendarApp.getDefaultCalendar().getTimeZone(), 'yyyyMMddHHmmssZ').slice(0, -2) + "'00) /ModDate (D" + Utilities.formatDate(new Date(),
      CalendarApp.getDefaultCalendar().getTimeZone(), 'yyyyMMddHHmmssZ').slice(0, -2) + "'00)>>\r\nendobj");
  var byteOffsets = [0];
  var bytes = [];
  var header = '%PDF-1.3\r\n';
  for (var headerIndex = 0; headerIndex < header.length; headerIndex++) {

    bytes.push(header.charCodeAt(headerIndex));

  }
  bytes.push('%'.charCodeAt(0));
  for (var characterCode = -127; characterCode < -123; characterCode++) {

    bytes.push(characterCode);

  }
  bytes.push('\r'.charCodeAt(0));
  bytes.push('\n'.charCodeAt(0));
  while (newObjects.length > 0) {

    byteOffsets.push(bytes.length);
    var object = newObjects.shift();
    if (Array.isArray(object)) {

      var streamKeyword = /stream\s*\n/.exec(object[0])[0];
      if (streamKeyword.indexOf('\n\n') > streamKeyword.length - 3) {

        streamKeyword = streamKeyword.slice(0, -1);

      } else if (streamKeyword.indexOf('\r\n\r\n') > streamKeyword.length - 5) {

        streamKeyword = streamKeyword.slice(0, -2);

      }
      var streamIndex = object[0].indexOf(streamKeyword) + streamKeyword.length;
      for (var objectIndex = 0; objectIndex < streamIndex; objectIndex++) {

        bytes.push(object[0].charCodeAt(objectIndex))

      }
      bytes = bytes.concat(object[1]);
      for (var objectIndex = streamIndex; objectIndex < object[0].length; objectIndex++) {

        bytes.push(object[0].charCodeAt(objectIndex));

      }

    } else {

      for (var objectIndex = 0; objectIndex < object.length; objectIndex++) {

        bytes.push(object.charCodeAt(objectIndex));

      }

    }
    bytes.push('\r'.charCodeAt(0));
    bytes.push('\n'.charCodeAt(0));

  }
  var xrefByteOffset = bytes.length;
  var xrefHeader = 'xref\r\n';
  for (var xrefHeaderIndex = 0; xrefHeaderIndex < xrefHeader.length; xrefHeaderIndex++) {

    bytes.push(xrefHeader.charCodeAt(xrefHeaderIndex));

  }
  var xrefSectionHeader = '0 ' + byteOffsets.length + '\r\n';
  for (var xrefSectionHeaderIndex = 0; xrefSectionHeaderIndex < xrefSectionHeader.length; xrefSectionHeaderIndex++) {

    bytes.push(xrefSectionHeader.charCodeAt(xrefSectionHeaderIndex));

  }
  for (var byteOffsetIndex = 0; byteOffsetIndex < byteOffsets.length; byteOffsetIndex++) {

    for (var byteOffsetStringIndex = 0; byteOffsetStringIndex < 10; byteOffsetStringIndex++) {

      bytes.push(Utilities.formatString('%010d', byteOffsets[byteOffsetIndex]).charCodeAt(byteOffsetStringIndex));

    }
    bytes.push(' '.charCodeAt(0));
    if (byteOffsetIndex == 0) {

      for (var generationStringIndex = 0; generationStringIndex < 5; generationStringIndex++) {

        bytes.push('65535'.charCodeAt(generationStringIndex));

      }
      for (var keywordIndex = 0; keywordIndex < 2; keywordIndex++) {

        bytes.push(' f'.charCodeAt(keywordIndex));

      }

    } else {

      for (var generationStringIndex = 0; generationStringIndex < 5; generationStringIndex++) {

        bytes.push('0'.charCodeAt(0));

      }
      for (var keywordIndex = 0; keywordIndex < 2; keywordIndex++) {

        bytes.push(' n'.charCodeAt(keywordIndex));

      }

    }
    bytes.push('\r'.charCodeAt(0));
    bytes.push('\n'.charCodeAt(0));

  }
  for (var trailerHeaderIndex = 0; trailerHeaderIndex < 9; trailerHeaderIndex++) {

    bytes.push('trailer\r\n'.charCodeAt(trailerHeaderIndex));

  }
  var idBytes = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, (new Date).toString());
  var id = '';
  for (var idByteIndex = 0; idByteIndex < idBytes.length; idByteIndex++) {

    id = id + ('0' + (idBytes[idByteIndex] & 0xFF).toString(16)).slice(-2);

  }
  var trailer = '<</Size ' + (byteOffsets.length) + ' /Root 1 0 R /Info 2 0 R /ID [<' + id + '> <' + id + '>]>>\r\nstartxref\r\n' + xrefByteOffset + '\r\n%%EOF';
  for (var trailerIndex = 0; trailerIndex < trailer.length; trailerIndex++) {

    bytes.push(trailer.charCodeAt(trailerIndex));

  }
  var directory = midProcessFolder;
  return directory.createFile(Utilities.newBlob(bytes, 'application/pdf', name));
  function getObject(objectAddress, objectByteOffsets, bytes) {

    objectAddress = objectAddress.split(/\s+/);
    for (var addressIndex = 0; addressIndex < 2; addressIndex++) {

      objectAddress[addressIndex] = +objectAddress[addressIndex];

    }
    var object = [];
    var byteIndex = 0;
    for (var index in objectByteOffsets) {

      var offset = objectByteOffsets[index];
      if (arrayEquals(objectAddress, offset[0])) {

        byteIndex = offset[1];
        break;

      }

    }
    object.push('');
    while (object[0].indexOf('endobj') <= -1) {

      if (/stream\s*\n/.test(object[0])) {

        var streamLength;
        var lengthFinder = object[0].slice(object[0].indexOf(/\/Length/.exec(object[0])[0]));
        if (/\/Length\s*\d+\s+\d+\s+R/.test(lengthFinder)) {

          var lengthObjectAddress = /\d+\s+\d+\s+R/.exec(/\/Length\s*\d+\s+\d+\s+R/.exec(lengthFinder)[0])[0].split(/\s+/);
          lengthObjectAddress[2] = 'obj';
          for (var addressIndex = 0; addressIndex < 2; addressIndex++) {

            lengthObjectAddress[addressIndex] = +lengthObjectAddress[addressIndex];

          }
          var lengthObject = ''
          var lengthByteIndex = 0;
          for (var index in objectByteOffsets) {

            var offset = objectByteOffsets[index];
            if (arrayEquals(lengthObjectAddress, offset[0])) {

              lengthByteIndex = offset[1];
              break;

            }

          }
          while (lengthObject.indexOf('endobj') <= -1) {

            lengthObject = lengthObject + String.fromCharCode(bytes[lengthByteIndex]);
            lengthByteIndex++;

          }
          streamLength = +(lengthObject.match(/obj\s*\n\s*\d+\s*\n\s*endobj/)[0].match(/\d+/)[0]);

        } else {

          streamLength = +(/\d+/.exec(lengthFinder)[0]);

        }
        var streamBytes = bytes.slice(byteIndex, byteIndex + streamLength);
        object.push(streamBytes);
        byteIndex += streamLength;
        while (object[0].indexOf('endobj') <= -1) {

          object[0] = object[0] + String.fromCharCode(bytes[byteIndex]);
          byteIndex++;

        }
        return object;

      }
      object[0] = object[0] + String.fromCharCode(bytes[byteIndex]);
      byteIndex++;

    }
    return object[0];

  }
  function arrayEquals(array1, array2) {

    if (array1 == array2) {

      return true;

    }
    if (array1 == null && array2 == null) {

      return true;

    } else if (array1 == null || array2 == null) {

      return false;

    }
    if (array1.length != array2.length) {

      return false;

    }
    for (var index = 0; index < array1.length; index++) {

      if (Array.isArray(array1[index])) {

        if (!arrayEquals(array1[index], array2[index])) {

          return false;

        }
        continue;

      }
      if (array1[index] != array2[index]) {

        return false;

      }

    }
    return true;

  }
  function getDependencies(objectString, objectByteOffsets, bytes) {

    var dependencies = [];
    var references = objectString.match(/\d+\s+\d+\s+R/g);
    if (references != null) {

      while (references.length > 0) {

        if (/\/Parent/.test(objectString.slice(objectString.indexOf(references[0]) - 8, objectString.indexOf(references[0])))) {

          references.shift();
          continue;

        }
        var dependency = getObject(references.shift().replace('R', 'obj'), objectByteOffsets, bytes);
        var dependencyExists = false;
        for (var index in dependencies) {

          var entry = dependencies[index];
          dependencyExists = (arrayEquals(dependency, entry)) ? true : dependencyExists;

        }
        if (!dependencyExists) {

          dependencies.push(dependency);

        }
        if (Array.isArray(dependency)) {

          dependencies = dependencies.concat(getDependencies(dependency[0], objectByteOffsets, bytes));

        } else {

          dependencies = dependencies.concat(getDependencies(dependency, objectByteOffsets, bytes));

        }

      }

    }
    return dependencies;

  }

}

/* -------------End External Tools------------------------- */
