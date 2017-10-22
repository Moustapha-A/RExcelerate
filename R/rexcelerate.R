devtools::use_package(XLConnect)
devtools::use_package(pander)


RExcelerate = function(rexcelerateIO, FUN, ...){

  #connect to the file
  workbook = loadWorkbook(filename = rexcelerateIO$file, password = rexcelerateIO$password)

  #read worksheet
  worksheet = readWorksheet(workbook, sheet = rexcelerateIO$sheet, startRow = rexcelerateIO$startRow, endRow = rexcelerateIO$endRow,startCol = rexcelerateIO$startCol, endCol = rexcelerateIO$endCol, region = rexcelerateIO$region)

  #get passed Function arguments as list
  #as.list(match.call()) returns this function signitureas a list including name [REXcelerate], first arg [rexcelerateIO],
  # and second arg [FUN] and the n the function args [...] we are intrested only in the function args [...] so we removed
  # the first three items using [-(1:3)] below
  passedArgs = as.list(match.call())[-(1:3)]

  #Initize the results matrix
  results = matrix(ncol = length(rexcelerateIO$resultCols),nrow = NROW(worksheet))
  colnames(results) = rexcelerateIO$resultCols

  #loop on the rows of the worksheet to replace the RExc arguments with thier corresponding value in the sheet
  for(i in seq(1:NROW(worksheet))){

    # Get the sheet value of each RExc argument
    arglist = lapply(passedArgs,function(x){
      if(is.character(x) && grepl('^rexc:',x)){
        x=gsub('rexc:',"",x)
        return(worksheet[i,x])
      }
      return(x)
    })

    #Apply the function and stack the results
    result = do.call(FUN, arglist)
    results[i,]=result
  }

  #write the results to the sheet

  # startCol, startRow ==> the position of writing data
  if(rexcelerateIO$startCol==0 && rexcelerateIO$endCol==0 && rexcelerateIO$startRow==0 && rexcelerateIO$endRow==0){ # The sheet bounderies are set automatically
    startCol = getBoundingBox(workbook,sheet = rexcelerateIO$sheet)[4,]+1 #Next to the Last Column in the sheet
    startRow = getBoundingBox(workbook,sheet = rexcelerateIO$sheet)[1,] #starting as the first Row in the sheet
  }
  else{ # The sheet boundaries are set by the user
    startCol = rexcelerateIO$endCol+1
    startRow =rexcelerateIO$startRow
  }

  writeWorksheet(workbook,data= data.frame(results), sheet=rexcelerateIO$sheet, startCol = startCol, startRow = startRow)
  saveWorkbook(workbook)

  if(rexcelerateIO$openWhenDone==TRUE){
    openFileInOS(rexcelerateIO$file)
  }

  return(results)

}

RExcelerateIO = function(excelFile, password= NULL, sheet=1, startCol=0, endCol=0, startRow=0, endRow=0, region=NULL, resultCols, openWhenDone = FALSE){
  #Validate Arguments #ToDo
  return(list('file'=excelFile,'password'=password,'sheet'=sheet, 'startCol'=startCol, 'endCol'=endCol, 'startRow' = startRow, 'endRow'=endRow, 'region'=region, 'resultCols'= resultCols, 'openWhenDone'=openWhenDone))
}



# try =function(s1,s2){
#   return(s1+s2)
# }


# io = RExcelerateIO('/home/mustapha/Desktop/try.xlsx',resultCols = 'Sum', openWhenDone = TRUE)
# RExcelerate(io,try,s1='rexc:op1',s2='rexc:op2')
