setwd("~/Documents/Code/2018/Rxlsm")

#install.packages("xlsx")
library(xlsx)
data<-read.xlsx2(file="a.xls",header=T,sheetIndex=1) 

dataColNum = ncol(data)

dataRowNum = nrow(data)
 
dataDF <- data.frame(lapply(data, as.character), stringsAsFactors=FALSE)




outputDF = data.frame()

for (i in 1:dataRowNum){
#for (i in 1:20){
  # fetch one row from raw data
  dataThisRow = dataDF[i,]
  
  outputRow = data.frame()
  
  outputRow[1,1] = dataThisRow$sec_cik
  outputRow[1,2] = dataThisRow$co_conm
  outputRow[1,3] = dataThisRow$fiscal_yr
  outputRow[1,4] = dataThisRow$sec_proxy_link
  outputRow[1,5] = dataThisRow$sec_8k_link
  outputRow[1,6] = dataThisRow$sec_10k_link
  outputRow[1,7] = dataThisRow$exec_name
  
  outputRow[2,1] = NA
  outputRow[2,2] = NA
  outputRow[2,3] = NA
  outputRow[2,4] = "metrics"
  outputRow[2,5] = "metric_description"
  outputRow[2,6] = "metric_classify"
  outputRow[2,7] = "metric_recon_prxy"
  outputRow[2,8] = "metric_recon_810K"
  
  outputRow[3,1] = NA
  outputRow[4,1] = NA
  outputRow[5,1] = NA
  outputRow[6,1] = NA
  
  
  currentCell = dataThisRow$metric_1
  if(currentCell =="" | is.na(currentCell) | is.null(currentCell))
  {
    # if empty
  } else
  {
    outputRow[3,4] = currentCell
  }
  
  currentCell = dataThisRow$metric_2
  if(currentCell =="" | is.na(currentCell) | is.null(currentCell))
  {
    # if empty
  } else
  {
    outputRow[4,4] = currentCell
  }
  
  currentCell = dataThisRow$metric_3
  if(currentCell =="" | is.na(currentCell) | is.null(currentCell))
  {
    # if empty
  } else
  {
    outputRow[5,4] = currentCell
  }
  
  
  currentCell = dataThisRow$metric_4
  if(currentCell =="" | is.na(currentCell) | is.null(currentCell))
  {
    # if empty
  } else
  {
    outputRow[6,4] = currentCell
  }
  
  
  
  currentCell = dataThisRow$metric_5
  if(currentCell =="" | is.na(currentCell) | is.null(currentCell))
  {
    # if empty
  } else
  {
    outputRow[7,4] = currentCell
  }
  
  
  currentCell = dataThisRow$metric_6
  if(currentCell =="" | is.na(currentCell) | is.null(currentCell))
  {
    # if empty
  } else
  {
    outputRow[8,4] = currentCell
  }
  
  
  currentCell = dataThisRow$metric_7
  if(currentCell =="" | is.na(currentCell) | is.null(currentCell))
  {
    # if empty
  } else
  {
    outputRow[9,4] = currentCell
  }
  
  
  currentCell = dataThisRow$metric_8
  if(currentCell =="" | is.na(currentCell) | is.null(currentCell))
  {
    # if empty
  } else
  {
    outputRow[10,4] = currentCell
  }
  
  
  currentCell = dataThisRow$metric_9
  if(currentCell =="" | is.na(currentCell) | is.null(currentCell))
  {
    # if empty
  } else
  {
    outputRow[11,4] = currentCell
  }
  
  
  
  currentCell = dataThisRow$metric_10
  if(currentCell =="" | is.na(currentCell) | is.null(currentCell))
  {
    # if empty
  } else
  {
    outputRow[12,4] = currentCell
  }
  
  outputDF = rbind(outputDF,outputRow)
  
}


outputRowHead = data.frame()
outputRowHead[1,1] = "sec_cik"
outputRowHead[1,2] = "co_conm"
outputRowHead[1,3] = "fiscal_yr"
outputRowHead[1,4] = "sec_proxy_link"
outputRowHead[1,5] = "sec_8k_link"
outputRowHead[1,6] = "sec_10k_link"
outputRowHead[1,7] = "exec_name"
outputRowHead[1,8] = ""

outputRowHead = rbind(outputRowHead,outputDF)

write.xlsx2(outputRowHead, "output.xls", sheetName="Sheet1", col.names=FALSE, row.names=FALSE, append=FALSE)
