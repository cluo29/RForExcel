#install.packages("XLConnect")

#install.packages("WriteXLS")
setwd("~/Desktop")

library(WriteXLS) #load the package
require("XLConnect")

MaxPerson=100
Month=2

p <- function(..., sep='') {
  paste(..., sep=sep, collapse=sep)
}

countSum<-function(MonthNumber,Result){
  #Jan is 1
  FileName = p(MonthNumber,".xls")
  data1 <- loadWorkbook(FileName, create = FALSE)
  dateFrame1 <- readWorksheet(data1, sheet = "Sheet1")
  
  for(i in 2:MaxPerson)
  {
    if(is.na(dateFrame1[i,2])==FALSE)
    {
      Result[i-1,1] = dateFrame1[i,2]
    }
    if(is.na(dateFrame1[i,3])==FALSE)
    {
      Result[i-1,2] = dateFrame1[i,3]
    }
    if(is.na(dateFrame1[i,4])==FALSE)
    {
      Result[i-1,3] = Result[i-1,3] + as.numeric(dateFrame1[i,4])
    }
    if(is.na(dateFrame1[i,5])==FALSE)
    {
      Result[i-1,4] = Result[i-1,4] + as.numeric(dateFrame1[i,5])
    }
  }
  
  return(Result)
}

personID=1:MaxPerson
Name=1:MaxPerson
GongZuoLiang=1:MaxPerson
GongXingKao=1:MaxPerson
Result=data.frame(personID,Name,GongZuoLiang,GongXingKao)

Result[,3]=0
Result[,4]=0

for(i in 1: Month)
{
  Result = countSum(i,Result)
}


WriteXLS(Result, ExcelFileName = "Result.xls", Encoding = "UTF-8"
         )

