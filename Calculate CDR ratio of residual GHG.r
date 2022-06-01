library(readxl)
library(ggplot2)
library(grid)
library(dplyr)
library(gridExtra)
library(openxlsx)
library(tidyr)
library(ggsci)
library(scales)
library(tcltk)
library(purrr)
library(writexl)
setwd("D:/OneDrive - HV/1-Equitable CDR pathway/3-Calculation/1-Allocation170/results")


path<-c("Allocation_CAP.xlsx","Allocation_CAPcap.xlsx","Allocation_EPC.xlsx",
        "Allocation_EPCcap.xlsx","Allocation_RES.xlsx","Allocation_REScap.xlsx")

sheet_input <-excel_sheets(path[1]) #get the sheet name (IAM scenario name)
sheet_input <-sheet_input[-c(1,2)]  #delete unrelated summary
country     <-pull(read_excel(path[1],sheet =sheet_input[1]),country) #get country name list
country     <-country[-171]
data<-data.frame()
for (i in 1:6){
    for (j in 1:13){
    sheet_data     <- read_excel(path[i],sheet =sheet_input[j])
    sheet_data     <- sheet_data[-171,]
    scenario       <- c(rep(sheet_input[j],170))
    principle      <- c(rep(path[i],170))
    sheet_data     <- cbind(scenario,principle,sheet_data)
    data           <- rbind(sheet_data,data)
  }
}
data<-data[-c(3,86)]  #delete 2020-2050 cumulative

setwd("D:/OneDrive - HV/1-Equitable CDR pathway/3-Calculation/1-Allocation170")
GHG_BL_input   <- read_excel("IE-summary-170.xlsx",sheet = "net GHG")
net_GHG<-pivot_longer(GHG_BL_input,col=4:84,names_to = "year",values_to = "net GHG")
eql_CDR<-pivot_longer(data,col=4:84,names_to = "year",values_to = "eql_CDR")
CDR_share_output<-data.frame()

#############################    calculation results for 170 countries     #################################################
## show the progress
pb<-tkProgressBar("Progress","Done %",0,100)
star_time <- Sys.time() 
for (i in 1:170){  ##select country
  c_CDR <-eql_CDR[eql_CDR$country == country[i],]
  c_GHG <-net_GHG[net_GHG$country == country[i],]
  cspt_share_t<-data.frame()
  for (s in 1:13){ ##select scenario
    cs_CDR <-c_CDR[c_CDR$scenario == sheet_input[s],]
    cs_GHG <-c_GHG[c_GHG$scenario == sheet_input[s],]
    for (p in 1:6){  ##select principle
      csp_CDR <- cs_CDR[cs_CDR$principle == path[p],]
      for (y in 2020:2100){
        cspt_CDR<-csp_CDR[csp_CDR$year==y,]
        cspt_GHG<-cs_GHG[cs_GHG$year==y,]
        cspt_share<-cspt_CDR$eql_CDR/(cspt_CDR$eql_CDR+cspt_GHG$`net GHG`)
        cspt_share<-data.frame(country[i],sheet_input[s],path[p],y,cspt_share)
        cspt_share_t<-rbind(cspt_share_t,cspt_share)
      }
    }
    info <- sprintf("Country calculated %d%%", round(i*100/170)) 
    setTkProgressBar(pb, i*100/170, sprintf("进度 (%s)", info),info) 
  }
  cspc_share<-pivot_wider(cspt_share_t,names_from = y,values_from = cspt_share)
  CDR_share_output<-rbind(CDR_share_output,cspc_share)
}
end_time <- Sys.time()  
close(pb)               
run_time <- end_time - star_time 

results <- createWorkbook()
addWorksheet(results, "CDR_share")
addWorksheet(results, "net_GHG")
addWorksheet(results, "eql_CDR")
writeData(results, sheet = "CDR_share", x = CDR_share_output)
writeData(results, sheet = "net_GHG", x = net_GHG)
writeData(results, sheet = "eql_CDR", x = eql_CDR)
saveWorkbook(results, "#CDR_share_country_1-170.xlsx")