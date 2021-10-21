library(tidyverse)
library(lubridate)

# Only change the company, if needed, do not alter anything else.

#################################################################

company <- rstudioapi::showPrompt("Input Company", "Input Company Code")

#################################################################

# Report must be started when in the next month or the below date lines will not be correct

mtd_start_POSIXct <- floor_date(now(), "month") - months(1)
mtd_end_POSIXct <- ceiling_date(now(), "month") - months(1) - days(1)
ytd_start_POSIXct <- floor_date(mtd_end_POSIXct, "year")

mtd_start <- format(mtd_start_POSIXct, "%Y/%m/%d")
mtd_end <- format(mtd_end_POSIXct, "%Y/%m/%d")
ytd_start <- format(ytd_start_POSIXct, "%Y/%m/%d")


pathInput <- rstudioapi::selectDirectory(caption = "Select Input Directory", 
                                    label = "Select",
                                    path = getwd())

# Invoice and credits line items manipulation

ic <- read_csv(rstudioapi::selectFile(caption = "Select INLineItem Report File",
                                      path = pathInput),
               col_types = cols(CustomerID = col_character(), 
                                SalesDate = col_date(format = "%d/%m/%Y"), 
                                VendorID = col_character()))

ic <- separate(ic,ProductLineDesc, into=c("ProductLine", "ProductLineDescr"), sep = " ", extra="merge")

ic$ProductLineDescr <- ifelse(is.na(ic$ProductLineDescr), ic$ProductLine, ic$ProductLineDescr)
  
ic$Quantity <- ic$Quantity1 * ic$ParameterOfEffect * -1

ic <- ic %>% 
  filter(!grepl("MEGA", CustomerID))

ic$LaborCost <- 0
ic$ExternalCost <- 0
ic$OtherCost <- 0
ic$ReconciliationCost <- 0
ic$AssemblyMinutes <- 0
ic$ProcessingMinutes <- 0
ic$BrandID <- ""
ic$ProductFamily2 <- ""
ic$ProductFamilyDescr2 <- ""
ic$ApplicationSector <- ""
ic$PriceUnit <- 1
ic$CompanyID <- ""

# State needs to be the full State Name

ic$State <- ifelse(ic$Country != "AUSTRALIA", ic$State <- ic$Country, ic$State <- ic$StateAustralia)

StateAbb <- c("NSW", "VIC", "QLD", "SA", "WA", "NT", "ACT", "TAS")
StateFull <- c("New South Wales", "Victoria", "Queensland", "South Australia", "Western Australia", "Northern Territory", "Australian Capital Territory", "Tasmanina")

for (i in 1:8){
  ic$State <- str_replace(ic$State, StateAbb[i], StateFull[i])
}
  
ic <- unite(ic, "SalesRef", c("TranType", "Reference"), sep = " ", remove = TRUE)

ic <- select(ic, c(SalesDate, ProductLine, ProductLineDescr, ProductFamily, ProductFamilyDesc, PartNumber, PartDescr, TotalAmount, Quantity, PurchaseCost, LaborCost, ExternalCost, OtherCost, AssemblyMinutes, ProcessingMinutes, CustomerID, CustomerName, State, DistributionChannel, BranchID, SalesmanID, BrandID, SalesRef, Currency, InvoiceCurrency, InvoiceAmount, ProductFamily2, ProductFamilyDescr2, VendorID, VendorName, ApplicationSector, ReconciliationCost, SalesPrice, PriceUnit, CompanyID
))

### JC Stock Code Manipulation

jca <- read_csv(rstudioapi::selectFile(caption = "Select JCStockCode Report File",
                                       path = pathInput), 
                col_types = cols(CustomerID = col_character(), 
                                 SalesDate = col_date(format = "%d/%m/%Y"),
                                 SalesDate_FirstInv = col_date(format = "%d/%m/%Y"),
                                 VendorID = col_character()))

jca <- separate(jca,`ProductLine/Desc`, into=c("ProductLine", "ProductLineDescr"), sep = " ", extra="merge")

RRPList <- read_csv(rstudioapi::selectFile(caption = "Select RRPList Report File",
                                           path = pathInput))%>% 
  drop_na()

jca <- left_join(jca, RRPList, by="PartNumber")

# JC Activity Code Manipulation

jcb <- read_csv(rstudioapi::selectFile(caption = "Select JCActivityCostLine Report File",
                                       path = pathInput), 
                col_types = cols(CustomerID = col_character(), 
                                 SalesDate = col_date(format = "%d/%m/%Y"),
                                 SalesDate_FirstInv = col_date(format = "%d/%m/%Y"),
                                 VendorID = col_character()))

jcb$ProductLine <- "P10"
jcb$ProductLineDescr <- "FleetFit"
jcb$ProductFamily <- "FleetFit"
jcb$ProductFamilyDesc <- "FleetFit"

# JC PO Variance Manipulation

jcc <- read_csv(rstudioapi::selectFile(caption = "Select JCPOVar Report File",
                                       path = pathInput), 
                col_types = cols(CustomerID = col_character(), 
                                 SalesDate = col_date(format = "%d/%m/%Y"),
                                 SalesDate_FirstInv = col_date(format = "%d/%m/%Y"),
                                 VendorID = col_character()))

jcc$ProductLine <- "P10"
jcc$ProductLineDescr <- "FleetFit"
jcc$ProductFamily <- "FleetFit"
jcc$ProductFamilyDesc <- "FleetFit"

jc <- bind_rows(jca, jcb, jcc) %>% 
   arrange(JobNumber)

CreditJobNumberData <- read_csv(rstudioapi::selectFile(caption = "Select JC Credit Note Report File",
                                                       path = pathInput), 
                                col_types = cols(SalesDate = col_date(format = "%d/%m/%Y")))

CreditJobNumberData$ValueTotal <- CreditJobNumberData$ValueTotal * -1

InvoiceJobNumberData <- read_csv(rstudioapi::selectFile(caption = "Select JCARInvoice Report File",
                                                        path = pathInput), 
                                 col_types = cols(SalesDate = col_date(format = "%d/%m/%Y")))

InvCred <- bind_rows(CreditJobNumberData, InvoiceJobNumberData)

JobNumberData1 <- InvCred %>% 
  group_by(JobNumber) %>% 
  summarise(SalesDate = min(SalesDate),
            ValueTotal1 = sum(ValueTotal), 
            SalesTotal = sum(SalesTotal),
            SalesRef = paste0("JC Issue"," ", max(Reference), " (", max(JobNumber), ")"),
            SalesmanID = max(SalesmanID),
            Currency = "AUD",
            InvoiceCurrency = "AUD")

# This sums the combined JC data
  
JobNumberData2 <- jc %>% 
  group_by(JobNumber) %>% 
  summarise(ValueTotal2 = sum(PurchaseCost))

JobNumberData <- inner_join(JobNumberData1, JobNumberData2, by="JobNumber") %>% 
  mutate(PercentCost = ifelse(ValueTotal1!=0, ValueTotal2/ValueTotal1, 0)) %>% 
  rename("ValueTotal" = "ValueTotal2") %>% 
  select(-ValueTotal1)

jc <- jc %>% 
  inner_join(JobNumberData, by="JobNumber") %>% 
  rename("SalesDate" = "SalesDate.y", "SalesmanID" = "SalesmanID.y", "Currency" = "Currency.y", "InvoiceCurrency" = "InvoiceCurrency.y", "ValueTotal" = "ValueTotal.y", "SalesTotal" = "SalesTotal.y") %>% 
  select(-ends_with(".x")) 

jc$LaborCost <- 0
jc$ExternalCost <- 0
jc$OtherCost <- 0
jc$ReconciliationCost <- 0
jc$AssemblyMinutes <- 0
jc$ProcessingMinutes <- 0
jc$BrandID <- ""
jc$ProductFamily2 <- ""
jc$ProductFamilyDescr2 <- ""
jc$ApplicationSector <- ""
jc$PriceUnit <- 1
jc$CompanyID <- ""

# State needs to be the full State Name

for (i in 1:8){
  jc$State <- str_replace(jc$State, StateAbb[i], StateFull[i])
}

jc$TotalAmount <- ifelse(jc$ValueTotal !=0, (round(jc$PurchaseCost/jc$ValueTotal*jc$SalesTotal, 2)), 0)

jc$InvoiceAmount <- jc$TotalAmount

jc <- select(jc, c(SalesDate, ProductLine, ProductLineDescr, ProductFamily, ProductFamilyDesc, PartNumber, PartDescr, 
                   TotalAmount, Quantity, PurchaseCost, LaborCost, ExternalCost, OtherCost, AssemblyMinutes, 
                   ProcessingMinutes, CustomerID, CustomerName, State, DistributionChannel, BranchID, SalesmanID, 
                   BrandID, SalesRef, Currency, InvoiceCurrency, InvoiceAmount, ProductFamily2, ProductFamilyDescr2, 
                   VendorID, VendorName, ApplicationSector, ReconciliationCost, SalesPrice, PriceUnit, CompanyID))

reportColumns <- colnames(ic)

fullSalesReport <- bind_rows(ic, jc)

fullSalesReport$SalesDate <- format(fullSalesReport$SalesDate,"%Y/%m/%d")
fullSalesReport$SalesPrice <- replace_na(fullSalesReport$SalesPrice, 0)
fullSalesReport$VendorID <- replace_na(fullSalesReport$VendorID, "")
fullSalesReport$VendorName <- replace_na(fullSalesReport$VendorName, "")
fullSalesReport$ProductLineDescr <- ifelse(is.na(fullSalesReport$ProductLineDescr), fullSalesReport$ProductLine, 
                                           fullSalesReport$ProductLineDescr)
fullSalesReport$ProductFamily <- ifelse(is.na(fullSalesReport$ProductFamily), fullSalesReport$ProductLineDescr, 
                                           fullSalesReport$ProductFamily)
fullSalesReport$ProductFamilyDesc <- ifelse(is.na(fullSalesReport$ProductFamilyDesc), fullSalesReport$ProductLineDescr, 
                                        fullSalesReport$ProductFamilyDesc)

hasNA <- fullSalesReport[rowSums(is.na(fullSalesReport))!=0, ]

summary(fullSalesReport)

pathOutput <- rstudioapi::selectDirectory(caption = "Select Output Directory", 
                                    label = "Select",
                                    path = pathInput)

mtd <- filter(fullSalesReport, SalesDate <= mtd_end & SalesDate >= mtd_start)

ytd <- filter(fullSalesReport, SalesDate <= mtd_end & SalesDate >= ytd_start)

all_to_mth_end <- filter(fullSalesReport, SalesDate <= mtd_end)

write_delim(mtd, paste0(pathOutput, "//", company, "_", gsub("\\/", "", mtd_start), "_", gsub("\\/", "", mtd_end), ".CSV"), delim=";", na = "", append = FALSE,
            col_names = TRUE, escape = "double")

write_delim(ytd, paste0(pathOutput, "//", company, "_", gsub("\\/", "", ytd_start), "_", gsub("\\/", "", mtd_end), ".CSV"), delim=";", na = "", append = FALSE,
            col_names = TRUE, escape = "double")

write_delim(all_to_mth_end, paste0(pathOutput, "//", company, "_20170201_", gsub("\\/", "", mtd_end), ".CSV"), delim=";", na = "", append = FALSE,
            col_names = TRUE, escape = "double")