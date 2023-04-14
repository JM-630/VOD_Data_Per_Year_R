#turn vod raster data to spatial points DF

library(readxl)
library(writexl)
library(tidyverse)
library(rio)
library(openxlsx)

library(dplyr)
library(magrittr)
library(raster)
library(rgdal)
library(tiff)
library(WriteXLS)
library("xlsx")

wb = createWorkbook()

sheet = createSheet(wb, "2019_VOD_swiss_sites")

# change depending on the number of days in chosen month
Jan <- 31
Feb <- 28
Mar <- 31
Apr <- 30
May <- 31
June <- 30
July <- 31
Aug <- 31
Sep <- 30
Oct <- 31
Nov <- 30
Dec <- 31

# vectors containg that will contain the averages for each month
JanV <- c()
FebV <- c()
MarV <- c()
AprV <- c()
MayV <- c()
JuneV <- c()
JulyV <- c()
AugV <- c()
SepV <- c()
OctV <- c()
NovV <- c()
DecV <- c()

for (x in 1:Jan) {
  
  # change the path to determine which geotiff is being converted
  str_name <-'/Users/student/Documents/VOD_Project/vod_summer_geotiff/vod_january_2019'
  str_name <- paste(str_name, x, sep = "_")
  str_name <- paste(str_name, "tiff", sep = ".")
  imported_raster = raster(str_name)
  
  # change the csv file here to change where the VOD data is being taken
  site <- read.csv("/Users/student/Documents/VOD_Project/Swiss_Site2.csv", header=TRUE)
  
  coords = data.frame(lat=site[,2],long=site[,3])
  coordinates(coords)=c("long","lat")
  
  VOD = extract(x=imported_raster,y=coords)
  
  # adds value to the vector if it is not NA
  if (!is.na(VOD))
    JanV <- c(JanV, VOD)
  
}

# averages the data from the month
mean <- mean(JanV)
avg <- as.numeric(mean)

# adds the data to the excel sheet
addDataFrame(data.frame(Month = "January", Average = avg), sheet = sheet, startRow = 1, startColumn = 1, row.names = FALSE, col.names = FALSE)



for (x in 1:Feb) {
  
  # change the path to determine which geotiff is being converted
  str_name <-'/Users/student/Documents/VOD_Project/vod_summer_geotiff/vod_february_2019'
  str_name <- paste(str_name, x, sep = "_")
  str_name <- paste(str_name, "tiff", sep = ".")
  imported_raster = raster(str_name)
  
  # change the csv file here to change where the VOD data is being taken
  site <- read.csv("/Users/student/Documents/VOD_Project/Swiss_Site2.csv", header=TRUE)
  
  coords = data.frame(lat=site[,2],long=site[,3])
  coordinates(coords)=c("long","lat")
  
  VOD = extract(x=imported_raster,y=coords)
  
  # adds value to the vector if it is not NA
  if (!is.na(VOD))
    FebV <- c(FebV, VOD)
  
}

# averages the data from the month
mean <- mean(FebV)
avg <- as.numeric(mean)

# adds the data to the excel sheet
addDataFrame(data.frame(Month = "February", Average = avg), sheet = sheet, startRow = 2, startColumn = 1, row.names = FALSE, col.names = FALSE)



for (x in 1:Mar) {
  
  # change the path to determine which geotiff is being converted
  str_name <-'/Users/student/Documents/VOD_Project/vod_summer_geotiff/vod_march_2019'
  str_name <- paste(str_name, x, sep = "_")
  str_name <- paste(str_name, "tiff", sep = ".")
  imported_raster = raster(str_name)
  
  # change the csv file here to change where the VOD data is being taken
  site <- read.csv("/Users/student/Documents/VOD_Project/Swiss_Site2.csv", header=TRUE)
  
  coords = data.frame(lat=site[,2],long=site[,3])
  coordinates(coords)=c("long","lat")
  
  VOD = extract(x=imported_raster,y=coords)
  
  # adds value to the vector if it is not NA
  if (!is.na(VOD))
    MarV <- c(MarV, VOD)
  
}

# averages the data from the month
mean <- mean(MarV)
avg <- as.numeric(mean)

# adds the data to the excel sheet
addDataFrame(data.frame(Month = "March", Average = avg), sheet = sheet, startRow = 3, startColumn = 1, row.names = FALSE, col.names = FALSE)



for (x in 1:Apr) {
  
  # change the path to determine which geotiff is being converted
  str_name <-'/Users/student/Documents/VOD_Project/vod_summer_geotiff/vod_april_2019'
  str_name <- paste(str_name, x, sep = "_")
  str_name <- paste(str_name, "tiff", sep = ".")
  imported_raster = raster(str_name)
  
  # change the csv file here to change where the VOD data is being taken
  site <- read.csv("/Users/student/Documents/VOD_Project/Swiss_Site2.csv", header=TRUE)
  
  coords = data.frame(lat=site[,2],long=site[,3])
  coordinates(coords)=c("long","lat")
  
  VOD = extract(x=imported_raster,y=coords)
  
  # adds value to the vector if it is not NA
  if (!is.na(VOD))
    AprV <- c(AprV, VOD)
  
}

# averages the data from the month
mean <- mean(AprV)
avg <- as.numeric(mean)

# adds the data to the excel sheet
addDataFrame(data.frame(Month = "April", Average = avg), sheet = sheet, startRow = 4, startColumn = 1, row.names = FALSE, col.names = FALSE)


for (x in 1:May) {
  
  # change the path to determine which geotiff is being converted
  str_name <-'/Users/student/Documents/VOD_Project/vod_summer_geotiff/vod_may_2019'
  str_name <- paste(str_name, x, sep = "_")
  str_name <- paste(str_name, "tiff", sep = ".")
  imported_raster = raster(str_name)
  
  # change the csv file here to change where the VOD data is being taken
  site <- read.csv("/Users/student/Documents/VOD_Project/Swiss_Site2.csv", header=TRUE)
  
  coords = data.frame(lat=site[,2],long=site[,3])
  coordinates(coords)=c("long","lat")
  
  VOD = extract(x=imported_raster,y=coords)
  
  # adds value to the vector if it is not NA
  if (!is.na(VOD))
    MayV <- c(MayV, VOD)
  
}

# averages the data from the month
mean <- mean(MayV)
avg <- as.numeric(mean)

# adds the data to the excel sheet
addDataFrame(data.frame(Month = "May", Average = avg), sheet = sheet, startRow = 5, startColumn = 1, row.names = FALSE, col.names = FALSE)



for (x in 1:June) {
  
  # change the path to determine which geotiff is being converted
  str_name <-'/Users/student/Documents/VOD_Project/vod_summer_geotiff/vod_june_2019'
  str_name <- paste(str_name, x, sep = "_")
  str_name <- paste(str_name, "tiff", sep = ".")
  imported_raster = raster(str_name)
  
  # change the csv file here to change where the VOD data is being taken
  site <- read.csv("/Users/student/Documents/VOD_Project/Swiss_Site2.csv", header=TRUE)
  
  coords = data.frame(lat=site[,2],long=site[,3])
  coordinates(coords)=c("long","lat")
  
  VOD = extract(x=imported_raster,y=coords)
  
  # adds value to the vector if it is not NA
  if (!is.na(VOD))
    JuneV <- c(JuneV, VOD)
  
}

# averages the data from the month
mean <- mean(JuneV)
avg <- as.numeric(mean)

# adds the data to the excel sheet
addDataFrame(data.frame(Month = "June", Average = avg), sheet = sheet, startRow = 6, startColumn = 1, row.names = FALSE, col.names = FALSE)


for (x in 1:July) {
  
  # change the path to determine which geotiff is being converted
  str_name <-'/Users/student/Documents/VOD_Project/vod_summer_geotiff/vod_july_2019'
  str_name <- paste(str_name, x, sep = "_")
  str_name <- paste(str_name, "tiff", sep = ".")
  imported_raster = raster(str_name)
  
  # change the csv file here to change where the VOD data is being taken
  site <- read.csv("/Users/student/Documents/VOD_Project/Swiss_Site2.csv", header=TRUE)
  
  coords = data.frame(lat=site[,2],long=site[,3])
  coordinates(coords)=c("long","lat")
  
  VOD = extract(x=imported_raster,y=coords)
  
  # adds value to the vector if it is not NA
  if (!is.na(VOD))
    JulyV <- c(JulyV, VOD)
  
}

# averages the data from the month
mean <- mean(JulyV)
avg <- as.numeric(mean)

# adds the data to the excel sheet
addDataFrame(data.frame(Month = "July", Average = avg), sheet = sheet, startRow = 7, startColumn = 1, row.names = FALSE, col.names = FALSE)



for (x in 1:Aug) {
  
  # change the path to determine which geotiff is being converted
  str_name <-'/Users/student/Documents/VOD_Project/vod_summer_geotiff/vod_august_2019'
  str_name <- paste(str_name, x, sep = "_")
  str_name <- paste(str_name, "tiff", sep = ".")
  imported_raster = raster(str_name)
  
  # change the csv file here to change where the VOD data is being taken
  site <- read.csv("/Users/student/Documents/VOD_Project/Swiss_Site2.csv", header=TRUE)
  
  coords = data.frame(lat=site[,2],long=site[,3])
  coordinates(coords)=c("long","lat")
  
  VOD = extract(x=imported_raster,y=coords)
  
  # adds value to the vector if it is not NA
  if (!is.na(VOD))
    AugV <- c(AugV, VOD)
  
}

# averages the data from the month
mean <- mean(AugV)
avg <- as.numeric(mean)

# adds the data to the excel sheet
addDataFrame(data.frame(Month = "August", Average = avg), sheet = sheet, startRow = 8, startColumn = 1, row.names = FALSE, col.names = FALSE)



for (x in 1:Sep) {
  
  # change the path to determine which geotiff is being converted
  str_name <-'/Users/student/Documents/VOD_Project/vod_summer_geotiff/vod_september_2019'
  str_name <- paste(str_name, x, sep = "_")
  str_name <- paste(str_name, "tiff", sep = ".")
  imported_raster = raster(str_name)
  
  # change the csv file here to change where the VOD data is being taken
  site <- read.csv("/Users/student/Documents/VOD_Project/Swiss_Site2.csv", header=TRUE)
  
  coords = data.frame(lat=site[,2],long=site[,3])
  coordinates(coords)=c("long","lat")
  
  VOD = extract(x=imported_raster,y=coords)
  
  # adds value to the vector if it is not NA
  if (!is.na(VOD))
    SepV <- c(SepV, VOD)
  
}

# averages the data from the month
mean <- mean(SepV)
avg <- as.numeric(mean)

# adds the data to the excel sheet
addDataFrame(data.frame(Month = "September", Average = avg), sheet = sheet, startRow = 9, startColumn = 1, row.names = FALSE, col.names = FALSE)



for (x in 1:Oct) {
  
  # change the path to determine which geotiff is being converted
  str_name <-'/Users/student/Documents/VOD_Project/vod_summer_geotiff/vod_october_2019'
  str_name <- paste(str_name, x, sep = "_")
  str_name <- paste(str_name, "tiff", sep = ".")
  imported_raster = raster(str_name)
  
  # change the csv file here to change where the VOD data is being taken
  site <- read.csv("/Users/student/Documents/VOD_Project/Swiss_Site2.csv", header=TRUE)
  
  coords = data.frame(lat=site[,2],long=site[,3])
  coordinates(coords)=c("long","lat")
  
  VOD = extract(x=imported_raster,y=coords)
  
  # adds value to the vector if it is not NA
  if (!is.na(VOD))
    OctV <- c(OctV, VOD)
  
}

# averages the data from the month
mean <- mean(OctV)
avg <- as.numeric(mean)

# adds the data to the excel sheet
addDataFrame(data.frame(Month = "October", Average = avg), sheet = sheet, startRow = 10, startColumn = 1, row.names = FALSE, col.names = FALSE)



for (x in 1:Nov) {
  
  # change the path to determine which geotiff is being converted
  str_name <-'/Users/student/Documents/VOD_Project/vod_summer_geotiff/vod_november_2019'
  str_name <- paste(str_name, x, sep = "_")
  str_name <- paste(str_name, "tiff", sep = ".")
  imported_raster = raster(str_name)
  
  # change the csv file here to change where the VOD data is being taken
  site <- read.csv("/Users/student/Documents/VOD_Project/Swiss_Site2.csv", header=TRUE)
  
  coords = data.frame(lat=site[,2],long=site[,3])
  coordinates(coords)=c("long","lat")
  
  VOD = extract(x=imported_raster,y=coords)
  
  # adds value to the vector if it is not NA
  if (!is.na(VOD))
    NovV <- c(NovV, VOD)
  
}

# averages the data from the month
mean <- mean(NovV)
avg <- as.numeric(mean)

# adds the data to the excel sheet
addDataFrame(data.frame(Month = "November", Average = avg), sheet = sheet, startRow = 11, startColumn = 1, row.names = FALSE,col.names = FALSE)


for (x in 1:Dec) {
  
  # change the path to determine which geotiff is being converted
  str_name <-'/Users/student/Documents/VOD_Project/vod_summer_geotiff/vod_december_2019'
  str_name <- paste(str_name, x, sep = "_")
  str_name <- paste(str_name, "tiff", sep = ".")
  imported_raster = raster(str_name)
  
  # change the csv file here to change where the VOD data is being taken
  site <- read.csv("/Users/student/Documents/VOD_Project/Swiss_Site2.csv", header=TRUE)
  
  coords = data.frame(lat=site[,2],long=site[,3])
  coordinates(coords)=c("long","lat")
  
  VOD = extract(x=imported_raster,y=coords)
  
  # adds value to the vector if it is not NA
  if (!is.na(VOD))
    DecV <- c(DecV, VOD)
  
}

# averages the data from the month
mean <- mean(DecV)
avg <- as.numeric(mean)

# adds the data to the excel sheet
addDataFrame(data.frame(Month = "December", Average = avg), sheet = sheet, startRow = 12, startColumn = 1, row.names = FALSE, col.names = FALSE)


saveWorkbook(wb, "VOD_2019_Switzerland_Site_1452.xlsx")



