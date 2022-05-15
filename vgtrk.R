
library(utf8)
Sys.setlocale("LC_ALL", "Russian")
library(tidyverse)
library(readr)
library(dplyr)
library(stringr)
library(readr)
library(tidytext)
library(openxlsx)
library(fuzzyjoin)
library(topicmodels)
library(ggplot2)
library(tidyr)

library(flextable)
library(GGally)
library(ggraph)
library(gutenbergr)
library(igraph)
library(Matrix)
library(network)
library(quanteda)
library(sna)
library(tidygraph)
library(tidyverse)
library(tm)
library(tibble)

### Importing pre-processed emails from csv ###

vgtrk <- read_delim("C:/Users/Savanna/Desktop/Code_projects/R_projects/Text_mining_with_R/all_messages.csv", col_names = FALSE)

## Use smaller chunks of emails and use rm() & gc() to free up RAM

vgtrk1a = vgtrk[1:10000,]
vgtrk1b = vgtrk[10001:20000,]
vgtrk1c = vgtrk[20001:30000,]
vgtrk1d = vgtrk[30001:40000,]
vgtrk1e = vgtrk[40001:50000,]
vgtrk2a = vgtrk[50001:60000,]
vgtrk2b = vgtrk[60001:70000,]
vgtrk2c = vgtrk[70001:80000,]
vgtrk2d = vgtrk[80001:90000,]
vgtrk2e = vgtrk[90001:100000,]

vgtrk3a = vgtrk[100001:110000,]
vgtrk3b = vgtrk[110001:120000,]
vgtrk3c = vgtrk[120001:130000,]
vgtrk3d = vgtrk[130001:140000,]
vgtrk3e = vgtrk[140001:150000,]
vgtrk4a = vgtrk[150001:160000,]
vgtrk4b = vgtrk[160001:170000,]
vgtrk4c = vgtrk[170001:180000,]
vgtrk4d = vgtrk[180001:190000,]
vgtrk4e = vgtrk[190001:200000,]

vgtrk5a = vgtrk[200001:210000,]
vgtrk5b = vgtrk[210001:220000,]
vgtrk5c = vgtrk[220001:230000,]
vgtrk5d = vgtrk[230001:240000,]
vgtrk5e = vgtrk[240001:250000,]
vgtrk6a = vgtrk[250001:260000,]
vgtrk6b = vgtrk[260001:270000,]
vgtrk6c = vgtrk[270001:280000,]
vgtrk6d = vgtrk[280001:290000,]
vgtrk6e = vgtrk[290001:300000,]

vgtrk7a = vgtrk[300001:310000,]
vgtrk7b = vgtrk[310001:320000,]
vgtrk7c = vgtrk[320001:330000,]
vgtrk7d = vgtrk[330001:340000,]
vgtrk7e = vgtrk[340001:350000,]
vgtrk8a = vgtrk[350001:360000,]
vgtrk8b = vgtrk[360001:370000,]
vgtrk8c = vgtrk[370001:380000,]
vgtrk8d = vgtrk[380001:390000,]
vgtrk8e = vgtrk[390001:400000,]

vgtrk9a = vgtrk[400001:410000,]
vgtrk9b = vgtrk[410001:420000,]
vgtrk9c = vgtrk[420001:430000,]
vgtrk9d = vgtrk[430001:440000,]
vgtrk9e = vgtrk[440001:450000,]
vgtrk10a = vgtrk[450001:460000,]
vgtrk10b = vgtrk[460001:470000,]
vgtrk10c = vgtrk[470001:480000,]
vgtrk10d = vgtrk[480001:490000,]
vgtrk10e = vgtrk[490001:500000,]

vgtrk11a = vgtrk[500001:510000,]
vgtrk11b = vgtrk[510001:520000,]
vgtrk11c = vgtrk[520001:530000,]
vgtrk11d = vgtrk[530001:540000,]
vgtrk11e = vgtrk[540001:550000,]
vgtrk12a = vgtrk[550001:560000,]
vgtrk12b = vgtrk[560001:570000,]
vgtrk12c = vgtrk[570001:580000,]
vgtrk12d = vgtrk[580001:590000,]
vgtrk12e = vgtrk[590001:600000,]

vgtrk13a = vgtrk[600001:610000,]
vgtrk13b = vgtrk[610001:620000,]
vgtrk13c = vgtrk[620001:630000,]
vgtrk13d = vgtrk[630001:640000,]
vgtrk13e = vgtrk[640001:650000,]
vgtrk14a = vgtrk[650001:660000,]
vgtrk14b = vgtrk[660001:670000,]
vgtrk14c = vgtrk[670001:680000,]
vgtrk14d = vgtrk[680001:690000,]
vgtrk14e = vgtrk[690001:700000,]

vgtrk15a = vgtrk[700001:710000,]
vgtrk15b = vgtrk[710001:720000,]
vgtrk15c = vgtrk[720001:730000,]
vgtrk15d = vgtrk[730001:740000,]
vgtrk15e = vgtrk[740001:750000,]
vgtrk16a = vgtrk[750001:760000,]
vgtrk16b = vgtrk[760001:770000,]
vgtrk16c = vgtrk[770001:780000,]
vgtrk16d = vgtrk[780001:790000,]
vgtrk16e = vgtrk[790001:800000,]

vgtrk17a = vgtrk[800001:810000,]
vgtrk17b = vgtrk[810001:820000,]
vgtrk17c = vgtrk[820001:830000,]
vgtrk17d = vgtrk[830001:840000,]
vgtrk17e = vgtrk[840001:850000,]
vgtrk18a = vgtrk[850001:860000,]
vgtrk18b = vgtrk[860001:870000,]
vgtrk18c = vgtrk[870001:880000,]
vgtrk18d = vgtrk[880001:890000,]
vgtrk18e = vgtrk[890001:900000,]

vgtrk19a = vgtrk[900001:910000,]
vgtrk19b = vgtrk[910001:920000,]
vgtrk19c = vgtrk[920001:930000,]
vgtrk19d = vgtrk[930001:940000,]
vgtrk19e = vgtrk[940001:950000,]
vgtrk20a = vgtrk[950001:960000,]
vgtrk20b = vgtrk[960001:970000,]
vgtrk20c = vgtrk[970001:980000,]
vgtrk20d = vgtrk[980001:990000,]
vgtrk20e = vgtrk[1000001:1010000,]

vgtrk21a = vgtrk[1010001:1020000,]
vgtrk21b = vgtrk[1020001:1030000,]
vgtrk21c = vgtrk[1030001:1040000,]
vgtrk21d = vgtrk[1040001:1050000,]
vgtrk21e = vgtrk[1050001:1060000,]
vgtrk22a = vgtrk[1060001:1070000,]
vgtrk22b = vgtrk[1070001:1080000,]
vgtrk22c = vgtrk[1080001:1090000,]
vgtrk22d = vgtrk[1090001:1100000,]
vgtrk22e = vgtrk[1100001:1110000,]

vgtrk23a = vgtrk[1110001:1120000,]
vgtrk23b = vgtrk[1120001:1130000,]
vgtrk23c = vgtrk[1130001:1140000,]
vgtrk23d = vgtrk[1140001:1150000,]
vgtrk23e = vgtrk[1150001:1160000,]
vgtrk24a = vgtrk[1160001:1170000,]
vgtrk24b = vgtrk[1170001:1180000,]
vgtrk24c = vgtrk[1180001:1190000,]
vgtrk24d = vgtrk[1190001:1200000,]
vgtrk24e = vgtrk[1200001:1210000,]

vgtrk25a = vgtrk[1210001:1220000,]
vgtrk25b = vgtrk[1220001:1230000,]
vgtrk25c = vgtrk[1230001:1240000,]
vgtrk25d = vgtrk[1240001:1250000,]
vgtrk25e = vgtrk[1250001:1260000,]
vgtrk26a = vgtrk[1260001:1270000,]
vgtrk26b = vgtrk[1270001:1280000,]
vgtrk26c = vgtrk[1280001:1290000,]
vgtrk26d = vgtrk[1290001:1300000,]
vgtrk26e = vgtrk[1300001:1310000,]

vgtrk27a = vgtrk[1310001:1320000,]
vgtrk27b = vgtrk[1320001:1330000,]
vgtrk27c = vgtrk[1330001:1340000,]
vgtrk27d = vgtrk[1340001:1350000,]
vgtrk27e = vgtrk[1350001:1360000,]
vgtrk28a = vgtrk[1360001:1370000,]
vgtrk28b = vgtrk[1370001:1380000,]
vgtrk28c = vgtrk[1380001:1390000,]
vgtrk28d = vgtrk[1390001:1400000,]
vgtrk28e = vgtrk[1400001:1410000,]

vgtrk29a = vgtrk[1410001:1420000,]
vgtrk29b = vgtrk[1420001:1430000,]
vgtrk29c = vgtrk[1430001:1440000,]
vgtrk29d = vgtrk[1440001:1450000,]
vgtrk29e = vgtrk[1450001:1460000,]
vgtrk30a = vgtrk[1460001:1470000,]
vgtrk30b = vgtrk[1470001:1480000,]
vgtrk30c = vgtrk[1480001:1490000,]
vgtrk30d = vgtrk[1490001:1500000,]
vgtrk30e = vgtrk[1500001:1510000,]

vgtrk31a = vgtrk[1510001:1520000,]
vgtrk31b = vgtrk[1520001:1530000,]
vgtrk31c = vgtrk[1530001:1540000,]
vgtrk31d = vgtrk[1540001:1550000,]
vgtrk31e = vgtrk[1550001:1560000,]
vgtrk32a = vgtrk[1560001:1570000,]
vgtrk32b = vgtrk[1570001:1580000,]
vgtrk32c = vgtrk[1580001:1590000,]
vgtrk32d = vgtrk[1590001:1600000,]
vgtrk32e = vgtrk[1600001:1610000,]

vgtrk33a = vgtrk[1610001:1620000,]
vgtrk33b = vgtrk[1620001:1630000,]
vgtrk33c = vgtrk[1630001:1640000,]
vgtrk33d = vgtrk[1640001:1650000,]
vgtrk33e = vgtrk[1650001:1660000,]
vgtrk34a = vgtrk[1660001:1670000,]
vgtrk34b = vgtrk[1670001:1680000,]
vgtrk34c = vgtrk[1680001:1690000,]
vgtrk34d = vgtrk[1690001:1700000,]
vgtrk34e = vgtrk[1700001:1710000,]

vgtrk35a = vgtrk[1710001:1720000,]
vgtrk35b = vgtrk[1720001:1730000,]
vgtrk35c = vgtrk[1730001:1740000,]
vgtrk35d = vgtrk[1740001:1750000,]
vgtrk35e = vgtrk[1750001:1760000,]
vgtrk36a = vgtrk[1760001:1770000,]
vgtrk36b = vgtrk[1770001:1777110,]

rm(vgtrk)
gc()

### Loading a spam-only filter ###

spam_only <- read.xlsx("C:/Users/Savanna/Desktop/Code_projects/R_projects/Text_mining_with_R/ru_spam.xlsx", colNames = TRUE)
spam_vector <- spam_only$value
spam <- str_trim(spam_vector, side = "both")
filter_spam <- str_c(spam, collapse = "|")

###### Steps for spam-filtering the data: ###########

### First: You need a list of dataframes from 1a to 36b 

vgtrk_names2 <- lapply(ls(pattern="^vgtrk[0-9]{1,2}[a-e]$"), function(x) get (x))

## Second: You need lapply() to filter each df for spam

vgtrk_spam_filtered <- lapply(vgtrk_names2, function(x) x %>% filter(!str_detect(X8, regex("spam", ignore_case = TRUE)) & !str_detect(X8, regex("mailto:support@vgtrk.com", ignore_case = TRUE)) & !str_detect(X7, regex(filter_spam, ignore_case = TRUE)) & !str_detect(X6, regex(filter_spam, ignore_case = TRUE))))

## Third: You need to condense the data into fewer tibbles and give them names

vgtrk1st100k_spam_filtered <- rbind(vgtrk_spam_filtered[[1]], vgtrk_spam_filtered[[2]], vgtrk_spam_filtered[[3]], vgtrk_spam_filtered[[4]], vgtrk_spam_filtered[[5]], vgtrk_spam_filtered[[6]], vgtrk_spam_filtered[[7]], vgtrk_spam_filtered[[8]], vgtrk_spam_filtered[[9]], vgtrk_spam_filtered[[10]])
vgtrk2nd100k_spam_filtered <- rbind(vgtrk_spam_filtered[[11]], vgtrk_spam_filtered[[12]], vgtrk_spam_filtered[[13]], vgtrk_spam_filtered[[14]], vgtrk_spam_filtered[[15]], vgtrk_spam_filtered[[16]], vgtrk_spam_filtered[[17]], vgtrk_spam_filtered[[18]], vgtrk_spam_filtered[[19]], vgtrk_spam_filtered[[20]])
vgtrk3rd100k_spam_filtered <- rbind(vgtrk_spam_filtered[[21]], vgtrk_spam_filtered[[22]], vgtrk_spam_filtered[[23]], vgtrk_spam_filtered[[24]], vgtrk_spam_filtered[[25]], vgtrk_spam_filtered[[26]], vgtrk_spam_filtered[[27]], vgtrk_spam_filtered[[28]], vgtrk_spam_filtered[[29]], vgtrk_spam_filtered[[30]])
vgtrk4th100k_spam_filtered <- rbind(vgtrk_spam_filtered[[31]], vgtrk_spam_filtered[[32]], vgtrk_spam_filtered[[33]], vgtrk_spam_filtered[[34]], vgtrk_spam_filtered[[35]], vgtrk_spam_filtered[[36]], vgtrk_spam_filtered[[37]], vgtrk_spam_filtered[[38]], vgtrk_spam_filtered[[39]], vgtrk_spam_filtered[[40]])
vgtrk5th100k_spam_filtered <- rbind(vgtrk_spam_filtered[[41]], vgtrk_spam_filtered[[42]], vgtrk_spam_filtered[[43]], vgtrk_spam_filtered[[44]], vgtrk_spam_filtered[[45]], vgtrk_spam_filtered[[46]], vgtrk_spam_filtered[[47]], vgtrk_spam_filtered[[48]], vgtrk_spam_filtered[[49]], vgtrk_spam_filtered[[50]])
vgtrk6th100k_spam_filtered <- rbind(vgtrk_spam_filtered[[51]], vgtrk_spam_filtered[[52]], vgtrk_spam_filtered[[53]], vgtrk_spam_filtered[[54]], vgtrk_spam_filtered[[55]], vgtrk_spam_filtered[[56]], vgtrk_spam_filtered[[57]], vgtrk_spam_filtered[[58]], vgtrk_spam_filtered[[59]], vgtrk_spam_filtered[[60]])
vgtrk7th100k_spam_filtered <- rbind(vgtrk_spam_filtered[[61]], vgtrk_spam_filtered[[62]], vgtrk_spam_filtered[[63]], vgtrk_spam_filtered[[64]], vgtrk_spam_filtered[[65]], vgtrk_spam_filtered[[66]], vgtrk_spam_filtered[[67]], vgtrk_spam_filtered[[68]], vgtrk_spam_filtered[[69]], vgtrk_spam_filtered[[70]])
vgtrk8th100k_spam_filtered <- rbind(vgtrk_spam_filtered[[71]], vgtrk_spam_filtered[[72]], vgtrk_spam_filtered[[73]], vgtrk_spam_filtered[[74]], vgtrk_spam_filtered[[75]], vgtrk_spam_filtered[[76]], vgtrk_spam_filtered[[77]], vgtrk_spam_filtered[[78]], vgtrk_spam_filtered[[79]], vgtrk_spam_filtered[[80]])
vgtrk9th100k_spam_filtered <- rbind(vgtrk_spam_filtered[[81]], vgtrk_spam_filtered[[82]], vgtrk_spam_filtered[[83]], vgtrk_spam_filtered[[84]], vgtrk_spam_filtered[[85]], vgtrk_spam_filtered[[86]], vgtrk_spam_filtered[[87]], vgtrk_spam_filtered[[88]], vgtrk_spam_filtered[[89]], vgtrk_spam_filtered[[90]])
vgtrk10th100k_spam_filtered <- rbind(vgtrk_spam_filtered[[91]], vgtrk_spam_filtered[[92]], vgtrk_spam_filtered[[93]], vgtrk_spam_filtered[[94]], vgtrk_spam_filtered[[95]], vgtrk_spam_filtered[[96]], vgtrk_spam_filtered[[97]], vgtrk_spam_filtered[[98]], vgtrk_spam_filtered[[99]], vgtrk_spam_filtered[[100]])
vgtrk11th100k_spam_filtered <- rbind(vgtrk_spam_filtered[[101]], vgtrk_spam_filtered[[102]], vgtrk_spam_filtered[[103]], vgtrk_spam_filtered[[104]], vgtrk_spam_filtered[[105]], vgtrk_spam_filtered[[106]], vgtrk_spam_filtered[[107]], vgtrk_spam_filtered[[108]], vgtrk_spam_filtered[[109]], vgtrk_spam_filtered[[110]])
vgtrk12th100k_spam_filtered <- rbind(vgtrk_spam_filtered[[111]], vgtrk_spam_filtered[[112]], vgtrk_spam_filtered[[113]], vgtrk_spam_filtered[[114]], vgtrk_spam_filtered[[115]], vgtrk_spam_filtered[[116]], vgtrk_spam_filtered[[117]], vgtrk_spam_filtered[[118]], vgtrk_spam_filtered[[119]], vgtrk_spam_filtered[[120]])
vgtrk13th100k_spam_filtered <- rbind(vgtrk_spam_filtered[[121]], vgtrk_spam_filtered[[122]], vgtrk_spam_filtered[[123]], vgtrk_spam_filtered[[124]], vgtrk_spam_filtered[[125]], vgtrk_spam_filtered[[126]], vgtrk_spam_filtered[[127]], vgtrk_spam_filtered[[128]], vgtrk_spam_filtered[[129]], vgtrk_spam_filtered[[130]])
vgtrk14th100k_spam_filtered <- rbind(vgtrk_spam_filtered[[131]], vgtrk_spam_filtered[[132]], vgtrk_spam_filtered[[133]], vgtrk_spam_filtered[[134]], vgtrk_spam_filtered[[135]], vgtrk_spam_filtered[[136]], vgtrk_spam_filtered[[137]], vgtrk_spam_filtered[[138]], vgtrk_spam_filtered[[139]], vgtrk_spam_filtered[[140]])
vgtrk15th100k_spam_filtered <- rbind(vgtrk_spam_filtered[[141]], vgtrk_spam_filtered[[142]], vgtrk_spam_filtered[[143]], vgtrk_spam_filtered[[144]], vgtrk_spam_filtered[[145]], vgtrk_spam_filtered[[146]], vgtrk_spam_filtered[[147]], vgtrk_spam_filtered[[148]], vgtrk_spam_filtered[[149]], vgtrk_spam_filtered[[150]])
vgtrk16th100k_spam_filtered <- rbind(vgtrk_spam_filtered[[151]], vgtrk_spam_filtered[[152]], vgtrk_spam_filtered[[153]], vgtrk_spam_filtered[[154]], vgtrk_spam_filtered[[155]], vgtrk_spam_filtered[[156]], vgtrk_spam_filtered[[157]], vgtrk_spam_filtered[[158]], vgtrk_spam_filtered[[159]], vgtrk_spam_filtered[[160]])
vgtrk17th100k_spam_filtered <- rbind(vgtrk_spam_filtered[[161]], vgtrk_spam_filtered[[162]], vgtrk_spam_filtered[[163]], vgtrk_spam_filtered[[164]], vgtrk_spam_filtered[[165]], vgtrk_spam_filtered[[166]], vgtrk_spam_filtered[[167]], vgtrk_spam_filtered[[168]], vgtrk_spam_filtered[[169]], vgtrk_spam_filtered[[170]])
vgtrk18th100k_spam_filtered <- rbind(vgtrk_spam_filtered[[171]], vgtrk_spam_filtered[[172]], vgtrk_spam_filtered[[173]], vgtrk_spam_filtered[[174]], vgtrk_spam_filtered[[175]], vgtrk_spam_filtered[[176]], vgtrk_spam_filtered[[177]])


## Fourth: you need to write the condensed data into xlsx

write.xlsx(vgtrk1st100k_spam_filtered, "vgtrk1st100k_spam_filtered.xlsx")
write.xlsx(vgtrk2nd100k_spam_filtered, "vgtrk2nd100k_spam_filtered.xlsx")
write.xlsx(vgtrk3rd100k_spam_filtered, "vgtrk3rd100k_spam_filtered.xlsx")
write.xlsx(vgtrk4th100k_spam_filtered, "vgtrk4th100k_spam_filtered.xlsx")
write.xlsx(vgtrk5th100k_spam_filtered, "vgtrk5th100k_spam_filtered.xlsx")
write.xlsx(vgtrk6th100k_spam_filtered, "vgtrk6th100k_spam_filtered.xlsx")
write.xlsx(vgtrk7th100k_spam_filtered, "vgtrk7th100k_spam_filtered.xlsx")
write.xlsx(vgtrk8th100k_spam_filtered, "vgtrk8th100k_spam_filtered.xlsx")
write.xlsx(vgtrk9th100k_spam_filtered, "vgtrk9th100k_spam_filtered.xlsx")
write.xlsx(vgtrk10th100k_spam_filtered, "vgtrk10th100k_spam_filtered.xlsx")
write.xlsx(vgtrk11th100k_spam_filtered, "vgtrk11th100k_spam_filtered.xlsx")
write.xlsx(vgtrk12th100k_spam_filtered, "vgtrk12th100k_spam_filtered.xlsx")
write.xlsx(vgtrk13th100k_spam_filtered, "vgtrk13th100k_spam_filtered.xlsx")
write.xlsx(vgtrk14th100k_spam_filtered, "vgtrk14th100k_spam_filtered.xlsx")
write.xlsx(vgtrk15th100k_spam_filtered, "vgtrk15th100k_spam_filtered.xlsx")
write.xlsx(vgtrk16th100k_spam_filtered, "vgtrk16th100k_spam_filtered.xlsx")
write.xlsx(vgtrk17th100k_spam_filtered, "vgtrk17th100k_spam_filtered.xlsx")
write.xlsx(vgtrk18th100k_spam_filtered, "vgtrk18th100k_spam_filtered.xlsx")

## Next make a list of spam-filtered dataframes:

vgtrk_nospam_names <- lapply(ls(pattern="vgtrk[0-9]+[a-z]{2}100k_spam_filtered$"), function(x) get (x))

## Now you can filter the spam-screened emails by keyword: ##

vgtrk_ibans <- lapply(vgtrk_nospam_names, function(x) x %>% filter(str_detect(X7, regex("[A-Z]{2}[0-9]{2} ?[0-9]{4} ?[0-9]{4} ?[0-9]{4} ?[0-9]{4} ?[0-9]{0,2}|IBAN", ignore_case = TRUE))))
vgtrk_ibans_df <- rbind(vgtrk_ibans[[1]], vgtrk_ibans[[2]], vgtrk_ibans[[3]], vgtrk_ibans[[4]], vgtrk_ibans[[5]], vgtrk_ibans[[6]], vgtrk_ibans[[7]], vgtrk_ibans[[8]], vgtrk_ibans[[9]], vgtrk_ibans[[10]], vgtrk_ibans[[11]], vgtrk_ibans[[12]], vgtrk_ibans[[13]], vgtrk_ibans[[14]], vgtrk_ibans[[15]], vgtrk_ibans[[16]], vgtrk_ibans[[17]], vgtrk_ibans[[18]])
write.xlsx(vgtrk_ibans_df, "vgtrk_ibans.xlsx")

vgtrk_iban_destination <- lapply(vgtrk_nospam_names, function(x) x %>% filter(str_detect(X4, regex("wkoswig@mx13.mail.ru|zapadrtr@mail.ru", ignore_case = TRUE))))
vgtrk_iban_destination <- vgtrk_iban_destination[sapply(vgtrk_iban_destination, function(x) dim(x)[1]) > 0]
vgtrk_iban_destination_df <- rbind(vgtrk_iban_destination[[1]], vgtrk_iban_destination[[2]])
write.xlsx(vgtrk_iban_destination_df, "vgtrk_iban_destination_df.xlsx")

ru_corrupt <- read.xlsx("C:/Users/Savanna/Desktop/Code_projects/R_projects/Text_mining_with_R/ru_corruption.xlsx", colNames = TRUE)
corrupt_vector <- ru_corrupt$value
corruption <- str_trim(corrupt_vector, side = "both")
filter_corrupt <- str_c(corruption, collapse = "|")

vgtrk_corruption <- lapply(vgtrk_nospam_names, function(x) x %>% filter(str_detect(X7, regex(filter_corrupt, ignore_case = FALSE))))
vgtrk_corruption_df <- rbind(vgtrk_corruption[[1]], vgtrk_corruption[[2]], vgtrk_corruption[[3]], vgtrk_corruption[[4]], vgtrk_corruption[[5]], vgtrk_corruption[[6]], vgtrk_corruption[[7]], vgtrk_corruption[[8]], vgtrk_corruption[[9]], vgtrk_corruption[[10]], vgtrk_corruption[[11]], vgtrk_corruption[[12]], vgtrk_corruption[[13]], vgtrk_corruption[[14]], vgtrk_corruption[[15]], vgtrk_corruption[[16]], vgtrk_corruption[[17]], vgtrk_corruption[[18]])
write.xlsx(vgtrk_corruption_df, "vgtrk_corruption.xlsx")

vgtrk_chomsky <- lapply(vgtrk_nospam_names, function(x) x %>% filter(str_detect(X7, regex("chomsky|Хомский|чомский|чомзкий|хомзкий", ignore_case = TRUE))))
vgtrk_chomsky_df <- rbind(vgtrk_chomsky[[1]], vgtrk_chomsky[[2]], vgtrk_chomsky[[3]], vgtrk_chomsky[[4]], vgtrk_chomsky[[5]], vgtrk_chomsky[[6]], vgtrk_chomsky[[7]], vgtrk_chomsky[[8]], vgtrk_chomsky[[9]], vgtrk_chomsky[[10]], vgtrk_chomsky[[11]], vgtrk_chomsky[[12]], vgtrk_chomsky[[13]], vgtrk_chomsky[[14]], vgtrk_chomsky[[15]], vgtrk_chomsky[[16]], vgtrk_chomsky[[17]], vgtrk_chomsky[[18]])
write.xlsx(vgtrk_chomsky_df, "vgtrk_chomsky.xlsx")

vgtrk_framing <- lapply(vgtrk_nospam_names, function(x) x %>% filter(str_detect(X7, regex(" я |!!", ignore_case = TRUE)) & !str_detect(X7, regex("hyperic", ignore_case = TRUE))))
### This one is problematic: vgtrk_framing[[3]] won't write to xlsx
vgtrk_framing_df <- rbind(vgtrk_framing[[1]], vgtrk_framing[[2]], vgtrk_framing[[4]], vgtrk_framing[[5]], vgtrk_framing[[6]], vgtrk_framing[[7]], vgtrk_framing[[8]], vgtrk_framing[[9]], vgtrk_framing[[10]], vgtrk_framing[[11]], vgtrk_framing[[12]], vgtrk_framing[[13]], vgtrk_framing[[14]], vgtrk_framing[[15]], vgtrk_framing[[16]], vgtrk_framing[[17]], vgtrk_framing[[18]])
write.xlsx(vgtrk_framing_df, "vgtrk_framing.xlsx")


ru_navalny <- read.xlsx("C:/Users/Savanna/Desktop/Code_projects/R_projects/Text_mining_with_R/ru_navalny.xlsx", colNames = TRUE)
navalny <- ru_navalny$value %>% str_trim(., side = "both")
filter_navalny <- str_c(navalny, collapse = "|")

vgtrk_navalny <- lapply(vgtrk_nospam_names, function(x) x %>% filter(str_detect(X7, regex(filter_navalny, ignore_case = TRUE))))
vgtrk_navalny_df <- rbind(vgtrk_navalny[[1]], vgtrk_navalny[[2]], vgtrk_navalny[[3]], vgtrk_navalny[[5]], vgtrk_navalny[[6]], vgtrk_navalny[[7]], vgtrk_navalny[[8]], vgtrk_navalny[[9]], vgtrk_navalny[[10]], vgtrk_navalny[[11]], vgtrk_navalny[[12]], vgtrk_navalny[[13]], vgtrk_navalny[[14]], vgtrk_navalny[[15]], vgtrk_navalny[[16]], vgtrk_navalny[[17]], vgtrk_navalny[[18]])
##  vgtrk_navalny[[4]] is corrupted
write.xlsx(vgtrk_navalny_df, "vgtrk_navalny.xlsx")


german_names <- read.xlsx("C:/Users/Savanna/Desktop/Code_projects/R_projects/Text_mining_with_R/ru_german_names.xlsx", colNames = TRUE)
german_name <- german_names$Name
filter_german_names <- str_c(german_name, collapse = "|")

vgtrk_german_names <- lapply(vgtrk_nospam_names, function(x) x %>% filter(str_detect(X7, regex(filter_german_names, ignore_case = TRUE))))
vgtrk_german_names_df <- rbind(vgtrk_german_names[[1]], vgtrk_german_names[[2]], vgtrk_german_names[[3]], vgtrk_german_names[[4]], vgtrk_german_names[[5]], vgtrk_german_names[[6]], vgtrk_german_names[[7]], vgtrk_german_names[[8]], vgtrk_german_names[[9]], vgtrk_german_names[[10]], vgtrk_german_names[[11]], vgtrk_german_names[[12]], vgtrk_german_names[[13]], vgtrk_german_names[[14]], vgtrk_german_names[[15]], vgtrk_german_names[[16]], vgtrk_german_names[[17]], vgtrk_german_names[[18]])
write.xlsx(vgtrk_german_names, "vgtrk_german_names.xlsx")

vgtrk_de <- lapply(vgtrk_nospam_names, function(x) x %>% filter(str_detect(X4, regex(".+@.+de$", ignore_case = TRUE))))
vgtrk_de_to_df <- rbind(vgtrk_de[[1]], vgtrk_de[[2]], vgtrk_de[[3]], vgtrk_de[[4]], vgtrk_de[[5]], vgtrk_de[[6]], vgtrk_de[[7]], vgtrk_de[[8]], vgtrk_de[[9]], vgtrk_de[[10]], vgtrk_de[[11]], vgtrk_de[[12]], vgtrk_de[[13]], vgtrk_de[[14]], vgtrk_de[[15]], vgtrk_de[[16]], vgtrk_de[[17]], vgtrk_de[[18]])


vgtrk_de <- lapply(vgtrk_nospam_names, function(x) x %>% filter(str_detect(X1, regex(".+@.+de$", ignore_case = TRUE))))
vgtrk_de_from_df <- rbind(vgtrk_de[[1]], vgtrk_de[[2]], vgtrk_de[[3]], vgtrk_de[[4]], vgtrk_de[[5]], vgtrk_de[[6]], vgtrk_de[[7]], vgtrk_de[[8]], vgtrk_de[[9]], vgtrk_de[[10]], vgtrk_de[[11]], vgtrk_de[[12]], vgtrk_de[[13]], vgtrk_de[[14]], vgtrk_de[[15]], vgtrk_de[[16]], vgtrk_de[[17]], vgtrk_de[[18]])

ru_WWII <- read.xlsx("C:/Users/Savanna/Desktop/Code_projects/R_projects/Text_mining_with_R/ru_WWII.xlsx", colNames = TRUE)
WWII <- ru_WWII$value
WWII <- str_trim(WWII, side = "both")
filter_WWII <- str_c(WWII, collapse = "|")

vgtrk_WWII <- lapply(vgtrk_nospam_names, function(x) x %>% filter(str_detect(X7, regex(filter_WWII, ignore_case = TRUE))))
### each of vgtrk_WWII[[4]], vgtrk_WWII[[7]], and vgtrk_WWII[[9]] has a corrupted file
vgtrk_WWII_df <- rbind(vgtrk_WWII[[1]], vgtrk_WWII[[2]], vgtrk_WWII[[3]], vgtrk_WWII[[5]], vgtrk_WWII[[6]], vgtrk_WWII[[8]], vgtrk_WWII[[10]], vgtrk_WWII[[11]], vgtrk_WWII[[12]], vgtrk_WWII[[13]], vgtrk_WWII[[14]], vgtrk_WWII[[15]], vgtrk_WWII[[16]], vgtrk_WWII[[17]], vgtrk_WWII[[18]])
write.xlsx(vgtrk_WWII_df, "vgtrk_WWII.xlsx")

ru_balticNATO <- read.xlsx("C:/Users/Savanna/Desktop/Code_projects/R_projects/Text_mining_with_R/ru_NATOBaltics.xlsx", colNames = TRUE)
balticNATO <- ru_balticNATO$value
balticNATO <- str_trim(balticNATO, side = "both")
filter_balticNATO <- str_c(balticNATO, collapse = "|")

vgtrk_baltic_nato <- lapply(vgtrk_nospam_names, function(x) x %>% filter(str_detect(X7, regex(filter_balticNATO, ignore_case = TRUE))))
##  vgtrk_baltic_nato[[4]], vgtrk_baltic_nato[[7]], and vgtrk_baltic_nato_df[[8]] each has a faulty file
vgtrk_baltic_nato_df <- rbind(vgtrk_baltic_nato[[1]], vgtrk_baltic_nato[[2]], vgtrk_baltic_nato[[3]], vgtrk_baltic_nato[[5]],  vgtrk_baltic_nato[[6]], vgtrk_baltic_nato[[10]], vgtrk_baltic_nato[[11]], vgtrk_baltic_nato[[12]], vgtrk_baltic_nato[[13]], vgtrk_baltic_nato[[14]], vgtrk_baltic_nato[[15]], vgtrk_baltic_nato[[16]],vgtrk_baltic_nato[[17]], vgtrk_baltic_nato[[18]])
write.xlsx(vgtrk_baltic_nato_df, "vgtrk_baltic_nato.xlsx")

vgtrk_de <- lapply(vgtrk_nospam_names, function(x) x %>% filter(str_detect(X4, regex(".+@.+de$", ignore_case = TRUE))))
vgtrk_de_to_df <- rbind(vgtrk_de[[1]], vgtrk_de[[2]], vgtrk_de[[3]], vgtrk_de[[4]], vgtrk_de[[5]], vgtrk_de[[6]], vgtrk_de[[7]], vgtrk_de[[8]], vgtrk_de[[9]], vgtrk_de[[10]], vgtrk_de[[11]], vgtrk_de[[12]], vgtrk_de[[13]], vgtrk_de[[14]], vgtrk_de[[15]], vgtrk_de[[16]], vgtrk_de[[17]], vgtrk_de[[18]])
write.xlsx(vgtrk_de_to_df, "vgtrk_de_to.xlsx")

ru_ukraine <- read.xlsx("C:/Users/Savanna/Desktop/Code_projects/R_projects/Text_mining_with_R/ru_ukraine.xlsx", colNames = TRUE)
ukraine_vector <- ru_ukraine$value
ukraine <- str_trim(ukraine_vector, side = "both")
filter_ukraine <- str_c(ukraine, collapse = "|")

vgtrk_ukraine_framing <- vgtrk_framing_df %>% filter(str_detect(X6, regex(filter_ukraine, ignore_case = TRUE)) | str_detect(X7, regex(filter_ukraine, ignore_case = TRUE)))
write.xlsx(vgtrk_ukraine_framing, "vgtrk_ukraine_framing.xlsx")


vgtrk_ukraine <- lapply(vgtrk_nospam_names, function(x) x %>% filter(str_detect(X6, regex(filter_ukraine, ignore_case = TRUE)) | str_detect(X7, regex(filter_ukraine, ignore_case = TRUE))))
vgtrk_ukraine_df <- rbind(vgtrk_ukraine[[1]], vgtrk_ukraine[[2]], vgtrk_ukraine[[3]], vgtrk_ukraine[[4]], vgtrk_ukraine[[5]], vgtrk_ukraine[[6]], vgtrk_ukraine[[7]], vgtrk_ukraine[[8]], vgtrk_ukraine[[9]], vgtrk_ukraine[[10]], vgtrk_ukraine[[11]], vgtrk_ukraine[[12]], vgtrk_ukraine[[13]], vgtrk_ukraine[[14]], vgtrk_ukraine[[15]], vgtrk_ukraine[[16]], vgtrk_ukraine[[17]], vgtrk_ukraine[[18]])
write.xlsx(vgtrk_ukraine_df, "vgtrk_ukraine.xlsx")


#### To perform LDA / sentiment analysis on keyword-filtered emails ####

## First import stopwords and sentiment lexicon ##

ru_stopwords <- read.xlsx("C:/Users/Savanna/Desktop/Code_projects/R_projects/Text_mining_with_R/ru_stopwords.xlsx", colNames = TRUE)

ru_sentiment <- read.xlsx("C:/Users/Savanna/Desktop/Code_projects/R_projects/Text_mining_with_R/ru_sentiment.xlsx", colNames = TRUE)

## Next unnest tokens & remove stopwords from each group of emails ##

vgtrk_ukraine_words <- vgtrk_ukraine_df %>% unnest_tokens("words", X7, to_lower = TRUE)
vgtrk_ukraine_keywords <- vgtrk_ukraine_words %>% stringdist_anti_join(ru_stopwords, c("words" = "words"))

de_words <- vgtrk_de_to_df %>% unnest_tokens("words", X7, to_lower = TRUE)
de_keywords <- de_words %>% stringdist_anti_join(ru_stopwords, c("words" = "words"))

## Next take counts of key words ##

ukraine_wordcounts <- vgtrk_ukraine_framing_keywords %>% 
  count(X1, words, sort = TRUE)

## Now you can perform LDA on unigrams ##

ukraine_dtm <- ukraine_wordcounts %>% cast_dtm(X1, words, n)

ukraine_lda <- LDA(ukraine_dtm, k=4, control = list(seed = 1234))
ukraine_topics <- tidy(ukraine_lda, matrix = "beta")
top_terms_ukraine <- ukraine_topics %>% group_by(topic) %>% slice_max(
  beta, n=15) %>% ungroup() %>% arrange(topic, -beta)

top_terms_ukraine %>% mutate(term = reorder_within(term, beta, topic)) %>%
  ggplot(aes(beta, term, fill = factor(topic))) +
  geom_col(show.legend = FALSE) +
  facet_wrap(~topic, scales = "free") +
  scale_y_reordered()

## In this step you manually translate the top words and reimport ##

write.xlsx(top_terms_ukraine, "top_terms_ukraine.xlsx")
top_terms_ukraine <- read.xlsx("top_terms_ukraine.xlsx")

## You can also look at LDA topics by sender (or other grouping variable)

ukraine_gamma <- tidy(ukraine_lda, matrix = "gamma")
ukraine_gamma_top <- filter(ukraine_gamma, gamma > 0.5)
ukraine_gamma %>% 
  mutate(X1 = reorder(document, gamma * topic)) %>%
  ggplot(aes(factor(topic), gamma)) +
  geom_boxplot() +
  facet_wrap(~ document) +
  labs(x = "topic", y = expression(gamma))

## To tokenize with bigrams instead, removing stop-words is tricky: ##

vgtrk_ukraine_bigrams <- vgtrk_ukraine_df %>% 
  unnest_tokens(bigram, X7, token = "ngrams", n = 2)

vgtrk_ukraine_bigrams_sep <- vgtrk_ukraine_bigrams %>% 
  separate(bigram, c("word1", "word2"), sep = " ")
vgtrk_ukraine_keybigrams <- vgtrk_ukraine_bigrams_sep %>% 
  stringdist_anti_join(ru_stopwords, c("word1" = "words")) %>% 
  stringdist_anti_join(ru_stopwords, c("word2" = "words"))
ukraine_bigrams_key <- vgtrk_ukraine_keybigrams %>% 
  unite(bigram, word1, word2, sep = " ")
ukraine_bigram_counts <- ukraine_bigrams_key %>% 
  count(X1, bigram, sort = TRUE)

## Topic modeling gives more interesting results with bigrams:

ukraine_bigram_dtm <- ukraine_bigram_counts %>% cast_dtm(X1, bigram, n)

ukraine_bigram_lda <- LDA(
  ukraine_bigram_dtm, k=6, control = list(seed = 1234))
ukraine_bigram_topics <- tidy(ukraine_bigram_lda, matrix = "beta")

top_bigrams_ukraine <- ukraine_bigram_topics %>% 
  group_by(topic) %>% 
  slice_max(beta, n=15) %>% ungroup() %>% arrange(topic, -beta)

top_bigrams_ukraine %>% 
  mutate(term = reorder_within(term, beta, topic)) %>%
  ggplot(aes(beta, term, fill = factor(topic))) +
  geom_col(show.legend = FALSE) +
  facet_wrap(~topic, scales = "free") +
  scale_y_reordered()

## Again you manually translate and reimport the key bigrams ##

write.xlsx(top_bigrams_ukraine, "top_bigrams_ukraine.xlsx")
top_bigrams_ukraine <- read.xlsx("top_bigrams_ukraine.xlsx")

## Again you can look at the topics by sender ##

ukraine_bigrams_gamma <- tidy(ukraine_bigram_lda, matrix = "gamma")
ukraine_bigrams_gamma_top <- filter(ukraine_gamma, gamma > 0.3)

ukraine_bigrams_gamma_top %>% 
  mutate(X1 = reorder(document, gamma * topic)) %>%
  ggplot(aes(factor(topic), gamma)) +
  geom_boxplot() +
  facet_wrap(~ document) +
  labs(x = "topic", y = expression(gamma))

### Visualizing topics for Baltic/NATO states emails ####

vgtrk_baltic_nato_bigrams <- vgtrk_baltic_nato_df %>% 
  unnest_tokens(bigram, X7, token = "ngrams", n = 2)

vgtrk_baltic_nato_bigrams_sep <- vgtrk_baltic_nato_bigrams %>% 
  separate(bigram, c("word1", "word2"), sep = " ")
vgtrk_baltic_nato_keybigrams <- vgtrk_baltic_nato_bigrams_sep %>% 
  stringdist_anti_join(ru_stopwords, c("word1" = "words")) %>% 
  stringdist_anti_join(ru_stopwords, c("word2" = "words"))
baltic_nato_bigrams_key <- vgtrk_baltic_nato_keybigrams %>% 
  unite(bigram, word1, word2, sep = " ")
baltic_nato_bigram_counts <- baltic_nato_bigrams_key %>% 
  count(X1, bigram, sort = TRUE)

baltic_nato_bigram_dtm <- baltic_nato_bigram_counts %>% cast_dtm(X1, bigram, n)

baltic_nato_bigram_lda <- LDA(
  baltic_nato_bigram_dtm, k=6, control = list(seed = 1234))
baltic_nato_bigram_topics <- tidy(baltic_nato_bigram_lda, matrix = "beta")

top_bigrams_baltic_nato <- baltic_nato_bigram_topics %>% 
  group_by(topic) %>% 
  slice_max(beta, n=15) %>% ungroup() %>% arrange(topic, -beta)

top_bigrams_baltic_nato %>% 
  mutate(term = reorder_within(term, beta, topic)) %>%
  ggplot(aes(beta, term, fill = factor(topic))) +
  geom_col(show.legend = FALSE) +
  facet_wrap(~topic, scales = "free") +
  scale_y_reordered()

write.xlsx(top_bigrams_baltic_nato, "top_bigrams_baltic_nato.xlsx")
top_bigrams_baltic_nato <- read.xlsx("top_bigrams_baltic_nato.xlsx")

top_bigrams_baltic_nato %>% 
  mutate(term = reorder_within(term, beta, topic)) %>%
  ggplot(aes(beta, translated, fill = factor(topic))) +
  geom_col(show.legend = FALSE) +
  facet_wrap(~topic, scales = "free") +
  scale_y_reordered()

baltic_nato_bigrams_gamma <- tidy(baltic_nato_bigram_lda, matrix = "gamma")
baltic_nato_bigrams_gamma_top <- filter(baltic_nato_gamma, gamma > 0.3)

baltic_nato_bigrams_gamma_top %>% 
  mutate(X1 = reorder(document, gamma * topic)) %>%
  ggplot(aes(factor(topic), gamma)) +
  geom_boxplot() +
  facet_wrap(~ document) +
  labs(x = "topic", y = expression(gamma))


### Visualizing topics for WWII emails ###

vgtrk_WWII_bigrams <- vgtrk_WWII_df %>% 
  unnest_tokens(bigram, X7, token = "ngrams", n = 2)

vgtrk_WWII_bigrams_sep <- vgtrk_WWII_bigrams %>% 
  separate(bigram, c("word1", "word2"), sep = " ")
vgtrk_WWII_keybigrams <- vgtrk_WWII_bigrams_sep %>% 
  stringdist_anti_join(ru_stopwords, c("word1" = "words")) %>% 
  stringdist_anti_join(ru_stopwords, c("word2" = "words"))
WWII_bigrams_key <- vgtrk_WWII_keybigrams %>% 
  unite(bigram, word1, word2, sep = " ")
WWII_bigrams_counts <- WWII_bigrams_key %>% 
  count(X1, bigram, sort = TRUE)

WWII_bigram_dtm <- WWII_bigrams_counts %>% cast_dtm(X1, bigram, n)

WWII_bigram_lda <- LDA(
  WWII_bigram_dtm, k=6, control = list(seed = 1234))
WWII_bigram_topics <- tidy(WWII_bigram_lda, matrix = "beta")

top_bigrams_WWII <- WWII_bigram_topics %>% 
  group_by(topic) %>% 
  slice_max(beta, n=15) %>% ungroup() %>% arrange(topic, -beta)

top_bigrams_WWII %>% 
  mutate(term = reorder_within(term, beta, topic)) %>%
  ggplot(aes(beta, term, fill = factor(topic))) +
  geom_col(show.legend = FALSE) +
  facet_wrap(~topic, scales = "free") +
  scale_y_reordered()

write.xlsx(top_bigrams_WWII, "top_bigrams_WWII.xlsx")
top_bigrams_WWII <- read.xlsx("top_bigrams_WWII.xlsx")

top_bigrams_WWII %>% 
  mutate(term = reorder_within(term, beta, topic)) %>%
  ggplot(aes(beta, translated, fill = factor(topic))) +
  geom_col(show.legend = FALSE) +
  facet_wrap(~topic, scales = "free") +
  scale_y_reordered()

WWII_bigrams_gamma <- tidy(WWII_bigram_lda, matrix = "gamma")

WWII_bigrams_gamma %>% 
  mutate(X1 = reorder(document, gamma * topic)) %>%
  ggplot(aes(factor(topic), gamma)) +
  geom_boxplot() +
  facet_wrap(~ document) +
  labs(x = "topic", y = expression(gamma))

vgtrk_WWII_words <- vgtrk_WWII_df %>% 
  unnest_tokens(word, X7, "words")

vgtrk_WWII_keywords <- vgtrk_WWII_words %>% 
  stringdist_anti_join(ru_stopwords, c("word" = "words"))

## You can also do sentiment analysis with unigrams

ru_sentiment <-  read.xlsx("C:/Users/Savanna/Desktop/Code_projects/R_projects/Text_mining_with_R/ru_sentiment_simplified.xlsx", colNames = TRUE)

WWII_sentiment_counts <- vgtrk_WWII_keywords %>%
  full_join(ru_sentiment, by = c("word" = "word"), copy = TRUE) %>%
  count(word, affect, sort = TRUE) %>%
  ungroup()

WWII_sentiment <- WWII_sentiment_counts %>%
  group_by(affect) %>%
  slice_max(n, n = 15) %>% 
  ungroup() %>%
  mutate(word = reorder(word, n))

## Again there is a manual translation step and you reimport the data:

write.xlsx(WWII_sentiment, "WWII_sentiment.xlsx")
WWII_sentiment <- read.xlsx("WWII_sentiment.xlsx")

WWII_sentiment %>%
  ggplot(aes(n, translated, fill = affect)) +
  geom_col(show.legend = FALSE) +
  facet_wrap(~affect, scales = "free_y") +
  labs(x = "Contribution to sentiment",
       y = NULL)

### Visualizing topics/sentiment for emails matching German names list ###

german_names_bigrams <- vgtrk_german_names_df %>% 
  unnest_tokens(bigram, X7, token = "ngrams", n = 2)

german_names_bigrams_sep <- german_names_bigrams %>% 
  separate(bigram, c("word1", "word2"), sep = " ")
german_names_keybigrams <- german_names_bigrams_sep %>% 
  stringdist_anti_join(ru_stopwords, c("word1" = "words")) %>% 
  stringdist_anti_join(ru_stopwords, c("word2" = "words"))
german_names_bigrams_key <- german_names_keybigrams %>% 
  unite(bigram, word1, word2, sep = " ")

german_names_bigram_counts <- german_names_bigrams_key %>% 
  count(X1, bigram, sort = TRUE)

german_names_bigram_dtm <- german_names_bigram_counts %>% 
  cast_dtm(X1, bigram, n)

german_names_bigram_lda <- LDA(
  german_names_bigram_dtm, k=6, control = list(seed = 1234))
german_names_bigram_topics <- tidy(german_names_bigram_lda, matrix = "beta")
top_bigrams_german_names <- german_names_bigram_topics %>% 
  group_by(topic) %>% 
  slice_max(beta, n=15) %>% ungroup() %>% arrange(topic, -beta)

top_bigrams_german_names %>% 
  mutate(term = reorder_within(term, beta, topic)) %>%
  ggplot(aes(beta, term, fill = factor(topic))) +
  geom_col(show.legend = FALSE) +
  facet_wrap(~topic, scales = "free") +
  scale_y_reordered()

write.xlsx(top_bigrams_german_names, "top_bigrams_german_names.xlsx")
top_bigrams_german_names <- read.xlsx("top_bigrams_german_names.xlsx")

top_bigrams_german_names %>% 
  mutate(term = reorder_within(term, beta, topic)) %>%
  ggplot(aes(beta, translated, fill = factor(topic))) +
  geom_col(show.legend = FALSE) +
  facet_wrap(~topic, scales = "free") +
  scale_y_reordered()

german_names_bigrams_gamma <- tidy(german_names_bigram_lda, matrix = "gamma")

german_names_bigrams_gamma %>% 
  mutate(X1 = reorder(document, gamma * topic)) %>%
  ggplot(aes(factor(topic), gamma)) +
  geom_boxplot() +
  facet_wrap(~ document) +
  labs(x = "topic", y = expression(gamma))

## for sentiment analysis, it is easier to use unigrams

vgtrk_german_names_words <- vgtrk_german_names_df %>% 
  unnest_tokens(word, X7, "words")

vgtrk_german_names_keywords <- vgtrk_german_names_words %>% 
  stringdist_anti_join(ru_stopwords, c("word" = "words"))


german_names_sentiment_counts <- vgtrk_german_names_keywords %>%
  full_join(ru_sentiment, by = c("word" = "word"), copy = TRUE) %>%
  count(word, affect, sort = TRUE) %>%
  ungroup()

german_names_sentiment <- german_names_sentiment_counts %>%
  group_by(affect) %>%
  slice_max(n, n = 15) %>% 
  ungroup() %>%
  mutate(word = reorder(word, n))


write.xlsx(german_names_sentiment, "german_names_sentiment.xlsx")
german_names_sentiment <- read.xlsx("german_names_sentiment.xlsx")

german_names_sentiment %>%
  ggplot(aes(n, translated, fill = affect)) +
  geom_col(show.legend = FALSE) +
  facet_wrap(~affect, scales = "free_y") +
  labs(x = "Contribution to sentiment",
       y = NULL)


#### Visualizing bigrams and sentiments for Navalny emails ###

navalny_bigrams <- vgtrk_navalny_df %>% 
  unnest_tokens(bigram, X7, token = "ngrams", n = 2)

navalny_bigrams_sep <- navalny_bigrams %>% 
  separate(bigram, c("word1", "word2"), sep = " ")
navalny_keybigrams <- navalny_bigrams_sep %>% 
  stringdist_anti_join(ru_stopwords, c("word1" = "words")) %>% 
  stringdist_anti_join(ru_stopwords, c("word2" = "words"))
navalny_bigrams_key <- navalny_keybigrams %>% 
  unite(bigram, word1, word2, sep = " ")

navalny_bigram_counts <- navalny_bigrams_key %>% 
  count(X1, bigram, sort = TRUE)

navalny_bigram_dtm <- navalny_bigram_counts %>% 
  cast_dtm(X1, bigram, n)

navalny_bigram_lda <- LDA(
  navalny_bigram_dtm, k=6, control = list(seed = 1234))
navalny_bigram_topics <- tidy(navalny_bigram_lda, matrix = "beta")
top_bigrams_navalny <- navalny_bigram_topics %>% 
  group_by(topic) %>% 
  slice_max(beta, n=5) %>% ungroup() %>% arrange(topic, -beta)

top_bigrams_navalny %>% 
  mutate(term = reorder_within(term, beta, topic)) %>%
  ggplot(aes(beta, term, fill = factor(topic))) +
  geom_col(show.legend = FALSE) +
  facet_wrap(~topic, scales = "free") +
  scale_y_reordered()

write.xlsx(top_bigrams_navalny, "top_bigrams_navalny.xlsx")
top_bigrams_navalny <- read.xlsx("top_bigrams_navalny.xlsx")


top_bigrams_navalny %>% 
  mutate(term = reorder_within(term, beta, topic)) %>%
  ggplot(aes(beta, translated, fill = factor(topic))) +
  geom_col(show.legend = FALSE) +
  facet_wrap(~topic, scales = "free") +
  scale_y_reordered()


navalny_bigrams_gamma <- tidy(navalny_bigram_lda, matrix = "gamma")


navalny_bigrams_gamma %>% 
  mutate(X1 = reorder(document, gamma * topic)) %>%
  ggplot(aes(factor(topic), gamma)) +
  geom_boxplot() +
  facet_wrap(~ document) +
  labs(x = "topic", y = expression(gamma))



vgtrk_navalny_words <- vgtrk_navalny_df %>% 
  unnest_tokens(word, X7, "words")

vgtrk_navalny_keywords <- vgtrk_navalny_words %>% 
  stringdist_anti_join(ru_stopwords, c("word" = "words"))


navalny_sentiment_counts <- vgtrk_navalny_keywords %>%
  full_join(ru_sentiment, by = c("word" = "word"), copy = TRUE) %>%
  count(word, affect, sort = TRUE) %>%
  ungroup()

navalny_sentiment <- navalny_sentiment_counts %>%
  group_by(affect) %>%
  slice_max(n, n = 15) %>% 
  ungroup() %>%
  mutate(word = reorder(word, n))


write.xlsx(navalny_sentiment, "navalny_sentiment.xlsx")
navalny_sentiment <- read.xlsx("navalny_sentiment.xlsx")

navalny_sentiment %>%
  ggplot(aes(n, translated, fill = affect)) +
  geom_col(show.legend = FALSE) +
  facet_wrap(~affect, scales = "free_y") +
  labs(x = "Contribution to sentiment",
       y = NULL)


### Comparing the network structure of news emails ###

ru_news <- read.xlsx("C:/Users/Savanna/Desktop/Code_projects/R_projects/Text_mining_with_R/ru_news_content.xlsx", colNames = TRUE)
news_vector <- ru_news$value
news <- str_trim(news_vector, side = "both")
filter_news <- str_c(news, collapse = "|")

vgtrk_news <- lapply(vgtrk_nospam_names, function(x) x %>% filter(str_detect(X7, regex(filter_news, ignore_case = TRUE))))
vgtrk_news_df <- rbind(vgtrk_news[[1]], vgtrk_news[[2]], vgtrk_news[[3]], vgtrk_news[[4]], vgtrk_news[[5]], vgtrk_news[[6]], vgtrk_news[[7]], vgtrk_news[[8]], vgtrk_news[[9]], vgtrk_news[[10]], vgtrk_news[[11]], vgtrk_news[[12]], vgtrk_news[[13]], vgtrk_news[[14]], vgtrk_news[[15]], vgtrk_news[[16]], vgtrk_news[[17]], vgtrk_news[[18]])

mail_from_news <- vgtrk_news_df %>% count(X1, X4) %>% filter(n > 25)

mail_matrix <- mail_from_news %>% pivot_wider(., names_from = "X4", values_from = "n")
mail_matrix_df <- as.data.frame(mail_matrix)
mail_matrix_df[is.na(mail_matrix_df)] = 0

mail_nodes <- mail_matrix_df %>% 
  dplyr::mutate(From = X1,
                Occurrences = rowSums(mail_matrix_df[-1])) %>%
  dplyr::select(From, Occurrences) %>%
  dplyr::filter(!str_detect(From, "X4"))

mail_edges <- mail_matrix_df %>%
  dplyr::mutate(from = X1) %>%
  tidyr::gather(to, Frequency, "Katie.Fisher@hermitagefund.com":"inquiry@rfis.tradeindia.com") %>%
  dplyr::mutate(Frequency = ifelse(Frequency == 0, NA, Frequency))
mail_edges <- subset(mail_edges, select = -c(X1))
head(mail_edges)
news_network <- igraph::graph_from_data_frame(
  d = mail_edges, directed = TRUE)
tidy_news_network <- tidygraph::as_tbl_graph(news_network) %>%
  tidygraph::activate(edges) %>%
  dplyr::mutate(name=from)

E(tidy_news_network)$weight <- E(tidy_news_network)$Frequency
head(E(tidy_news_network)$weight, 10)

set.seed(12345)

tidy_news_network %>%
  ggraph(layout = "fr") +
  geom_edge_arc(
    colour = "gray50", 
    lineend = "round", 
    strength = 0.1, 
    aes(edge_width = weight, 
        alpha = weight)) +
  geom_node_text(aes(label = name),
                 repel = TRUE,
                 point.padding = unit(0.2, "lines"),
                 colour = "gray10") +
  theme_graph(background = "white") +
  guides(edge_width = FALSE,
         edge_alpha = FALSE)

### Anonymized image:

tidy_news_network %>%
  ggraph(layout = "fr") +
  geom_edge_arc(
    colour = "gray50", 
    lineend = "round", 
    strength = 0.1, 
    aes(edge_width = weight, 
        alpha = weight)) +
  theme_graph(background = "white") +
  guides(edge_width = FALSE,
         edge_alpha = FALSE)

### Same network graph for just the framing emails ####


news_framing_df <- vgtrk_news_df %>% filter(str_detect(X7, regex(" я |!!", ignore_case = TRUE)) & !str_detect(X7, regex("hyperic", ignore_case = TRUE)))

mail_from_news_framing <- news_framing_df %>% count(X1, X4) %>% filter(n > 2)

framing_mail_matrix <- mail_from_news_framing %>% pivot_wider(., names_from = "X4", values_from = "n")
framing_mail_matrix_df <- as.data.frame(framing_mail_matrix)
framing_mail_matrix_df[is.na(framing_mail_matrix_df)] = 0

framing_mail_nodes <- framing_mail_matrix_df %>% 
  dplyr::mutate(From = X1,
                Occurrences = rowSums(framing_mail_matrix_df[-1])) %>%
  dplyr::select(From, Occurrences) %>%
  dplyr::filter(!str_detect(From, "X4"))

head(framing_mail_matrix_df)

framing_mail_edges <- framing_mail_matrix_df %>%
  dplyr::mutate(from = X1) %>%
  tidyr::gather(to, Frequency, "LMarkova@tv-culture.ru":"aredkina@vgtrk.com") %>%
  dplyr::mutate(Frequency = ifelse(Frequency == 0, NA, Frequency))
framing_mail_edges <- subset(framing_mail_edges, select = -c(X1))

framing_news_network <- igraph::graph_from_data_frame(
  d = framing_mail_edges, directed = TRUE)
tidy_framing_news_network <- tidygraph::as_tbl_graph(
  framing_news_network) %>%
  tidygraph::activate(edges) %>%
  dplyr::mutate(name=from)

E(tidy_framing_news_network)$weight <- E(
  tidy_framing_news_network)$Frequency

set.seed(12345)

tidy_framing_news_network %>%
  ggraph(layout = "fr") +
  geom_edge_arc(
    colour = "gray50", 
    lineend = "round", 
    strength = 0.1, 
    aes(edge_width = weight, 
        alpha = weight)) +
  geom_node_text(aes(label = name),
                 repel = TRUE,
                 point.padding = unit(0.2, "lines"),
                 colour = "gray10") +
  theme_graph(background = "white") +
  guides(edge_width = FALSE,
         edge_alpha = FALSE)

##Anonymized image:

tidy_framing_news_network %>%
  ggraph(layout = "fr") +
  geom_edge_arc(
    colour = "gray50", 
    lineend = "round", 
    strength = 0.1, 
    aes(edge_width = weight, 
        alpha = weight)) +
  theme_graph(background = "white") +
  guides(edge_width = FALSE,
         edge_alpha = FALSE)



