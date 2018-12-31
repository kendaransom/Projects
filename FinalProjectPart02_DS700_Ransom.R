# File for cleansing data for DS700 Final Project

#read in Excel Workbook
#install.packages("openxlsx")
library("openxlsx")

# read in sheet 2 - Abbeville, LA
Abbeville = read.xlsx("Dataset.xlsx", sheet = 2, startRow = 1, colNames = TRUE)
attach(Abbeville)
# what's in the data
summary(Abbeville)

# cleanse Incoming Examinations - remove non numeric characters
non_numeric_rows = which(is.na(as.numeric(as.character(Incoming.Examinations))))
non_numeric_rows
Incoming.Examinations[non_numeric_rows] = ''

# cleanse Incoming Examinations - remove outlier numeric data
bad_numeric_rows = which(as.numeric(as.character(Incoming.Examinations)) > 10000)
bad_numeric_rows
Incoming.Examinations[bad_numeric_rows] = ''

# create "cleaned" data frame for Abbeville worksheet 
abbeville.df = data.frame(Incoming.Examinations, Year, Month, stringsAsFactors = FALSE)

# sort "cleaned" data by year then month and plot

# create a function to:
# read in sheets 3-6 - 4 HC's
# find Original Hospital Location = Abbeville
# find heart related examinations using a list of heart related conditions
# remove duplicates
# return a data frame

cleanse = function(i,y){
  hc.data = read.xlsx("Dataset.xlsx", sheet = i, startRow = 1, colNames = TRUE)
  a.rows = which(hc.data$Original.Hospital.Location == "Abbeville")
  hc.data = hc.data[a.rows, ]
  heart.rows = which(hc.data$Examination %in% y)
  a.df = data.frame(hc.data[heart.rows, 1:4])
  rm_dups = unique(a.df, by = a.df[ ,4])
  a.df = rm_dups[ ,1:3]
  return(assign(paste("heart.data",i,sep='_'),a.df, envir = .GlobalEnv))
}


# y = heart related condtions
y = c("Myocarditis", "Cardiac","Aortic Valve Stenosis","Cor Pulmonale", "Angina",
      "Ischemic Heart Disease", "Myocardial Infraction", "Myocardial Ischemia",
      "Ventricular Septal Defect (VSD)", "Premature Ventricular Contraction", "coronary Artery Disease (CAD)",
      "CAD", "Arrhythmia", "Cardiovascular", "VSD", "Heart", "Heart Palpitations","Endocarditis", "Cur Pulmonale")

# loop to call the function to find the heart examinations
for (i in 3:6){
  cleanse(i, y)
}

# Count of Violet, LA heart related exams (all rows are May 2007)
heart.0507.exams_3 = dim(heart.data_3)[1]

# Count of New Orleans, LA heart related exams (all rows are May 2013)
heart.0513.exams_4 = dim(heart.data_4)[1]

# Count of Lafayette, LA heart related exams (sort by dates)
# 6 rows out of the Lafayette (5) and Baton Rouge (6) HC are formatted
# different, but have the same dates
dates = c('12 May, 2007','13 May, 2007','14 May, 2007','15 May, 2007','17 May, 2007','23 May, 2007')
dc.rows_5 = which(heart.data_5$Date %in% dates)
heart.exams_5 = length(dc.rows_5)

# Find remaining May 2007 dates - dates based on origin since 1899-12-30
rows_5 = which(as.numeric(as.character(heart.data_5[-dc.rows_5,3])) < 39500)
heart.0507.exams_5 = heart.exams_5 + length(rows_5)

# Find May 2013 dates - dates based on origin since 1899-12-30
rows_5 = which(as.numeric(as.character(heart.data_5[-dc.rows_5,3])) > 41394 &
                 as.numeric(as.character(heart.data_5[-dc.rows_5,3])) < 41426)
heart.0513.exams_5 = length(rows_5)

# Find June 2013 dates - dates based on origin since 1899-12-30
rows_5 = which(as.numeric(as.character(heart.data_5[-dc.rows_5,3])) > 41425 &
                 as.numeric(as.character(heart.data_5[-dc.rows_5,3])) < 41456)
heart.0613.exams_5 = length(rows_5)

# Find July 2013 dates - dates based on origin since 1899-12-30
rows_5 = which(as.numeric(as.character(heart.data_5[-dc.rows_5,3])) > 41455 &
                 as.numeric(as.character(heart.data_5[-dc.rows_5,3])) < 41487)
heart.0713.exams_5 = length(rows_5)

# Count of Baton Rouge, LA heart related exams (sort by dates)
# 6 rows out of the Lafayette (5) and Baton Rouge (6) HC are formatted
# different, but have the same dates

dc.rows_6 = which(heart.data_6$Date %in% dates)
heart.exams_6 = length(dc.rows_6)

# Find remaining May 2007 dates - dates based on origin since 1899-12-30
rows_6 = which(as.numeric(as.character(heart.data_6[-dc.rows_6,3])) < 39500)
heart.0507.exams_6 = heart.exams_6 + length(rows_6)

# Find May 2013 dates - dates based on origin since 1899-12-30
rows_6 = which(as.numeric(as.character(heart.data_6[-dc.rows_6,3])) > 41394 &
                 as.numeric(as.character(heart.data_6[-dc.rows_6,3])) < 41426)
heart.0513.exams_6 = length(rows_6)

# Find June 2013 dates - dates based on origin since 1899-12-30
rows_6 = which(as.numeric(as.character(heart.data_6[-dc.rows_6,3])) > 41425 &
                 as.numeric(as.character(heart.data_6[-dc.rows_6,3])) < 41456)
heart.0613.exams_6 = length(rows_6)

# Find July 2013 dates - dates based on origin since 1899-12-30
rows_6 = which(as.numeric(as.character(heart.data_6[-dc.rows_6,3])) > 41455 &
                 as.numeric(as.character(heart.data_6[-dc.rows_6,3])) < 41487)
heart.0713.exams_6 = length(rows_6)

# Parse the December 2013 data
hc.data = read.xlsx("Dataset.xlsx", sheet = 7, startRow = 1, colNames = TRUE)

# first 7998 rows are Abbeville data based on the start and end codes
hc.a.data = hc.data[1:7998, ]

# heart conditions codes
hc.codes = c('VVN284', 'PJU008', 'ABN441', 'UMP621', 'UMX710', 'TUX333', 'KPN015',
             'RLX001', 'WPC608', 'LLA092', 'LLN112', 'TNR628', 'LOR159', 'KON421',
             'KOZ198', 'ROB001')

# Abbeville heart conditions for December 2013
rows_7 = vector()
for (i in 1:length(hc.codes)){
  rows_7[i] = length(grep(hc.codes[i], hc.a.data))
}
heart.1213.exams = sum(rows_7) 

# All Excel sheets handled, add in new data into appropriate abbeville.df rows
# can only add data for 5/2007, 5/2013, 6/2013, 7/2013, and 12/2013 from Excel sheets

# Impute 5/2007 missing data - Violet, LA, Lafayette, LA, and Baton Rouge, LA
may07.row = as.numeric(as.character(which(abbeville.df[ ,2] == '2007' & abbeville.df[ ,3]  == '5')))
may07.data = heart.0507.exams_3 + heart.0507.exams_5 + heart.0507.exams_6
may07 = as.numeric(as.character(abbeville.df$Incoming.Examinations[may07.row])) + as.numeric(may07.data)
abbeville.df$Incoming.Examinations[may07.row] = may07

# Impute 5/2013 missing data - New Orleans, LA, Lafayette, LA, and Baton Rouge, LA
may13.row = as.numeric(as.character(which(abbeville.df[ ,2] == '2013' & abbeville.df[ ,3]  == '5')))
may13.data = heart.0513.exams_4 + heart.0513.exams_5 + heart.0513.exams_6
may13 = as.numeric(as.character(abbeville.df$Incoming.Examinations[may13.row])) + as.numeric(may13.data)
abbeville.df$Incoming.Examinations[may13.row] = may13

# Impute 6/2013 missing data - Lafayette, LA, and Baton Rouge, LA
june13.row = as.numeric(as.character(which(abbeville.df[ ,2] == '2013' & abbeville.df[ ,3]  == '6')))
june13.data = heart.0613.exams_5 + heart.0613.exams_6
june13 = as.numeric(as.character(abbeville.df$Incoming.Examinations[june13.row])) + as.numeric(june13.data)
abbeville.df$Incoming.Examinations[june13.row] = june13

# Impute 7/2013 missing data - Lafayette, LA, and Baton Rouge, LA
july13.row = as.numeric(as.character(which(abbeville.df[ ,2] == '2013' & abbeville.df[ ,3]  == '7')))
july13.data = heart.0713.exams_5 + heart.0713.exams_6
july13 = as.numeric(as.character(abbeville.df$Incoming.Examinations[july13.row])) + as.numeric(july13.data)
abbeville.df$Incoming.Examinations[july13.row] = july13

# Impute 12/2013 missing data - Lafayette, LA, and Baton Rouge, LA
dec13.row = as.numeric(as.character(which(abbeville.df[ ,2] == '2013' & abbeville.df[ ,3]  == '12')))
abbeville.df$Incoming.Examinations[dec13.row] = as.numeric(heart.1213.exams)

# Fill missing data with NA and convert exams to numeric
abbeville.df$Incoming.Examinations = as.numeric(abbeville.df$Incoming.Examinations)
data.frame(apply(abbeville.df, 2, is.na))

#plot of Excel cleaned data
missing_data =order(abbeville.df$Year, abbeville.df$Month)
plot(abbeville.df$Incoming.Examinations[missing_data], ylab = "Heart Exams", xlab = "Years", main = "Abbeville Heart Exams Post Excel Clean", type = 'o')

# Data Imputation - Year by Year
library(mice)
library(VIM)
md.pattern(abbeville.df)

# 2006 Data Imputation - MICE with method = norm for linear regression
hc2006.rows = which(abbeville.df[ ,2] == '2006')
hc2006.data = abbeville.df[hc2006.rows, ]
aggr_plot = aggr(hc2006.data[ ,c(1,3)], col=c('navyblue','red'), numbers=TRUE, sortVars=TRUE,
                 cex.axis=.7, gap=3, ylab=c("Histogram of missing data","Pattern"), main = c("2006 Data"))

temp2006.data = mice(hc2006.data, m =5, method = 'norm', maxit = 5, seed=500)
temp2006.data$imp$Incoming.Examinations

# Use the median of the 5 imputations to get the missing values for March and June
med = apply(temp2006.data$imp$Incoming.Examinations, 1, median)

# Impute March data
hc2006.na.rows = which(is.na(hc2006.data[ ,1]))
hc2006.data[hc2006.na.rows[1],1] = round(med[1],0)

# Impute June data
hc2006.data[hc2006.na.rows[2],1] = round(med[2],0)
# Order 2006 data
hc2006.data = hc2006.data[order(hc2006.data$Year, hc2006.data$Month), ]
#plot
plot(hc2006.data[,3],hc2006.data[,1], ylab = "Heart Exams", xlab = "Months", main = "2006 Heart Exams")

# No missing 2007 data, but order
hc2007.rows = which(abbeville.df[ ,2] == '2007')
hc2007.data = abbeville.df[hc2007.rows, ]
hc2007.data = hc2007.data[order(hc2007.data$Year, hc2007.data$Month), ]
#plot
plot(hc2007.data[,3],hc2007.data[,1], ylab = "Heart Exams", xlab = "Months", main = "2007 Heart Exams")


# 2008 Data Imputation - MICE with sample due to October 2008 data
hc2008.rows = which(abbeville.df[ ,2] == '2008')
hc2008.data = abbeville.df[hc2008.rows, ]
aggr_plot = aggr(hc2008.data[ ,c(1,3)], col=c('navyblue','red'), numbers=TRUE, sortVars=TRUE,
                 cex.axis=.7, gap=3, ylab=c("Histogram of missing data","Pattern"))

temp2008.data = mice(hc2008.data, m =5, method = 'sample', maxit = 5, seed=500)


# Randomly select an imputate that seems plausible for the data
temp2008.data$imp$Incoming.Examinations

# Impute December data - use complete becasue there is only one value to fill
hc2008.data = complete(temp2008.data,3)

# Order 2008 data
hc2008.data = hc2008.data[order(hc2008.data$Year, hc2008.data$Month), ]
#plot
plot(hc2008.data[,3],hc2008.data[,1], ylab = "Heart Exams", xlab = "Months", main = "2008 Heart Exams")


# 2009 Data Imputation - MICE with norm
hc2009.rows = which(abbeville.df[ ,2] == '2009')
hc2009.data = abbeville.df[hc2009.rows, ]

aggr_plot = aggr(hc2009.data[ ,c(1,3)], col=c('navyblue','red'), numbers=TRUE, sortVars=TRUE,
                 cex.axis=.7, gap=3, ylab=c("Histogram of missing data","Pattern"))

temp2009.data = mice(hc2009.data, m =5, method = 'norm', maxit = 5, seed=500)

# Imputed Values 
temp2009.data$imp$Incoming.Examinations
med = apply(temp2009.data$imp$Incoming.Examinations, 1, median)

# Impute May data - use median for May 2009
hc2009.na.rows = which(is.na(hc2009.data[ ,1]))
hc2009.data[hc2009.na.rows[1],1] = round(med[1],0)

# Order 2009 data
hc2009.data = hc2009.data[order(hc2009.data$Year, hc2009.data$Month), ]
#plot
plot(hc2009.data[,3],hc2009.data[,1], ylab = "Heart Exams", xlab = "Months", main = "2009 Heart Exams Excluding December")


# 2010 Data Imputation - MICE with norm
hc2010.rows = which(abbeville.df[ ,2] == '2010')
hc2010.data = abbeville.df[hc2010.rows, ]

aggr_plot = aggr(hc2010.data[ ,c(1,3)], col=c('navyblue','red'), numbers=TRUE, sortVars=TRUE,
                 cex.axis=.7, gap=3, ylab=c("Histogram of missing data","Pattern"))

temp2010.data = mice(hc2010.data, m =5, method = 'norm', maxit = 5, seed=500)

# Imputed Values
temp2010.data$imp$Incoming.Examinations

# Impute June data - use median of 5 imputations
med = apply(temp2010.data$imp$Incoming.Examinations, 1, median)
hc2010.na.rows = which(is.na(hc2010.data[ ,1]))
hc2010.data[hc2010.na.rows[1],1] = round(med[1],0)

# Order 2010 data
hc2010.data = hc2010.data[order(hc2010.data$Year, hc2010.data$Month), ]
#plot
plot(hc2010.data[,3],hc2010.data[,1], ylab = "Heart Exams", xlab = "Months", main = "2010 Heart Exams Excl. January & February")


# Handling Dec 2009-Feb 2010 data

# get total exams for 2007, 2008, and 2009 (minus 12/2009)
# solve for yearly standard deviation
# get Dec, Jan, and Feb percentages of exams for 12/2007-2/2008, 12/2008-2/2009
# get % average for Dec, Jan, and Feb month
# get sd average average exams for 2007, 2008, 2009

# total exams in 2007
hc2007.rows = which(abbeville.df[ ,2] == '2007')
hc2007.data = abbeville.df[hc2007.rows, ]
total07.exams = sum(hc2007.data[ ,1])
sd07 = sd(hc2007.data[ ,1])

# total exams in 2008
total08.exams = sum(hc2008.data[ ,1])
sd08 = sd(hc2008.data[ ,1])

# total exams in 2009
total09.exams = sum(hc2009.data[ ,1], na.rm = TRUE)
sd09 = sd(hc2009.data[ ,1], na.rm = TRUE)


# 12/2007 percentage
dec.row = which(hc2007.data[ ,3] == 12)
dec07.pct = hc2007.data[dec.row,1]/total07.exams
# 1/2008 percentage
jan08.pct = hc2008.data[1, 1]/total08.exams
# 2/2008 percentage
feb08.pct = hc2008.data[2, 1]/total08.exams

# 12/2008 percentage
dec08.pct = hc2008.data[12, 1]/total08.exams
# 1/2009 percentage
jan09.pct = hc2009.data[1,1]/total09.exams
# 2/2009 percentage
feb09.pct = hc2009.data[2, 1]/total09.exams

# Dec avg
dec.avg = mean(c(dec07.pct,dec08.pct))
# Jan avg
jan.avg = mean(c(jan08.pct,jan09.pct))
# Feb avg
feb.avg = mean(c(feb08.pct,feb09.pct))
# SD avg
sd.avg = mean(c(sd08, sd09))
# exam avg
exams.avg = mean(c(total08.exams, total09.exams))

# Dec 2009
dec09 = (dec.avg*exams.avg) + sd.avg

# Jan 2010
jan10 = (jan.avg*exams.avg) + sd.avg

# Feb 2010
feb10 = (feb.avg*exams.avg) + sd.avg

# Add December 2009 data to 2009 dataset
hc2009.data[12, 1] = round(dec09,0)

# Add January and February 2010 data to 2010 dataset 
hc2010.jan.row = which(hc2010.data[ ,3] == 1) 
hc2010.feb.row = which(hc2010.data[ ,3] == 2) 
hc2010.data[hc2010.jan.row,1] = round(jan10,0)
hc2010.data[hc2010.feb.row,1] = round(feb10,0)


# 2011 Data Imputation - MICE with norm
hc2011.rows = which(abbeville.df[ ,2] == '2011')
hc2011.data = abbeville.df[hc2011.rows, ]

aggr_plot = aggr(hc2011.data[ ,c(1,3)], col=c('navyblue','red'), numbers=TRUE, sortVars=TRUE,
                 cex.axis=.7, gap=3, ylab=c("Histogram of missing data","Pattern"))

temp2011.data = mice(hc2011.data, m =5, method = 'norm', maxit = 5, seed=500)

# Imputed Values - missing Jan and Dec 2011
temp2011.data$imp$Incoming.Examinations

# Impute - use average of 5 imputations
med = apply(temp2011.data$imp$Incoming.Examinations, 1, median)
hc2011.na.rows = which(is.na(hc2011.data[ ,1]))

# Add in imputed Dec 2011 data
hc2011.data[hc2011.na.rows[1],1] = round(med[1],0)

# Add in imputed Jan 2011 data
hc2011.data[hc2011.na.rows[2],1] = round(med[2],0)

# Order 2011 data
hc2011.data = hc2011.data[order(hc2011.data$Year, hc2011.data$Month), ]
#plot
plot(hc2011.data[,3],hc2011.data[,1], ylab = "Heart Exams", xlab = "Months", main = "2011 Heart Exams")


# Order 2012 and 2013 data
hc2012.rows = which(abbeville.df[ ,2] == '2012')
hc2012.data = abbeville.df[hc2012.rows, ]
hc2012.data = hc2012.data[order(hc2012.data$Year, hc2012.data$Month), ]
plot(hc2012.data[,3],hc2012.data[,1], ylab = "Heart Exams", xlab = "Months", main = "2012 Heart Exams")


hc2013.rows = which(abbeville.df[ ,2] == '2013')
hc2013.data = abbeville.df[hc2013.rows, ]
hc2013.data = hc2013.data[order(hc2013.data$Year, hc2013.data$Month), ]
plot(hc2013.data[,3],hc2013.data[,1], ylab = "Heart Exams", xlab = "Months", main = "2013 Heart Exams")


# Create final data frame of cleansed data to forcast with
hc.data = data.frame(rbind(hc2006.data,hc2007.data,hc2008.data,hc2009.data,
                           hc2010.data,hc2011.data,hc2012.data,hc2013.data))

# Write to .csv file
write.csv(hc.data, "Abbeville_Cleansed.csv")

# Forecasting Code for Abbeville Heart Condition Data - DS700 Final Project

hc.data = read.csv("Abbeville_Cleansed.csv", header = TRUE)
hc.data = hc.data[ ,-1]
attach(hc.data)


#plot dat
plot(hc.data[ ,1], ylab = "Heart Exams", xlab = "Index", main ="Scatterplot of Heart Exams")

require(forecast)

# time series
hc.ts = ts(hc.data[ ,1], start=2006, frequency=12)
plot(hc.ts, ylab = "Heart Exams", xlab = "Years", main = "Time Series Plot of Heart Exams")

# Holt-Winters Expotential Smoothing with trending
hc.hw = HoltWinters(hc.ts, gamma = FALSE)

# Holt-Winters Forecasting
hc.hw.forecast = forecast(hc.hw, h=12)
plot(hc.hw)
plot(hc.hw.forecast, ylab = "Heart Exams", xlab = "Years")
summary(hc.hw.forecast)


# ARIMA forecast

# Correlation plots - not stationary - confirms trend
acf(hc.ts) # autoregression present, large spike at initial that decays towards 0 is autoregressive
pacf(hc.ts) # autoregressive and moving average, large spike in 1st 2 variables that decays to 0 is MA

# Diffing to transform data to account for October 2008 spike in data
# fixes time series with trends or drift


# Optimal number for diffing
ndiffs(hc.ts)

# Plot with diffing
plot(diff(hc.ts,1))

# Fit ARIMA model - let R decide best model
hc.arima = auto.arima(hc.ts)
hc.arima
#fitted model
fit = hc.arima$fitted
# plot of fit
plot(hc.ts, ylab = "Heart Exams", xlab = "Years", main = "Time Series Plot of Heart Exams")
lines(fit, col='red')

# Plot residuals in acf and pacf - look vs bounds
acf(hc.arima$residuals)
pacf(hc.arima$residuals)

# ARIMA Forecast
hc.arima.forecast = forecast(hc.arima, h=12)
plot(hc.arima.forecast,ylab = "Heart Exams", xlab = "Years")
summary(hc.arima.forecast)


