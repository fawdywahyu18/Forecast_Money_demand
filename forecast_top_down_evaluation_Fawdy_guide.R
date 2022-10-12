library(ggplot2)
library(forecast)
library(bsts)
library(repr)
library(readxl)
library(writexl)
library(rlist)
library(dplyr)

# Metode: Proyeksi rasio pemusnahan uang thdp inflow, konversi, lalu dibandingkan dengan proyeksi BI
# Creating the function
forecast_all <- function(x, nahead) {
  
  fstl_arima <- stlf(x, h=nahead, method="arima", level=95)
  fstl_rw <- stlf(x, h=nahead, method="rwdrift", level=95)
  fstl_ets <- stlf(x, h=nahead, method="ets", level=95)
  fhw <- forecast(hw(x), h = nahead)
  fets <- forecast(ets(x), h = nahead)
  
  forecast_result <- list(fstl_arima$mean, fstl_rw$mean, fstl_ets$mean,
                          fhw$mean, fets$mean)
  
  for (i in 1:5) {
    forecast_result[[i]] <- data.frame(forecast_result[[i]])
  }
  
  return(forecast_result)
}

# Forecast pemusnahan uang Top Down (Nasional)
mainwd = 'E:/Kerja\Asisten Akademik/Riset Bank Indonesia/DPU 2021/Transfer Knowledge Riset DPU 2021'
setwd(".") # Set your working directory here
data_pu <- data.frame(read_xlsx("pemusnahan_uang.xlsx", sheet="NAS_com"))$pemusnahan
data_inflow <- data.frame(read_xlsx("inflow_kpw.xlsx", sheet="NAS_com"))$inflow
monthly_pu <- ts(data_pu, start=c(2010, 1), end=c(2021, 9), frequency=12)
monthly_inflow <- ts(data_inflow, start=c(2010, 1), end=c(2021, 9), frequency=12)
monthly_ratio <- monthly_pu/monthly_inflow
# quarterly <- aggregate(monthly, nfrequency=4)
forecast_result_ratio <- forecast_all(monthly_ratio, 9)
forecast_result_inflow <- forecast_all(monthly_inflow, 9)
# forecast_result_inflow[forecast_result_inflow<0] <- 0

mat_indeks <- matrix(seq(1,25), ncol = 5, byrow = TRUE)
forecast_result_pu <- list()
for (i in 1:5) {
  for (o in 1:5) {
    convertion_result <- forecast_result_ratio[[i]] * forecast_result_inflow[[o]]
    indeks <- mat_indeks[i,o]
    forecast_result_pu[[indeks]] <- data.frame(convertion_result)
    colnames(forecast_result_pu[[indeks]]) <- "pemusnahan"
  }
}

raw_pemusnahan_nasional <- tail(data.frame(read_xlsx("pemusnahan_uang.xlsx", sheet="NAS")$pemusnahan), 3)
colnames(raw_pemusnahan_nasional) <- "pemusnahan"

annually_result <- list()
for (p in 1:25) {
  rbind_result <- rbind(raw_pemusnahan_nasional, forecast_result_pu[[p]])
  ts_object <- ts(rbind_result, start=c(2021, 1), end=c(2021, 12), frequency=12)
  annually_object <- aggregate(ts_object, nfrequency = 1)
  annually_result[[p]] <- data.frame(annually_object/1000000)
}

matriks <- matrix(c(271.2), ncol = 1)
proyeksi_bi <- data.frame(matriks)
colnames(proyeksi_bi) <- "pemusnahan"

diff_result <- list()
for (u in 1:25) {
  diff_result[[u]] <- annually_result[[u]]-proyeksi_bi
}
diff_result_matrix <- matrix(unlist(diff_result), ncol=5, byrow = FALSE)
diff_result_matrix


# Metode: Proyeksi langsung inflow dan pemusnahan uang secara top down, kemudian dibandingkan dengan proyeksi BI
data_pu <- data.frame(read_xlsx("pemusnahan_uang.xlsx", sheet="NAS"))$pemusnahan
data_inflow <- data.frame(read_xlsx("inflow_kpw.xlsx", sheet="NAS"))$inflow
monthly_pu <- ts(data_pu, start=c(2010, 1), end=c(2021, 3), frequency=12)
monthly_inflow <- ts(data_inflow, start=c(2010, 1), end=c(2021, 3), frequency=12)
forecast_direct_pu <- forecast_all(monthly_pu, 9)
forecast_direct_inflow <- forecast_all(monthly_inflow, 9)

for (y in 1:5) {
  colnames(forecast_direct_pu[[y]]) <- "pemusnahan"
  colnames(forecast_direct_inflow[[y]]) <- "inflow"
}

annually_result2 <- list()
for (p in 1:5) {
  rbind_result <- rbind(raw_pemusnahan_nasional, forecast_direct_pu[[p]])
  ts_object <- ts(rbind_result, start=c(2021, 1), end=c(2021, 12), frequency=12)
  annually_object <- aggregate(ts_object, nfrequency = 1)
  annually_result2[[p]] <- data.frame(annually_object/1000000)
}

diff_result2 <- list()
for (u in 1:5) {
  diff_result2[[u]] <- annually_result2[[u]]-proyeksi_bi
}
diff_result2_matrix <- matrix(unlist(diff_result2), ncol=1, byrow = FALSE)
diff_result2_matrix

raw_inflow_nasional <- tail(data.frame(read_xlsx("inflow_kpw.xlsx", sheet="NAS")$inflow), 3)
colnames(raw_inflow_nasional) <- "inflow"
annually_result3 <- list()
for (p in 1:5) {
  rbind_result <- rbind(raw_inflow_nasional, forecast_direct_inflow[[p]])
  ts_object <- ts(rbind_result, start=c(2021, 1), end=c(2021, 12), frequency=12)
  annually_object <- aggregate(ts_object, nfrequency = 1)
  annually_result3[[p]] <- data.frame(annually_object/1000000)
}
matriks <- matrix(c(830.5), ncol = 1)
proyeksi_inflow_bi <- data.frame(matriks)
colnames(proyeksi_inflow_bi) <- "inflow"

diff_result3 <- list()
for (u in 1:5) {
  diff_result3[[u]] <- annually_result3[[u]]-proyeksi_inflow_bi
}
diff_result3_matrix <- matrix(unlist(diff_result3), ncol=1, byrow = FALSE)
diff_result3_matrix

# Running the bottom up projection using the forecast_bottom_up_Fawdy_guide.R 
# for inflow and pemusnahan uang

# Evaluasi MASE dan MAPE
evaluasi_sumflow <- function(data.test, y_test, h, metode){
  sumflow <- data.frame(data.test[, c(y_test)])
  sumflow.ts <- ts(sumflow, start = c(2010, 1), end = c(2021, 3), frequency = 12)
  x <- sumflow.ts
  
  # create training data
  train2 <- window(x, end = c(2020, 12))
  
  # create specific test data of interest
  test <- window(x, start = c(2021, 1), end = c(2021, 3))
  
  # Forecast
  stlf.arima <- stlf(train2, method="arima", h=h)
  stlf.ets <- stlf(train2, method="ets", h=h)
  stlf.rwdrift <- stlf(train2, method="rwdrift", h=h)
  fhw <- forecast(hw(train2), h=h)
  fets <- forecast(ets(train2), h = h)
  f.list <- list(stlf.arima, stlf.ets, stlf.rwdrift,
                 fhw, fets)
  
  # Accuracy 2
  rata2_mape <- list()
  for (i in f.list) {
    mean1_mape <- accuracy(i, x)["Test set", "MAPE"]
    rata2_mape <- list.append(rata2_mape, mean1_mape)
  }
  rata2_mape.df <- data.frame(rata2_mape)
  colnames(rata2_mape.df) <- c("stlf arima", "stlf ets", "stlf rwdrift",
                               "holt-winter", "ets")
  mape <- rata2_mape.df
  
  # Accuracy 3
  rata2_mase <- list()
  for (i in f.list) {
    mean1_mase <- accuracy(i, x)["Test set", "MASE"]
    rata2_mase <- list.append(rata2_mase, mean1_mase)
  }
  rata2_mase.df <- data.frame(rata2_mase)
  colnames(rata2_mase.df) <- c("stlf arima", "stlf ets", "stlf rwdrift",
                               "holt-winter", "ets")
  mase <- rata2_mase.df
  
  if (metode=="mape") {
    output <- mape
  } else if (metode=="mase") {
    output <- mase
  } else {
    output <- "Error dalam memasukkan nama metode forecast"
  }
  
  return(output)
}

data.test <- data.frame(monthly_ratio)
colnames(data.test) <- c("Rasio PU Nasional")
colomn.names <- colnames(data.test)

# Evaluation MAPE dan MASE
eval.list.mape <- list()
for (i in colomn.names) {
  eval <- evaluasi_sumflow(data.test, i, 3, "mape")
  eval.list.mape[[i]] <- data.frame(eval)
}

eval.list.mase <- list()
for (i in colomn.names) {
  eval <- evaluasi_sumflow(data.test, i, 3, "mase")
  eval.list.mase[[i]] <- data.frame(eval)
}

mape_pu_nasional <- eval.list.mape[["Rasio PU Nasional"]]
mase_pu_nasional <- eval.list.mase[["Rasio PU Nasional"]]
evaluasi_pu <- rbind(mape_pu_nasional, mase_pu_nasional)
rownames(evaluasi_pu) <- c("MAPE", "MASE")
setwd(".") # set your working directory to export file here
write_xlsx(evaluasi_pu, "Evaluasi PU Top Down.xlsx")
setwd(".") # set your working directory

# Melakukan proyeksi outflow dan inflow secara langsung untuk mendapatkan netflow, 
# kemudian dibandingkan dengan proyeksi milik BI

# Creating modified forecast function
forecast_all_mod <- function(x, nahead) {
  
  fstl_arima <- stlf(x, h=nahead, method="arima", level=95)
  fstl_rw <- stlf(x, h=nahead, method="rwdrift", level=95)
  fstl_ets <- stlf(x, h=nahead, method="ets", level=95)
  fhw <- hw(x, h=nahead)
  fets <- forecast(ets(x), h = nahead)
  
  forecast_mean <- list(fstl_arima$mean, fstl_rw$mean, fstl_ets$mean,
                        fhw$mean, fets$mean)
  forecast_upper <- list(fstl_arima$upper, fstl_rw$upper, fstl_ets$upper,
                         fhw$upper[,"95%"], fets$upper[,"95%"])
  forecast_lower <- list(fstl_arima$lower, fstl_rw$lower, fstl_ets$lower,
                         fhw$lower[,"95%"], fets$lower[,"95%"])
  
  for (i in 1:5) {
    forecast_mean[[i]] <- data.frame(forecast_mean[[i]])
    forecast_upper[[i]] <- data.frame(forecast_upper[[i]])
    forecast_lower[[i]] <- data.frame(forecast_lower[[i]])
  }
  
  forecast_result <- list(forecast_mean, forecast_upper, forecast_lower)
  
  return(forecast_result)
}

# Melakukan proyeksi outflow dan inflow secara langsung menggunakan fungsi proyeksi yang sudah diubah
setwd(".") # set your working directory
data_outflow <- data.frame(read_xlsx("outflow_kpw.xlsx", sheet="NAS"))$outflow
data_inflow <- data.frame(read_xlsx("inflow_kpw.xlsx", sheet="NAS"))$inflow
data_netflow <- data_outflow - data_inflow
monthly_outflow <- ts(data_outflow, start=c(2010, 1), end=c(2021, 3), frequency=12)
monthly_inflow <- ts(data_inflow, start=c(2010, 1), end=c(2021, 3), frequency=12)
monthly_netflow <- monthly_outflow-monthly_inflow
forecast_result_outflow <- forecast_all_mod(monthly_outflow, 33)
forecast_result_inflow <- forecast_all_mod(monthly_inflow, 33)
forecast_result_netflow <- forecast_all_mod(monthly_netflow, 33)

# Keluarkan mean, upper, dan lower dari hasil proyeksi
# [[1]] adalah mean, [[2]] adalah upper, dan [[3]] adalah lower

raw_outflow <- data.frame(tail(data_outflow, 3))
raw_inflow <- data.frame(tail(data_inflow, 3))
raw_netflow <- data.frame(tail(data_netflow, 3))
colnames(raw_outflow) <- c("outflow")
colnames(raw_inflow) <- c("inflow")
colnames(raw_netflow) <- c("netflow")


# Renaming colomns name
for (i in 1:3) {
  for (u in 1:5) {
    colnames(forecast_result_inflow[[i]][[u]]) <- c("inflow")
    colnames(forecast_result_outflow[[i]][[u]]) <- c("outflow")
    colnames(forecast_result_netflow[[i]][[u]]) <- c("netflow")
  }
}

annually_result_outflow <- list(list(), list(), list())
annually_result_inflow <- list(list(), list(), list())
annually_result_netflow <- list(list(), list(), list())

# Rbind dan konversi bulanan ke tahunan, dan dalam triliun rupiah
for (i in 1:3) {
  for (u in 1:5) {
    rbind_result_inflow <- rbind(raw_inflow, forecast_result_inflow[[i]][[u]])
    ts_object_inflow <- ts(rbind_result_inflow, start=c(2021, 1), end=c(2023, 12), frequency=12)
    annually_object_inflow <- aggregate(ts_object_inflow, nfrequency = 1)
    annually_result_inflow[[i]][[u]] <- data.frame(annually_object_inflow/1000000)
    
    rbind_result_outflow <- rbind(raw_outflow, forecast_result_outflow[[i]][[u]])
    ts_object_outflow <- ts(rbind_result_outflow, start=c(2021, 1), end=c(2023, 12), frequency=12)
    annually_object_outflow <- aggregate(ts_object_outflow, nfrequency = 1)
    annually_result_outflow[[i]][[u]] <- data.frame(annually_object_outflow/1000000)
    
    rbind_result_netflow <- rbind(raw_netflow, forecast_result_netflow[[i]][[u]])
    ts_object_netflow <- ts(rbind_result_netflow, start=c(2021, 1), end=c(2023, 12), frequency=12)
    annually_object_netflow <- aggregate(ts_object_netflow, nfrequency = 1)
    annually_result_netflow[[i]][[u]] <- data.frame(annually_object_netflow/1000000)
  }
}

annually_outflow_mean <- data.frame(bind_rows(annually_result_outflow[[1]]))
annually_outflow_upper <- data.frame(bind_rows(annually_result_outflow[[2]]))
annually_outflow_lower <- data.frame(bind_rows(annually_result_outflow[[3]]))

annually_inflow_mean <- data.frame(bind_rows(annually_result_inflow[[1]]))
annually_inflow_upper <- data.frame(bind_rows(annually_result_inflow[[2]]))
annually_inflow_lower <- data.frame(bind_rows(annually_result_inflow[[3]]))

annually_netflow_mean <- data.frame(bind_rows(annually_result_netflow[[1]]))
annually_netflow_upper <- data.frame(bind_rows(annually_result_netflow[[2]]))
annually_netflow_lower <- data.frame(bind_rows(annually_result_netflow[[3]]))

outflow_annual <- cbind(annually_outflow_lower, annually_outflow_mean, annually_outflow_upper)
inflow_annual <- cbind(annually_inflow_lower, annually_inflow_mean, annually_inflow_upper)
netflow_annual <- cbind(annually_netflow_lower, annually_netflow_mean, annually_netflow_upper)

setwd(".") # set your working directory to export file here
write_xlsx(outflow_annual, "forecast outflow nasional.xlsx")
write_xlsx(inflow_annual, "forecast inflow nasional.xlsx")
write_xlsx(netflow_annual, "forecast netflow nasional.xlsx")

# Proyeksi untuk setiap pecahan uang nasional, outflow dan inflow
forecast_all_mean <- function(x, nahead) {
  
  fstl_arima <- stlf(x, h=nahead, method="arima", level=95)
  fstl_rw <- stlf(x, h=nahead, method="rwdrift", level=95)
  fstl_ets <- stlf(x, h=nahead, method="ets", level=95)
  fhw <- hw(x, h=nahead)
  fets <- forecast(ets(x), h = nahead)
  
  forecast_mean <- list(fstl_arima$mean, fstl_rw$mean, fstl_ets$mean,
                        fhw$mean, fets$mean)
  
  for (i in 1:5) {
    forecast_mean[[i]] <- data.frame(forecast_mean[[i]])
  }
  
  return(forecast_mean)
}

# Ganti outflow atau inflow untuk melakukan spesifikasi proyeksi
setwd(".") # set your working directory here
data_outflow <- data.frame(read_xlsx("outflow_kpw.xlsx", sheet="NAS"))
data_inflow <- data.frame(read_xlsx("inflow_kpw.xlsx", sheet="NAS"))
data_netflow <- data_outflow - data_inflow
data.test <- data_netflow # Ganti di sini
colomn.names <- colnames(data.test)
colomn.names <- colomn.names[-1]

# Menghasilkan proyeksi untuk setiap pecahan uang
forecast_sumflow <- list()
forecast_annually_sumflow <- list()
for (i in colomn.names) {
  sumflow <- data.frame(data.test[, c(i)])
  sumflow.ts <- ts(sumflow, start = c(2010, 1), end = c(2021, 3), frequency = 12)
  forecast_result <- forecast_all_mean(sumflow.ts, 33)
  forecast_sumflow[[i]] <- data.frame(forecast_result)
  raw_sumflow <- data.frame(tail(data.test[, c(i)], 3))
  matriks_raw_sumflow <- data.frame(matrix(nrow=3, ncol=5))
  for (p in 1:5) {
    matriks_raw_sumflow[,p] <- raw_sumflow
    raw_sumflow_df <- data.frame(matriks_raw_sumflow)
  }
  colnames(raw_sumflow_df) <- c("stlf_arima", "stlf_rw", "stlf_ets", "hw", "ets")
  colnames(forecast_sumflow[[i]]) <- c("stlf_arima", "stlf_rw", "stlf_ets", "hw", "ets")
  rbind_result_sumflow <- rbind(raw_sumflow_df, forecast_sumflow[[i]])
  ts_object_sumflow <- ts(rbind_result_sumflow, start=c(2021, 1), end=c(2023, 12), frequency=12)
  annually_object_sumflow <- aggregate(ts_object_sumflow, nfrequency = 1)
  forecast_annually_sumflow[[i]] <- data.frame(annually_object_sumflow/1000000)
}

annually_sumflow_df <- data.frame(list.cbind(forecast_annually_sumflow))

matriks_sumflow<- list()
for (n in colomn.names) {
  matriks_sumflow[[n]] <- unlist(forecast_annually_sumflow[[n]])
}

annually_sumflow_df <- data.frame(list.cbind(matriks_sumflow))

setwd(".") # set your working directory to export file here
write_xlsx(annually_sumflow_df, "forecast netflow per bank notes nasional.xlsx") # Ganti outflow dan inflow di sini

# Evaluation MAPE and MASE of outflow and inflow using predefined function
data.test <- data.frame(data_netflow) # Ganti outflow atau inflow di sini
colomn.names <- colnames(data.test)
colomn.names <- colomn.names[-1]

# Evaluation MAPE dan MASE
eval.list.mape <- list()
for (i in colomn.names) {
  eval <- evaluasi_sumflow(data.test, i, 4, "mape")
  eval.list.mape[[i]] <- data.frame(eval)
}

eval.list.mase <- list()
for (i in colomn.names) {
  eval <- evaluasi_sumflow(data.test, i, 4, "mase")
  eval.list.mase[[i]] <- data.frame(eval)
}

mape_sumflow <- list()
mase_sumflow <- list()
for (n in colomn.names) {
  mape_sumflow[[n]] <- unlist(eval.list.mape[[n]])
  mase_sumflow[[n]] <- unlist(eval.list.mase[[n]])
}

eval_mape_df <- data.frame(list.cbind(mape_sumflow))
eval_mase_df <- data.frame(list.cbind(mase_sumflow))

setwd(".") # set your working directory to export file here
write_xlsx(eval_mape_df, "mape netflow per bank notes nasional.xlsx") # Ganti outflow dan inflow di sini
write_xlsx(eval_mase_df, "mase netflow per bank notes nasional.xlsx") # Ganti outflow dan inflow di sini
