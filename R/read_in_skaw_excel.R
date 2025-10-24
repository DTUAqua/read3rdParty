#' Title
#'
#' @param file 
#'
#' @returns
#' @export
#'
#' @examples
read_in_skaw_excel <- function(input_path, file, output_path) {
  
  library(readxl)
  library(stringr)
  library(dplyr)
  
  # Create folders for outputting
  
  dir.create(file.path(output_path, "to_many_sheets"))
  dir.create(file.path(output_path, "read"))
  dir.create(file.path(output_path, "sworn_cant_read_yet"))
  dir.create(file.path(output_path, "skaw_dtu_19_cant_read_yet"))
  
  print(file)
  
  sheet_names <- excel_sheets(path = paste0(input_path, file))
  
  if (length(sheet_names) > 1) {
    file.copy(
      from = paste0(input_path, file),
      to = file.path(output_path, "to_many_sheets"),
      overwrite = T
    )
    
    dat_done <- c()
    
  } else if (sheet_names[1] %in% c("Tobis",
                                   "Brisling 16+",
                                   "Sild",
                                   "Fjæsing",
                                   "Brisling",
                                   "Sild 16+")) {
    file.copy(
      from = paste0(input_path, file),
      to = file.path(output_path, "sworn_cant_read_yet"),
      overwrite = T
    )
    
    dat_done <- c()
    
  } else if (sheet_names[1] %in% c("DTU_19")) {
    
    file.copy(
      from = paste0(input_path, file),
      to = file.path(output_path, "skaw_dtu_19_cant_read_yet"),
      overwrite = T
    )
    
    dat_done <- c()
    
  } else if (length(sheet_names) == 1 &
             !(
               sheet_names[1] %in% c(
                 "Tobis",
                 "Brisling 16+",
                 "Sild",
                 "Fjæsing",
                 "Brisling",
                 "Sild 16+",
                 "DTU_19"
               )
             )) {
    file.copy(
      from = paste0(input_path, file),
      to = file.path(output_path, "read"),
      overwrite = T
    )
    
    
    dat <-  read_excel(paste0(input_path, file),
                       sheet = 1,
                       col_names = FALSE)
    
  position_col <- grep("Prøveoversigt", x = dat)
  
  dat_1 <- dat[, -c(1:(position_col-1))]
  
  # Find univesal info ----
  
  position_head_1 <- which(apply(dat_1, 1, function(x) any(grepl("Skib:", x))))
  head_1 <- dat_1[position_head_1, ]
  
  fid <- head_1[, grep("Skib:", head_1) + 1]
  fid_no <- str_extract(fid, "[[:digit:]]{2,3}")
  fid_char <- gsub("[[:digit:]].*","", fid)
  fid_char_1 <- gsub(" ","", fid_char)
  fid <- paste0(fid_char_1, fid_no)
  fid <- as.data.frame(fid)
  
  buyer <- head_1[, grep("Køber:", head_1) + 2]
  colnames(buyer) <- c("buyer")
  if (is.na(buyer$buyer)) {
    buyer <- head_1[, grep("Køber:", head_1) + 1]
    colnames(buyer) <- c("buyer")
  }
  
  position_head_2 <- which(apply(dat_1, 1, function(x) any(grepl("Total vejet fisk:", x))))
  head_2 <- dat_1[position_head_2, ]
  
  tot_wgt <- head_2[, grep("Total vejet fisk:", head_2) + 2]
  colnames(tot_wgt) <- c("tot_wgt")
  tot_wgt$tot_wgt <- as.numeric(tot_wgt$tot_wgt)
  
  date <- head_2[, grep("Dato:", head_2) + 2]
  colnames(date) <- c("date")
  date$date <- as.Date(date$date, format = "%d.%m.%Y")
  if (is.na(date$date)) {
    date <- head_2[, grep("Dato:", head_2) + 1]
    colnames(date) <- c("date")
    date$date <- as.Date(date$date, format = "%d.%m.%Y")
  }
  
  head_2 <- dat_1[position_head_2, ]
  
  head_done <- cbind(fid, buyer, tot_wgt, date)
  
  
  # Find samples ----
  
  position_samples_head <- which(apply(dat_1, 1, function(x) any(grepl("Fiskeart", x))))
  samples_head <- dat_1[position_samples_head, ]
  samples_head <- as.data.frame(as.character(samples_head))
  samples_head$`as.character(samples_head)`[samples_head$`as.character(samples_head)`== "Gobler\r\n(Vandmæ-nd)AJQ"] <- "AJQ"
  samples_head[is.na(samples_head) | samples_head == "NA"] <- "X"
  
  position_samples <- which(apply(dat_1, 1, function(x) any(grepl("Prøve nr.", x))))
  
  samples <- dat_1[c(position_samples:nrow(dat_1)), ]
  samples_1 <-  samples   #subset(samples, !is.na(...9))
  
  colnames(samples_1) <- samples_head[1:nrow(samples_head), ]
  
  # Remove columns after Select number of rows:
  position_sel_row <- which(colnames(samples_1)=="Inspektor(s)")
  if (length(position_sel_row) == 1) {
    samples_1 <- samples_1[, -c(position_sel_row:ncol(samples_1))]
  } else {
    position_sel_row_2 <- which(colnames(samples_1)=="Total bi-\r\nfangst (kg)")
    samples_1 <- samples_1[, -c(position_sel_row_2:ncol(samples_1))]
  }
  
  
  col_names <- names(samples_1)
  col_names <- gsub(".*\\s", "", col_names)
  
  for (i in seq_along(col_names)) {
    
    if (substr(col_names[i], 1, 1) == "X") {
      if (is.na(samples_1[1, i])) {
        col_names[i] <- col_names[i]
      } else if (samples_1[1, i] == "(kg)") {
        col_names[i] <- col_names[i-1]
      } else if (samples_1[1, i] == "stk.")
        col_names[i] <- col_names[i-1]
    } }
  
  for (i in seq_along(col_names)) {   
    if (is.na(samples_1[1, i])) {
      col_names[i] <- col_names[i]
    } else if (samples_1[1, i] == "(kg)") {
      col_names[i] <- paste0(col_names[i], "_kg")
    } else if (samples_1[1, i] == "stk.") {
      col_names[i] <- paste0(col_names[i], "_stk")
    } else {
      col_names[i] <- col_names[i]
    }
  }
  
  colnames(samples_1) <- col_names
  
  # Clean up a bit
  samples_2 <- samples_1 %>% select(-contains("X"), -contains("(kg)"), -contains("Prøven"), -contains("snit"), -contains("Tid"))
  samples_2 <- subset(samples_2, !(`Kode:` %in% c("Prøve nr.", "Total:")))
  samples_2 <- rename(samples_2, sample_no = `Kode:`)
  samples_2$sample_no <- as.numeric(samples_2$sample_no)
  samples_2 <- subset(samples_2, !(is.na(sample_no)))
  
  # Transpose 
  samples_3 <- tidyr::gather(samples_2, key = spp_fao, value = amount, -sample_no)
  samples_3 <- subset(samples_3, !(amount %in% c(".", "-", "=", "", "´0,20")))
  samples_3$amount <- gsub("\\+", "", samples_3$amount) # this is not so nice
  samples_3$kg <- NA
  samples_3$kg[grepl("kg", samples_3$spp_fao) == T] <- 
    as.numeric(samples_3$amount[grepl("kg", samples_3$spp_fao) == T])
  samples_3$pcs <- NA
  samples_3$pcs[grepl("stk", samples_3$spp_fao) == T] <- 
    as.numeric(samples_3$amount[grepl("stk", samples_3$spp_fao) == T])
  samples_3$spp_fao <- substr(samples_3$spp_fao, 1, 3)
  
  kg <- subset(samples_3, !is.na(kg))
  kg <- select(kg, sample_no, spp_fao, kg)
  stk <- subset(samples_3, !is.na(pcs))
  stk <- select(stk, sample_no, spp_fao, pcs)
  
  samples_4 <- full_join(kg, stk)
  
  # Add global info
  
  samples_4$fishery <- substr(colnames(samples_2)[2], 1, 3)
  samples_4$filename <- paste0(input_path, file)
  
  samples_5 <- cross_join(samples_4, head_done)
  
  dat_done <- samples_5
  }
  return(dat_done)
  
}
