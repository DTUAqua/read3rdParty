#' Title
#'
#' @param file 
#'
#' @returns
#' @export
#'
#' @examples
read_in_dan_eye_excel <- function(input_path, id, output_path) {
  
  library(readxl)
  library(stringr)
  library(dplyr)
  
  # Create folders for outputting
  
  dir.create(file.path(output_path, "to_many_sheets"))
  dir.create(file.path(output_path, "read"))
  dir.create(file.path(output_path, "sworn_cant_read_yet"))
  dir.create(file.path(output_path, "skaw_dtu_19_cant_read_yet"))
  
  print(id)
  
  # Find relevant files
  file_list <- list.files(path = input_path, pattern = id, recursive = T)
  
  bifang_file <- subset(file_list, grepl(pattern = "Bifang", x = file_list))
  inspek_file <- subset(file_list, grepl(pattern = "Inspek", x = file_list))
  
  dat_b <- read_excel(paste0(input_path, bifang_file),
                      sheet = 1,
                      col_names = FALSE)
  dat_i <- read_excel(paste0(input_path, inspek_file),
                      sheet = 1,
                      col_names = FALSE)
  
  # Get Total kg from inspection report
  position_tot_wgt_1 <- which(apply(dat_i, 1, function(x) any(grepl("Tilmeldt art", x))))
  tot_wgt_i <- dat_i[position_tot_wgt_1, grep("Tilmeldt art", dat_i) + 4]
  colnames(tot_wgt_i) <- c("tot_wgt_i")
  
  # Get header info from inspection report
  position_head_1 <- which(apply(dat_i, 1, function(x) any(grepl("Rapport Nr.:", x))))
  head_1 <- dat_i[position_head_1, ]
  report_no <- head_1[, grep("Rapport Nr.:", head_1) + 3]
  colnames(report_no) <- c("report_no")
  
  position_head_2 <- which(apply(dat_i, 1, function(x) any(grepl("Nr.:", x))))
  position_head_2 <- subset(position_head_2, !(grepl(pattern = position_head_1, x = position_head_2)))
  head_2 <- dat_i[position_head_2, ]

  fid <- head_2[, grep("Nr.:", head_2) + 3]
  fid_no <- str_extract(fid, "[[:digit:]]{2,3}")
  fid_char <- gsub("[[:digit:]].*","", fid)
  fid_char_1 <- gsub(" ","", fid_char)
  fid <- paste0(fid_char_1, fid_no)
  fid <- as.data.frame(fid)
  
  position_head_3 <- which(apply(dat_i, 1, function(x) any(grepl("Dato.:|Start dato.:", x))))
  head_3 <- dat_i[position_head_3, ]

  date <- head_3[, grep("Dato.:|Start dato.:", head_3) + 3] # This need to be translated into a proper date
  colnames(date) <- c("date")
  
  head_done <- cbind(fid, report_no, tot_wgt_i, date)
  
  # Find samples ----
  position_head_1 <- which(apply(dat_b, 1, function(x) any(grepl("Art:", x))))
  head_1 <- dat_b[position_head_1[1], ]
  fishery <- str_extract(head_1[, grep("Art:", head_1) + 1], '[A-Za-z]+')
  fishery <- species[species$art == fishery, ]$engkode
  
  
  position_samples_head <- which(apply(dat_b, 1, function(x) any(grepl("Total bifangst:", x))))
  samples_head <- dat_b[position_samples_head, ]
  samples_head <- as.data.frame(as.character(samples_head))
  samples_head$`as.character(samples_head)`[samples_head$`as.character(samples_head)`== "Gobler\r\n(Vandmæ-nd)AJQ"] <- "AJQ"
  samples_head[is.na(samples_head) | samples_head == "NA"] <- "X"

  position_samples <- which(apply(dat_b, 1, function(x) any(grepl("Prøve nr.", x))))

  samples <- dat_b[c(position_samples:nrow(dat_b)), ]
  samples_1 <-  samples   #subset(samples, !is.na(...9))

  colnames(samples_1) <- samples_head[1:nrow(samples_head), ]

  # Remove columns after Select number of rows:
  position_sel_row <- which(colnames(samples_1)=="Alle arter")
  if (length(position_sel_row) == 1) {
    samples_1 <- samples_1[, -c((position_sel_row + 1):ncol(samples_1))]
  } 
  # else {
  #   position_sel_row_2 <- which(colnames(samples_1)=="Total bi-\r\nfangst (kg)")
  #   samples_1 <- samples_1[, -c(position_sel_row_2:ncol(samples_1))]
  # }

  

col_names <- names(samples_1)
col_names <- gsub(".*\\s", "", col_names)
col_names <- gsub("Art:", "All", col_names)

for (i in seq_along(col_names)) {
  
  fao_code <- species[species$art == col_names[i], ]
  if (length(fao_code ==1)) {
  col_names[i] <- fao_code$engkode
  }
}

  # for (i in seq_along(col_names)) {
  # 
  #   if (substr(col_names[i], 1, 1) == "X") {
  #     if (is.na(samples_1[2, i])) {
  #       col_names[i] <- col_names[i]
  #     } else if (samples_1[2, i] == "kg") {
  #       col_names[i] <- col_names[i-1]
  #     } else if (samples_1[2, i] %in% c("%"))
  #       col_names[i] <- col_names[i-1]
  #   } }

  for (i in seq_along(col_names)) {
    if (is.na(samples_1[2, i])) {
      col_names[i] <- col_names[i]
    } else if (samples_1[2, i] == "kg") {
      col_names[i] <- paste0(col_names[i], "_kg")
    } else if (samples_1[2, i] %in% c("%")) {
      col_names[i] <- paste0(col_names[i], "_pct")
    } else {
      col_names[i] <- col_names[i]
    }
  }

  colnames(samples_1) <- col_names
  if (fishery == "MAC") {
    samples_1$MAC_kg <- as.numeric(samples_1$All) - as.numeric(samples_1$arter)
  }
  
  

  # # Clean up a bit
  samples_2 <- samples_1 %>% select(-contains("X"), -contains("(kg)"), -contains("Prøven"), -contains("snit"), -contains("Tid"))
  samples_2 <- subset(samples_2, !(`bifangst:` %in% c("Prøve nr.", "Total:")))
  samples_2 <- rename(samples_2, sample_no = `bifangst:`)
  samples_2$sample_no <- as.numeric(samples_2$sample_no)
  samples_2 <- subset(samples_2, !(is.na(sample_no)))
  samples_2 <- subset(samples_2, !(is.na(All)))
  
  samples_2 <- select(samples_2, -All, -arter)

  # Transpose
  samples_3 <- tidyr::gather(samples_2, key = spp_fao, value = amount, -sample_no)
  samples_3$kg <- NA
  samples_3$kg[grepl("kg", samples_3$spp_fao) == T] <-
    as.numeric(samples_3$amount[grepl("kg", samples_3$spp_fao) == T])
  samples_3$spp_fao <- substr(samples_3$spp_fao, 1, 3)

kg <- subset(samples_3, !is.na(kg))
kg <- select(kg, sample_no, spp_fao, kg)

samples_4 <- kg

  # # Add global info

  samples_4$fishery <- fishery
  samples_4$path <- input_path
  samples_4$filename <- id

  samples_5 <- cross_join(samples_4, head_done)

  return(samples_5)
  
}
