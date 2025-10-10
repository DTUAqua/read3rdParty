library(pacman)
p_load(data.table,ggplot2,ggrepel,tidyverse,haven,zoo,RColorBrewer,
       rmarkdown,gridExtra,knitr)

library(readxl)
library(openxlsx)
library(visdat)
library(dplyr)

file_path <- "Q:/20-forskning/20-dfad/data/Data/additional_data/SKAW/get_data_from_AQUA-artsdata/Artsdata_download_2025/310725-125252_RI 376 til DTU.xlsx"


### Function to look through different sheets

search_sheet_by_name <- function(file_path) {
  sheet_names <- excel_sheets(file_path)
  
  
  if ("Ark1" %in% sheet_names && "Table 1" %in% sheet_names && "Ark2" %in% sheet_names) {
    sheet_name <- "Ark1"
  } else if ("Ark1" %in% sheet_names) {
    sheet_name <- "Ark1"
  } else if ("Table 1" %in% sheet_names) {
    sheet_name <- "Table 1"
  } else if ("Ark2" %in% sheet_names) {
    sheet_name <- "Ark2"
  } else if ("DTU_7" %in% sheet_names) {
    sheet_name <- "DTU_7"
  } else {
    print("Sheets 'Ark1' and 'Table1' not found in the workbook.")
    return(NULL)
  }
  
  
  
  my_data <- read_excel(file_path, sheet = sheet_name,col_names = FALSE, )
  
  my_data <- my_data[ , -c(1:7)]
  
  # print(paste("Found sheet '", sheet_name, "'", sep = ""))
  
  row_with_word <- which(apply(my_data, 1, function(x) any(grepl("Prøve nr.", x))))
  # print(row_with_word)
  colnames(my_data) <- as.character(my_data[row_with_word,])
  
  new_data <- my_data[row_with_word:nrow(my_data), ]
  ignored_data <- my_data[1:(row_with_word - 1), ]
  
  
  
  
  
  # new_data <- new_data[-1, ]
  new_data <- new_data[-1, ]
  
  
  # 
  # a <- names(new_data)[!is.na(names(new_data))]
  # 
  # # colnames(new_data) <- as.character(new_data[1, ])
  # b <- a[a != "NA" ]
  # 
  
  col_names <- names(new_data)
  
  for (i in seq_along(col_names)) {
    if (is.na(col_names[i]) || grepl("^NA\\.?\\d*$", col_names[i]) || col_names[i] == "") {
      if (i > 1) {
        col_names[i] <- paste0(col_names[i - 1], "_KG")
      } else {
        col_names[i] <- "Unknown_KG"  # fallback for first column if it's NA
      }
    }
  }
  
  
  names(new_data) <- col_names
  
  final_df <- new_data
  
  final_df <- data.frame(lapply(final_df, function(x) type.convert(x, as.is = TRUE)))
  ignored_df <- data.frame(lapply(ignored_data, function(x) type.convert(x, as.is = TRUE)))  
  
  
  colnames(final_df) <- gsub("\\.", "", colnames(final_df))
  final_df <- final_df[rowSums(is.na(final_df)) != ncol(final_df),]
  
  
  
  if (ncol(final_df) > 0 && colnames(final_df)[1] %in% c("FiskeartKode","FiskeartKade","Fiskeart","Fiskeart: Kode:","Fiskeart:")) {
    fiskeart_kode_value <- final_df[1, 1]  # Get the value in the first row and first column
    
    # Replace the column name with the actual value
    colnames(final_df)[1] <- fiskeart_kode_value
    
    # Delete the first row
    final_df <- final_df[-1, ]
  }
  
  # final_df2 <- final_df
  columns_to_remove <- grepl("ProcentPrøve|ProcentPrøven|Procentprøve|Procentprøven|ProcentPrnve|ProoentPrøve|Tid", colnames(final_df), ignore.case = TRUE)
  columns_to_remove <- which(columns_to_remove)
  final_df <- final_df[, 1:(min(columns_to_remove, na.rm = TRUE) - 1)]
  
  
  ignored_columns <- colnames(ignored_df)[colSums(!is.na(ignored_df)) == 0]
  ignored_df<-(ignored_df[, !(colnames(ignored_df) %in% ignored_columns)])
  
  ignored_df_t <- data.frame(t(ignored_df),row.names = NULL)
  
  first_row <-   unlist(ignored_df_t[1,])
  
  colnames(ignored_df_t) <- as.character(first_row)
  ignored_df_t <- ignored_df_t[-1, ]
  
  ### This is due to it being observed that the bad formated files do not have all the columns so to avoid that
  ### we ignore the files with columns less then 7

  if ( ncol(ignored_df_t) < 9 || is.null(ncol(ignored_df_t))) {
    print("Ignoring file due to insufficient columns and copying to another folder")
    return(file_path) 
  } else if(ncol(ignored_df_t) <= 9) {
    
    ignored_df_t <- ignored_df_t[1,1:7]
    
    # Getting the boat IDs out from Fartøj
    
    
    ignored_df_t$Fartøj <- sub("^([A-Z]+\\s+\\d+).*", "\\1", ignored_df_t$Fartøj)
    
    
    
    # Combining the values of the dataframes
    combined_df <- cbind(final_df, ignored_df_t[rep(seq_len(nrow(ignored_df_t)), each = nrow(final_df)), ])
    colnames(combined_df) <- gsub(" ", "", colnames(combined_df))
    
    
    
    
    
    
    
    
    
    # Fixing the date
    origin_date <- as.Date("1899-12-30")
    
    combined_df <- combined_df %>%
      mutate(
        Dato = format((origin_date + days(as.numeric(Dato))),format = "%d-%m-%Y"))
    
    
    #Changing the column name so it can be in the same column in final file
    
    column_replacements <- c("Totalvejet" = "Totalvejetfisk",
                             "PrøvenrKode" = "Prøvenr:",
                             "Prøvenr" = "Prøvenr:",
                             "Prøvenr." = "Prøvenr:",
                             "FiskeartKodePrøvenr" = "Prøvenr:",
                             "Pøvenr." = "Prøvenr:",
                             "Prnvenr" = "Prøvenr:",
                             "Pr0venr" = "Prøvenr:",
                             "Prnvenr:" = "Prøvenr:",
                             "Prnvenr." = "Prøvenr:",
                             "Pr0venr." = "Prøvenr:"
    )
    
    # Iterate over the vector of column names and perform substitutions
    for (col_name in names(column_replacements)) {
      if (col_name %in% colnames(combined_df)) {
        colnames(combined_df) <- sub(col_name, column_replacements[col_name], colnames(combined_df))
      }
    }
    
    # Removing Columns which are of no use
    column_to_drop <- c("Totalbifangstkg","Totalheleprøvekg","Totalheleprøvenkg","Totalbiangstkg",
                        "Totalhelepøvenkg","Tolalheleprøvenkg","TotalhelePrøvekg","Total bifangst (kg)","Tid")
    
    combined_df <- combined_df[ , !(names(combined_df) %in% column_to_drop)]
    
    
    
    
    ###CHECK PATH BEFORE RUNNING!! Using sas file to get the species code and match to columns
    sad_code <- read_sas("C:/Users/emilb/Downloads/art (3).sas7bdat",NULL)
    species_code <- sad_code[-1,c("start","art","engkode")]
    
    new_column_names <- list()
    # Iterate over the column names
    for (i in seq_along(colnames(combined_df))) {
      column_name <- colnames(combined_df)[i]
      first_word <- unlist(strsplit(column_name, "(?<=[a-z])(?=[A-Z]|kg|$)", perl = TRUE))[1]
      
      
      matched_row <- species_code[species_code$art == first_word, ]
      
      if (nrow(matched_row) > 0) {
        if (!is.na(matched_row$engkode[1])) {
          new_column_name <- matched_row$engkode[1]
        } else {
          new_column_name <- matched_row$start[1]
        }
        
        colnames(combined_df)[i] <- new_column_name
        
        new_column_names <- c(new_column_names, new_column_name)
        
        
      }
    }
    
    print(new_column_names)
    
    # If there is no value present or if there is a misspell we check if its present in this
    
    species_code_list <- c("SPR","MAC","USO","HER","WHG","GUG","GUU","CRA","CEP","DAB","PLA","WHB","HAD","BOR","HOM","ANE",
                           "PLE","CPR","SAN","MY","NOP","GOG","MYG","ELP","ARY","GDG","SOL","WEG","MON","PRA","MZZ",
                           "OPH","LAU","AJQ","REG","RED","USO")
    
    
    # Extract code list part from column names
    new_colnames <- sub(paste0(".+(", paste(species_code_list, collapse = "|"), ").+"), "\\1", colnames(combined_df))
    
    
    colnames(combined_df) <- new_colnames
    combined_df2 <- combined_df
    combined_df <- combined_df2
    ### To avoid having a double dot value in the species column
    # combined_df <- combined_df[grep("Total",combined_df$`Prøvenr:`,invert = TRUE),]
    combined_df <- combined_df[grep("Bifangst",combined_df$`Prøvenr:`,invert = TRUE),]
    combined_df <- combined_df[grep("Fordeling",combined_df$`Prøvenr:`,invert = TRUE),]
    combined_df <- combined_df[grep("Fordelng",combined_df$`Prøvenr:`,invert = TRUE),]
    combined_df <- combined_df[grep("Fordellng",combined_df$`Prøvenr:`,invert = TRUE),]
    combined_df <- combined_df[grep("%",combined_df$`Prøvenr:`,invert = TRUE),]
    combined_df <- combined_df[grep("kg",combined_df$`Prøvenr:`,invert = TRUE),]
    combined_df <- combined_df[grep("ka",combined_df$`Prøvenr:`,invert = TRUE),]
    # 
    # # FIX OF the issue related to the fullstop values
    # for (col_name in colnames(combined_df)) {
    #   if (col_name == "Totalvejetfisk" ) {
    #     if (any(sapply(combined_df[[col_name]], function(x) grepl("^\\d+\\.\\d+\\.\\d+$", x)))) {
    #       combined_df[[col_name]] <- as.numeric(gsub("\\.", "", combined_df[[col_name]]))
    #     } else {
    #       combined_df[[col_name]] <- round(as.numeric(combined_df[[col_name]]), 3)
    #     }
    #   } else if (col_name == "Totalvejet") {
    #     combined_df[[col_name]] <- as.numeric(gsub("\\.", "", combined_df[[col_name]]))
    #   }
    # }
    
    
    miss_spelled <- c(    "Lillefjæsingkg" = "TOZ",
                          "Sølvoksekg" = "QGE",
                          "Blahvillingkg" = "WHB",
                          "Spelingkg" = "NOP",
                          "Lampretkg" = "LAS",
                          "Rejerkg" = "DCP",
                          "Goplerkg" = "SZY",
                          "Tangålkg" = "SWY",
                          "Ålebrosmekg" = "JXQ",
                          "Hestemakelkg" = "HOM"
    )
    
    # Iterate over the vector of column names and perform substitutions
    for (col_name in names(miss_spelled)) {
      if (col_name %in% colnames(combined_df)) {
        colnames(combined_df) <- sub(col_name, miss_spelled[col_name], colnames(combined_df))
      }
    }
    
    
    
    
    column_kg_to_drop <- c("Xkg","Xkg1","Xkg2","Xkg3")
    
    combined_df <- combined_df[ , !(names(combined_df) %in% column_kg_to_drop)]
    
    
    
    
    
    
    ## Converting Species Column names to Numeric
    
    
    
    # Apply the conversion using lapply
    for (col_name in species_code_list) {
      if (exists(col_name, combined_df)) {
        combined_df[[col_name]] <- round(as.numeric(combined_df[[col_name]]), 2)
      }
    }
    
    
    # Resetting the row index
    rownames(combined_df) <- NULL
    
    col_names <- names(combined_df)
    
    # Initialize a new vector for renamed columns
    new_names <- col_names
    
    # Loop through column names
    for (i in seq_along(col_names)) {
      base_name <- sub("\\.\\d+$", "", col_names[i])  # remove .1, .2 etc. to get base name
      
      # Count how many times base_name appears (including .1, .2 etc.)
      matching_cols <- grep(paste0("^", base_name, "(\\.\\d+)?$"), col_names)
      
      # If there are exactly 2 matching (e.g. Apple and Apple.1)
      if (length(matching_cols) == 2) {
        first_idx <- matching_cols[1]
        second_idx <- matching_cols[2]
        
        new_names[first_idx] <- paste0(base_name, " (stk)")
        new_names[second_idx] <- paste0(base_name, " (kg)")
      }
    }
    
    # Assign the new names to the dataframe
    names(combined_df) <- new_names
    
    
    
    
    
    
    
    return(combined_df)
  }
}

### For checking Individual files
# file_path <- "C:/Users/muhaadu/Desktop/SKAW/For_testing_comb/211207-084614-055080-hg62_CLEAN  - Kopi.xlsx"
file_path <- "C:/Users/emilb/Desktop/DPPO/SKAW data/2025/pdf/pdf/060525-080847_HG 62 Beinur - DTU_CLEAN.xlsx"
file_path <-"C:/Users/emilb/Desktop/DPPO/SKAW data/2025/Til Nis, sandeel/pdf/CLEAN/020525-030129_S 204 RAPPORT_CLEAN.xlsx"

# 
result <- search_sheet_by_name(file_path)


###### To process all the files (Working for the right formated ones)


### Path to all files
dir_path_all <- "C:/Users/emilb/Desktop/DPPO/SKAW data/2025/Til Nis, sandeel/pdf/CLEAN/" 
# 
# ### For debugging
# 
# # output_path <- "C:/Users/muhaadu/Desktop/SKAW/Processed/"
# # unprocesed <- "C:/Users/muhaadu/Desktop/SKAW/weird/"
# 
# ### The path to put the wrong formatted files 
# un_procesed <- "//ait-pdfs.win.dtu.dk/Qdrev/aqua/dfad/data/Data/additional_data/SKAW/muhaadu/Unprocessed/"
# 
# 
file_list <- list.files(path = dir_path_all, pattern = "\\.xlsx$", full.names = TRUE)
# 
combined_df <- data.frame()
# 
# 
for (file in file_list) {
  print(file)

  final_df <- search_sheet_by_name(file)
  final_df$Filename <- file

  if (is.data.frame(final_df)) {
    # The value of final_df is a dataframe
    combined_df <- bind_rows(combined_df, final_df)

  } else {
    # The value of final_df is a file path
    # Move the file to another folder
    file.copy(file, un_procesed, overwrite = TRUE)
  }
}


combined_df <- combined_df %>% select( "Prøvenr:", SAN, `HER (stk)` , `HER (kg)`, `CEP (stk)`, `CEP (kg)`, `WHG (stk)`, `WHG (kg)`, `MAC (stk)`, `MAC (kg)`, `GUG (stk)`, `GUG (kg)`, `SPR (stk)`, `SPR (kg)`, `HAD (stk)`, `HAD (kg)`, `AJQ (stk)`, `AJQ (kg)`, `DAB (stk)`, `DAB (kg)`, `HOM (stk)`, `HOM (kg)`, `WIT (stk)`, `WIT (kg)`, USO, UsorterbartUSO, Fiskeri, Fartøj, Dato,  Status, Køber, Filename)

# 
# 
# 
# 
# 
write.csv2(combined_df, file = "SKAW_SAN_24_25.csv", row.names = FALSE,fileEncoding = "UTF-8")
# 
# test <- read.csv2("test.csv")
# 
























