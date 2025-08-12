# -3. Table of Contents
#  -3) Table of Contents
#  -2) Room for improvement
#  -1) Instructions before running
#   0) Setup
#   1) Function to create Upload files
#   2) Detect what needs to be updated
#   3) Loop through the lines of business that need to be updated, applying the function to each
#   4) Create LDF tool
#   5) Documentation
#   6) Next year's Exhibits folder
#   7) Duration log


# -2. Room for improvement
#   A) Create "Combine" tab with different aggregation for each LOB (maybe allow input to determine aggregation bases). Import "Combine" results into UploadTable.
#   B) Format numeric cells in label columns as numbers in a way that doesn't take too long.
#   C) Apply LDF_style more


# -1. Instructions before running
#   A) For CA, use "Custom Exhibits"
#   B) Ensure the ISO Exhibits are located in sprintf(
#     "G:\\Actuarial\\commercial lines     division\\Resources\\ISO\\%s\\LDF\\Exhibits"
#     , currentYear)
#   C) Ensure every ISO Exhibit has the exact string, "Exhibit" somewhere in its name


# 0. Setup
start_time <- Sys.time()

##install.packages("dplyr")
##install.packages("openxlsx")
##install.packages("officer")
##install.packages("rstudioapi")
library(dplyr)
library(openxlsx)
library(officer)
library(rstudioapi) # Am I actually using this? If yes, it's by the documentation

currentYear <- format(Sys.Date(), "%Y")
##currentYear <- as.character(as.integer(currentYear) - 1) # Activate for baseline creation; otherwise, comment out
lastYear <- as.character(as.integer(currentYear) - 1)
nextYear <- as.character(as.integer(currentYear) + 1)

# Number of development columns we want
LDF_width <- 14
LDF_cols <- as.character(3 + 12 * (1:LDF_width))

# Decimal format for development
LDF_style <- createStyle(numFmt = "0.00")

# Functions for building and saving workbooks
addSheet <- function(workbook, sheet_name, data, visibility = TRUE, writeColNames = TRUE, filter = TRUE) {
  
  data <- as.data.frame(data)
  
  addWorksheet(workbook, sheet_name, visible = visibility)
  writeData(workbook, sheet_name, data, startCol = 1, startRow = 1, colNames = writeColNames)
  
  # Name the region
  end_row <- nrow(data)
  if (writeColNames) {
    end_row <- end_row + 1
  }
  
  createNamedRegion(wb = workbook, sheet = sheet_name, name = gsub(" ", "_", sheet_name), rows = 1:end_row, cols = 1:ncol(data))
  
  # Add filters if appropriate
  if (filter) {
    addFilter(workbook, sheet_name, row = 1, cols = 1:ncol(data))
  }
  
  # Style LDF columns
  col_indices <- which(names(data) %in% LDF_cols)
  
  for (col in col_indices) {
    addStyle(
      workbook,
      sheet = sheet_name,
      style = LDF_style,
      cols = col,
      rows = 2:end_row,
      gridExpand = TRUE
    )
  }          
  
}

clear <- function(file_path) {
  if (file.exists(file_path)) {
    file.remove(file_path)
  }
}

revise <- function(file_path) {
  original_file_path <- file_path
  base <- substr(original_file_path, 1, nchar(original_file_path) - 5)
  ending <- substr(original_file_path,
                   nchar(original_file_path) - 4,
                   nchar(original_file_path)) # Running Revise immediately after Clear allows the program to run even if this year's previous version is open
  
  n <- 1
  while (file.exists(file_path)) {
    file_path <- paste0(base, "(", n, ")", ending)
    n <- n + 1
  }
  
  return(file_path)
}

# Function for formatting column headers
capitalize_first <- function(x) {
  sapply(x, function(str) {
    if (is.na(str) || str == "")
      return(str)
    paste0(toupper(substr(str, 1, 1)), tolower(substr(str, 2, nchar(str))))
  }, USE.NAMES = FALSE)
}

# Function for retrieving the script's name
get_script_name <- function() {
  if (requireNamespace("rstudioapi", quietly = TRUE) && rstudioapi::isAvailable()) {
    context <- rstudioapi::getActiveDocumentContext()
    return(basename(context$path))
  } else {
    return(NULL)  # Not running in RStudio or rstudioapi not available
  }
}


# 1. Function to create Upload files
# A) Table of Contents
#   B. Construct file paths
#   C. Load new data
#   D. Use new data to construct new DZ Data tab
#   E. Use new DZ Data tab to construct new UploadTable tab
#   F. If relevant, use old and new UploadTables to construct new ComptoPrior tab
#   G. Build the final product


update <- function(LOB, exhibit_folder) {
  # B) Define file paths
  new_data_path <- file.path(exhibit_folder, LOB)
  
  LOB_title <- substr(LOB, 4, 5)
  if (LOB_title == "CA") {
    LOB_title <- "AL"
  }
  
  old_workbook_path <- sprintf(
    "G:\\Actuarial\\commercial lines division\\Resources\\ISO\\%s\\LDF\\Uploads\\%s_Upload_%s.xlsx",
    lastYear,
    LOB_title,
    lastYear
  )
  ##new_workbook_path <- paste0(new_folder_path, "\\", sprintf("%s_Upload_%s.xlsx", LOB_title, currentYear)) # Moved to #save section at end of function
  # sprintf("G:\\Actuarial\\commercial lines division\\Resources\\ISO\\%s\\LDF\\Uploads\\%s_Upload_%s - Test.xlsx", currentYear, LOB_title, currentYear)
  
  # print(new_data_path)
  print(LOB_title)
  # print(old_workbook_path)
  # print(new_workbook_path)
  
  ##cat(paste(new_data_path, old_workbook_path, new_workbook_path, sep = "\n"))
  
  # C) Load new data
  new_data <- loadWorkbook(new_data_path) %>%
    read.xlsx(sheet = "DZ Data", colNames = FALSE)
  
  ##print(new_data)
  
  
  # D) Use new data to construct new DZ Data tab
  print("Building DZ Data tab")
  
  # Tab Structure
  # Row and column clean
  row_index <- which(apply(new_data, 1, function(row)
    any(grepl(
      "year", row, ignore.case = TRUE
    ))))[1]
  
  colnames(new_data) <- as.character(unlist(new_data[row_index, ]))
  
  new_DZ_Data <- new_data[(row_index + 1):nrow(new_data), ]
  
  threshold <- 0.99
  non_empty_cols <- sapply(new_DZ_Data, function(col) {
    mean(is.na(col) | col == "") < threshold
  })
  new_DZ_Data <- new_DZ_Data[, non_empty_cols]
  
  # Format column names
  colnames(new_DZ_Data) <- capitalize_first(colnames(new_DZ_Data))
  
  # Reorder columns
  year_col <- names(new_DZ_Data)[grepl("year", names(new_DZ_Data), ignore.case = TRUE)]
  loss_cols <- names(new_DZ_Data)[grepl("month", names(new_DZ_Data), ignore.case = TRUE)]
  label_cols <- setdiff(names(new_DZ_Data), c(year_col, loss_cols))
  
  new_DZ_Data <- new_DZ_Data[, c(label_cols, year_col, loss_cols)]
  
  # Reorder rows to ensure triangles
  new_DZ_Data <- new_DZ_Data %>%
    arrange(across(all_of(label_cols)), across(all_of(year_col)))
  
  rownames(new_DZ_Data) <- NULL
  
  # Format columns
  new_DZ_Data[, c(year_col, loss_cols)] <- sapply(new_DZ_Data[, c(year_col, loss_cols)], function(x)
    suppressWarnings(as.numeric(x)))
  
  if (new_DZ_Data[1, year_col] > 10000) {
    new_DZ_Data[[year_col]] <- as.Date(new_DZ_Data[[year_col]], origin = "1899-12-30")
  }
  
  # Add analysis columns
  new_columns <- as.data.frame(matrix(NA, nrow = nrow(new_DZ_Data), ncol = 16))
  colnames(new_columns) <- c("Blank", LDF_cols, "Tag")
  
  new_DZ_Data <- cbind(new_DZ_Data, new_columns)
  
  # Triangle clean
  triangle_ends <- which((!is.na(new_DZ_Data[[loss_cols[1]]])) &
                           (is.na(new_DZ_Data[[loss_cols[2]]])))
  triangle_height <- triangle_ends[2] - triangle_ends[1]
  
  rows_to_delete <- c()  # Initialize an empty vector to store row indices
  
  for (start_row in (triangle_ends - triangle_height + 1)) {
    test_value <- suppressWarnings(as.numeric(as.matrix(new_DZ_Data[start_row, loss_cols[[1]]])))
    
    if (is.na(test_value) || test_value == 0) {
      print(test_value)
      Tag <- paste(new_DZ_Data[start_row, label_cols], collapse = " ")
      rows_to_delete <- c(rows_to_delete, start_row + (0:(triangle_height - 1)))
      cat("Marked for deletion:", Tag, "\n")
    }
  }
  
  if (length(rows_to_delete) > 0) {
    rows_to_delete <- sort(unique(rows_to_delete))
    new_DZ_Data <- new_DZ_Data[-rows_to_delete, ]
  }
  
  
  
  ##print(colnames(new_DZ_Data))
  ##print(new_DZ_Data) # structure without output
  
  # Triangle and output dimensions
  triangle_ends <- which((!is.na(new_DZ_Data[[loss_cols[1]]])) &
                           (is.na(new_DZ_Data[[loss_cols[2]]])))
  ##print(triangle_ends)
  
  triangle_length <- length(loss_cols)
  
  output_length <- min(triangle_height - 3, triangle_length - 1, LDF_width)
  ##cat("Output length:", output_length, "\n")
  
  for (end_row in triangle_ends) {
    # LDFs
    for (i in 1:LDF_width) {
      if (i <= output_length) {
        early <- loss_cols[i]
        late <- loss_cols[i + 1]
        
        rows <- (end_row - i - 2):(end_row - i)
        
        early_vals <- new_DZ_Data[rows, early]
        late_vals <- new_DZ_Data[rows, late]
        
        ##cat("Rows:", rows, "\n")
        ##cat("Early:", early, "\n")
        ##cat("Late:", late, "\n")
        ##cat("Early Values:", early_vals, "\n")
        ##cat("Late Values:", late_vals, "\n")
        
        LDF <- sum(late_vals) / sum(early_vals)
        
      } else {
        LDF <- 1.000
        
      }
      
      new_DZ_Data[end_row, LDF_cols[i]] <- LDF
      
    }
    
    ##print(new_DZ_Data[end_row, LDF_cols])
    
    # ATUs
    for (i in rev(seq_along(LDF_cols))) {
      if (i == LDF_width) {
        ATU <- new_DZ_Data[end_row, LDF_cols[i]]
      } else {
        ATU <- (new_DZ_Data[end_row, LDF_cols[i]] * ATU)
      }
      
      new_DZ_Data[end_row - 1, LDF_cols[i]] <- ATU
      
    }
    
    ##print(new_DZ_Data[endrow - 1, LDF_cols])
    
  }
  
  # Tag column
  new_DZ_Data$Tag <- apply(new_DZ_Data[label_cols], 1, function(row)
    paste(as.character(row), collapse = " "))
  
  ##print(new_DZ_Data)
  
  
  # E) Use new DZ Data tab to construct new UploadTable tab -- Missing combined rows and "NEW" in column W
  print("Building UploadTable tab")
  
  selected_rows <- triangle_ends - 1
  selected_cols <- unique(c(label_cols, LDF_cols, "Tag"))
  
  new_UploadTable <- new_DZ_Data[selected_rows, selected_cols]
  new_UploadTable <- cbind(LOB = LOB_title, new_UploadTable)
  
  
  rownames(new_UploadTable) <- NULL
  
  ##print(new_UploadTable)
  
  # addWorksheet(tool_workbook, LOB_title)
  # writeData(tool_workbook, LOB_title, new_UploadTable)
  
  
  # F) If relevant, use old and new UploadTables to construct new ComptoPrior tab
  ##print(old_workbook_path)
  if (exists("new_ComptoPrior")) {
    rm("new_ComptoPrior", envir = .GlobalEnv)
  }
  
  if (file.exists(old_workbook_path)) {
    old_workbook_wb <- loadWorkbook(old_workbook_path)
    sheet_names <- names(old_workbook_wb)
    
    if ("Metadata" %in% sheet_names) {
      metadata <- readWorkbook(old_workbook_wb, sheet = "Metadata")
      
      if ("Note" %in% names(metadata) &&
          any(metadata$Note == "Created by R script")) {
        print("Building ComptoPrior")
        
        old_UploadTable <- loadWorkbook(old_workbook_path) %>%
          read.xlsx(sheet = "UploadTable", colNames = TRUE)
        
        blank_rows <- as.data.frame(matrix(
          NA,
          nrow = 2,
          ncol = ncol(old_UploadTable)
        ))
        blank_rows[2, ] <- colnames(old_UploadTable)
        colnames(blank_rows) <- colnames(old_UploadTable)
        
        new_ComptoPrior <- rbind(old_UploadTable, blank_rows, old_UploadTable)
        
        
        for (i in seq_len(nrow(old_UploadTable))) {
          tag_value <- old_UploadTable$Tag[i]
          matching_row <- new_UploadTable[new_UploadTable$Tag == tag_value, LDF_cols]
          
          if (length(matching_row) > 0) {
            old_LDFs <- as.numeric(old_UploadTable[i, LDF_cols])
            new_LDFs <- as.numeric(matching_row[1, ])
            
            percent_change <- new_LDFs / old_LDFs - 1  # Keep as numeric
            
          } else {
            percent_change <- rep("Discontinued", length(LDF_cols))
          }
          
          new_ComptoPrior[nrow(old_UploadTable) + 2 + i, LDF_cols] <- percent_change
          
        }
        
        
        ##print(new_ComptoPrior)
        
      }
    }
  } else {
    print("No prior data to use to build ComptoPrior")
    if (exists("new_ComptoPrior")) {
      rm("new_ComptoPrior", envir = .GlobalEnv)
    }
  }
  
  
  # G) Build the final product
  print("Building final product")
  finalProduct <- createWorkbook()
  
  if (exists("new_ComptoPrior")) {
    print("Adding ComptoPrior tab")
    
    addSheet(finalProduct, "ComptoPrior", new_ComptoPrior)
    
    percent_style <- createStyle(numFmt = "0.00%")
    
    for (col_name in LDF_cols) {
      col_index <- which(names(new_ComptoPrior) == col_name)
      start_row <- nrow(old_UploadTable) + 3  # +3 to account for offset
      end_row <- start_row + nrow(old_UploadTable) - 1
      
      addStyle(
        finalProduct,
        sheet = "ComptoPrior",
        style = percent_style,
        rows = start_row:end_row,
        cols = col_index,
        gridExpand = TRUE
      )
    }
    
  }
  
  print("Adding UploadTable Tab")
  addSheet(finalProduct, "UploadTable", new_UploadTable)
  
  print("Adding DZ Data tab")
  addSheet(finalProduct, "DZ Data", new_DZ_Data)
  
  print("Adding Metadata tab")
  metadata <- data.frame(Note = "Created by R script")
  addSheet(finalProduct, "Metadata", metadata, filter = FALSE)
  metadata_index <- which(names(finalProduct) == "Metadata")
  sheetVisibility(finalProduct)[metadata_index] <- "hidden"
  
  # Format
  bold_style <- createStyle(textDecoration = "bold")
  label_style <- createStyle(numFmt = "0")
  lUT_style <- createStyle(numFmt = "0.000")
  
  for (sheet_name in names(finalProduct)) {
    sheet_data <- read.xlsx(finalProduct, sheet = sheet_name, colNames = FALSE)
    num_cols <- ncol(sheet_data)
    
    addStyle(
      finalProduct,
      sheet = sheet_name,
      style = bold_style,
      rows = 1,
      cols = 1:num_cols,
      gridExpand = TRUE
    )
  }
  
  # I'm trying to format numeric data as numbers, but it takes too long
  # sheet_names <- names(finalProduct)
  #
  # for (sheet in sheet_names) {
  #   data <- readWorkbook(finalProduct, sheet = sheet)
  #   col_names <- colnames(data)
  #
  #   for (j in seq_along(col_names)) {
  #     col_name <- col_names[j]
  #
  #     # Determine which style to apply
  #     if (col_name %in% label_cols) {
  #       style_to_apply <- label_style
  #     } else if (col_name %in% LDF_cols) {
  #       style_to_apply <- lUT_style
  #     }
  #
  #     # Loop through each row in the column
  #     for (i in seq_len(nrow(data))) {
  #       cell_value <- data[i, j]
  #       num <- suppressWarnings(as.numeric(as.character(cell_value)))
  #
  #       if (!is.na(num)) {
  #         addStyle(
  #           finalProduct, sheet = sheet, style = style_to_apply,
  #           rows = i + 1, cols = j, gridExpand = TRUE, stack = TRUE
  #         )
  #       }
  #     }
  #   }
  # }
  
  # Save
  new_workbook_path <- paste0(new_folder_path,
                              "\\",
                              sprintf("%s_Upload_%s.xlsx", LOB_title, currentYear))
  
  # Delete/revise path
  # Optional: Delete duplicate
  clear(new_workbook_path)
  # Revise path
  new_workbook_path <- revise(new_workbook_path)
  
  saveWorkbook(finalProduct, file = new_workbook_path, overwrite = TRUE)
  cat("Saved", LOB_title, "at", new_workbook_path, "\n", sep = " ")
  
}


# 2. Detect what needs to be updated
exhibit_folder <- sprintf(
  "G:\\Actuarial\\commercial lines division\\Resources\\ISO\\%s\\LDF\\Exhibits",
  currentYear
)

exhibits <- unique(list.files(
  path = exhibit_folder,
  pattern = "Exhibit",
  full.names = FALSE
))

print(exhibits)


# 3. Create new Upload folder and apply the function to each LOB that needs to be updated
new_folder_path <- sprintf(
  "G:\\Actuarial\\commercial lines division\\Resources\\ISO\\%s\\LDF\\Uploads",
  currentYear
)

# Delete/revise path
# Optional: Delete duplicate
if (dir.exists(new_folder_path)) {
  unlink(new_folder_path, recursive = TRUE, force = TRUE)
}

# Revise path
original_new_folder_path <- new_folder_path
n <- 1
while (dir.exists(new_folder_path)) {
  new_folder_path <- paste0(original_new_folder_path, "(", n, ")")
  n <- n + 1
}

dir.create(new_folder_path)

for (LOB in exhibits) {
  update(LOB, exhibit_folder)
}

##print(Sys.time() - start_time)

##stop("Still working on LDF_tool")
# 4. Create LDF_tool
print("Creating LDF_tool")

# Get UploadTables from new uploads
new_uploads <- list.files(new_folder_path, pattern = "\\.xlsx$", full.names = TRUE)
LOB_titles <- substr(list.files(new_folder_path, pattern = "\\.xlsx$", full.names = FALSE), 1, 2)

UT_list <- list()

for (i in seq_along(new_uploads)) {
  file_path <- new_uploads[i]
  
  UT_list[[i]] <- read.xlsx(file_path, sheet = "UploadTable")
}

names(UT_list) <- LOB_titles

full_width <- max(sapply(UT_list, ncol))
LDF_and_tag_width <- LDF_width + 1
label_width <- full_width - LDF_and_tag_width

for (i in seq_along(UT_list)) {
  UT <- UT_list[[i]]
  
  padding_width <- full_width - ncol(UT)
  
  if (padding_width > 0) {
    padding <- as.data.frame(matrix(
      "",
      nrow = nrow(UT),
      ncol = padding_width,
      dimnames = list(NULL, as.character(1:padding_width))
    ))
    
    UT_list[[i]] <- cbind(padding, UT)
  }
  
}

# Build
tool_workbook <- createWorkbook()

# UploadTables
buffer <- 5
for (i in seq_along(UT_list)) {
  
  addSheet(tool_workbook, names(UT_list)[i], UT_list[[i]])
  
  for (j in 1:(label_width)) {
    
    writeData(tool_workbook, names(UT_list[i]),
              colnames(UT_list[[i]])[j],
              startCol = full_width + buffer + j,
              startRow = 1)
    
  }
  
  formula <- paste0("=UNIQUE(A$2:A$",
                    nrow(UT_list[[i]]) + 1,
                    ")")
  
  unique_col <- full_width + buffer + 1
  writeData(tool_workbook, names(UT_list)[i],
            x = formula,
            startCol = unique_col,
            startRow = 2)
  
  for (j in unique_col + (1:(label_width - 1))) {
    writeData(tool_workbook, names(UT_list)[i],
              x = "Drag",
              startCol = j,
              startRow = 2)
  }
  
  max_unique <- max(sapply(UT_list[[i]][1:(label_width)], function(col) length(unique(col))))
  createNamedRegion(wb = tool_workbook, sheet = names(UT_list)[i], name = paste0(names(UT_list)[i], "_labels"),
                    rows = 1:(max_unique + 1),
                    cols = full_width + 5 + 1:(label_width))
  
}

# Stacked UploadTables -- Full tab
full_UT <- data.frame()
for (UT in UT_list) {
  names(UT) <- as.character(1:full_width)
  full_UT <- bind_rows(full_UT, UT)
}
addSheet(tool_workbook, "Full", full_UT, writeColNames = FALSE)

# Interpolation tab
interp_UT <- full_UT
names(interp_UT)[1:label_width] <- paste0("Label", 1:label_width)
names(interp_UT)[label_width + 1:LDF_width] <- 3 + 12*(1:LDF_width)
names(interp_UT)[full_width] <- "Tag"

n_rows <- nrow(interp_UT)

prefix1 <- c("Blank", "Interpolation")
prefix2 <- c("Blank", "Linear Interpolation")
interp_ages <- paste(3 + 12:(12 * LDF_width), "Interp")
linear_ages <- paste(3 + 12:(12 * LDF_width), "Linear")

interp_names <- c(prefix1, interp_ages)
linear_names <- c(prefix2, linear_ages)

interp_df <- as.data.frame(matrix(NA, nrow = n_rows, ncol = length(interp_names)))
names(interp_df) <- interp_names

linear_df <- as.data.frame(matrix(NA, nrow = n_rows, ncol = length(linear_names)))
names(linear_df) <- linear_names

tryCatch({
  for (row in 1:nrow(interp_df)) {
    
    ATUs <- unlist(as.numeric(interp_UT[row, label_width + 1:LDF_width]))
    ##cat("ATUs:", unlist(ATUs), "\n")
    inverse_factors <- 1 - (1 / as.numeric(ATUs))
    ##cat("Inverse factors:", unlist(inverse_factors), "\n")
    
    ratios <- inverse_factors[-1] / (inverse_factors[-length(inverse_factors)] + 1e-200)
    ##cat("Ratios:", ratios, "\n")
    steps <- c(ifelse(ratios >= -1e-199, ratios^(1 / 12), 1e100), 1)

    linear_steps <- c((ATUs[-1] - ATUs[-length(ATUs)])/12, 1)
    
    for (j in 3:(3 + 12 * (LDF_width-1))) {
      
      months <- j - 3 + 15
      
      x <- months - 3
      m <- x %/% 12
      d <- x %% 12
      
      if (d == 0){
        insert <- ATUs[m]
      }else if (steps[[m]] < 1e99){
        preconversion <- inverse_factors[[m]] * steps[[m]]^d
        ##cat("Preconversion:", preconversion, "\n")
        insert <- 1/(1-preconversion)
      }else{
        insert <- "Not available"
      }
      
      interp_df[row, j] <- insert
      
      linear_df[row, j] <- ATUs[m] + linear_steps[m]*d
      
    }
  
  }
  
}, error = function(e) {
  print("Interpolation error")
  cat("Steps:", steps, "\n")
  cat("row:", row, "months:", months, "m:", m, "---- d:", d, "\n")
})

##print(interp_df)

interp <- cbind(interp_UT, interp_df, linear_df)

addSheet(tool_workbook, "Interpolation", interp, visibility = FALSE, filter = FALSE)

# selectionTool tab -- uses hidden Dropdowns tab for slicer options
addSheet(tool_workbook, "Dropdowns", names(UT_list), visibility = FALSE, writeColNames = FALSE, filter = FALSE)

addWorksheet(tool_workbook, "selectionTool")

writeData(
  tool_workbook,
  "selectionTool",
  "Select LOB:",
  startCol = 1,
  startRow = 1
)

writeData(tool_workbook, "selectionTool", names(UT_list)[1], startCol = 2, startRow = 1)

dataValidation(
  wb = tool_workbook,
  sheet = "selectionTool",
  cols = 2,
  rows = 1,
  type = "list",
  value = "=dropdowns",
  allowBlank = TRUE,
  showInputMsg = TRUE,
  showErrorMsg = TRUE
)

# Label options
writeData(tool_workbook, "selectionTool",
          x = '=IF(
             INDIRECT($B$1 & "_labels") <> "", 
             INDIRECT($B$1 & "_labels"), 
             "")',
          startCol = 1,
          startRow = 25)

for (col in 1:full_width) {
  writeFormula(
    tool_workbook,
    "selectionTool",
    x = paste0("=INDEX(INDIRECT($B$1), 1, ", col, ")"),
    startCol = col,
    startRow = 3
  )
}

# for (col in 1:(label_width)) {
#   count_unique_col <- col - 1 + full_width + buffer + (label_width) + 2
#   formula <- paste0("= 25 + INDIRECT(\"'\" & $B$1 & \"'!",
#                     int2col(count_unique_col),
#                     "2\")"
#                     )
#   writeFormula(tool_workbook, "selectionTool",
#                x = formula,
#                startCol = col,
#                startRow = 1000)
# }
# LDF match formula
col <- label_width + 1
row <- 4

formula <- paste0("=INDEX(Full!",
                  int2col(col),
                  "$1:",
                  int2col(col),
                  "$",
                  nrow(full_UT),
                  ", MATCH(1,
                      (Full!$A$1:$A$", nrow(full_UT), "=$A4)")

for (label_col in 2:label_width) {
  formula <- paste0(formula,
                    " * (Full!$",
                    int2col(label_col),
                    "$1:$",
                    int2col(label_col),
                    "$",
                    nrow(full_UT),
                    "=$",
                    int2col(label_col),
                    row,
                    ")")
}
formula <- paste0(formula, ", 0))")

writeData(
  tool_workbook,
  "selectionTool",
  x = formula,
  startCol = col,
  startRow = row
)

for (row in 4:23) {  
  for (col in 1:(label_width)) {
    
    col_letter <- int2col(col)
    
    range <- paste0("=$",
                    col_letter,
                    "26:$",
                    col_letter,
                    "100")
    
    dataValidation(
      wb = tool_workbook,
      sheet = "selectionTool",
      cols = col,
      rows = row,
      type = "list",
      value = range,
      allowBlank = TRUE,
      showInputMsg = TRUE,
      showErrorMsg = TRUE
    )
  }
  
  for (col in (label_width + 1):full_width) {
    if ((row != 4 || col != label_width + 1) && (col - label_width)%%3 == 1) {
      writeData(tool_workbook, "selectionTool",
                "Drag",
                startCol = col, startRow = row)
    } else if ((row != 4 || col != label_width + 1) && (col - label_width)%%3 == 2) {
      writeData(tool_workbook, "selectionTool",
                "without",
                startCol = col, startRow = row)
    } else if ((row != 4 || col != label_width + 1) && (col - label_width)%%3 == 0) {
      writeData(tool_workbook, "selectionTool",
                "formatting",
                startCol = col, startRow = row)
    }
  }
}

# Style selectionTool
light_green_fill <- createStyle(fgFill = "#CCFFCC")
light_red_fill <- createStyle(fgFill = "#FF6666")
bold_bottom_border <- createStyle(border = "bottom", borderStyle = "thick")

addStyle(
  tool_workbook,
  sheet = "selectionTool",
  style = LDF_style,
  cols = (label_width + 1):full_width,
  rows = 4:23,
  gridExpand = TRUE
)

addStyle(
  tool_workbook,
  sheet = "selectionTool",
  style = light_green_fill,
  cols = (label_width + 1):(full_width - 1),
  rows = 3,
  gridExpand = FALSE
)

addStyle(
  tool_workbook,
  sheet = "selectionTool",
  style = light_green_fill,
  cols = 1:(label_width),
  rows = 4:23,
  gridExpand = TRUE
)

addStyle(
  tool_workbook,
  sheet = "selectionTool",
  style = light_red_fill,
  cols = full_width,
  rows = 4:23,
  gridExpand = TRUE
)

addStyle(
  tool_workbook,
  sheet = "selectionTool",
  style = light_red_fill,
  cols = full_width,
  rows = 4:23,
  gridExpand = TRUE
)

addStyle(
  tool_workbook,
  sheet = "selectionTool",
  style = bold_bottom_border,
  cols = 1:(label_width),
  rows = 3,
  gridExpand = TRUE
)

addStyle(
  tool_workbook,
  sheet = "selectionTool",
  style = bold_bottom_border,
  cols = 1:(label_width),
  rows = 25,
  gridExpand = TRUE
)


# Save
#Folder
tool_folder_path <- sprintf(
  "G:\\Actuarial\\commercial lines division\\Resources\\ISO\\%s\\ISO LDF Tool %s",
  currentYear, currentYear
)

# Delete/revise path
# Optional: Delete duplicate
if (dir.exists(tool_folder_path)) {
  unlink(tool_folder_path, recursive = TRUE, force = TRUE)
}

# Revise path
original_tool_folder_path <- tool_folder_path
n <- 1
while (dir.exists(tool_folder_path)) {
  tool_folder_path <- paste0(original_tool_folder_path, "(", n, ")")
  n <- n + 1
}

dir.create(tool_folder_path)

#Tool
tool_path <- file.path(tool_folder_path, sprintf("001 - %s ISO LDF tool.xlsx",
                                                 currentYear))

# Delete/revise path
# Optional: Delete duplicate
clear(tool_path)
# Revise path
tool_path <- revise(tool_path)

activeSheet(tool_workbook) <- "selectionTool"
print(saveWorkbook(tool_workbook, file = tool_path, overwrite = TRUE))
##shell.exec(tool_path)


# 5. Documentation
# Readme
readMe_original_path <- "G:\\Actuarial\\commercial lines division\\Resources\\ISO\\ISO LDF Tool Original Documentation\\Read Me.docx"
readMe_copy_path <- file.path(tool_folder_path, "003 - Read Me.docx")
clear(readMe_copy_path)
revise(readMe_copy_path)
file.copy(from = readMe_original_path, to = readMe_copy_path, overwrite = TRUE)

# Assembly instructions
instructions_original_path <- "G:\\Actuarial\\commercial lines division\\Resources\\ISO\\ISO LDF Tool Original Documentation\\Assembly Instructions -- For E&S Team Only.docx"
instructions_copy_path <- file.path(tool_folder_path, "002 - Assembly Instructions - For the E&S Team Only.docx")
clear(instructions_copy_path)
revise(instructions_copy_path)
file.copy(from = instructions_original_path, to = instructions_copy_path, overwrite = TRUE)

shell.exec(tool_folder_path)


# 6. Next year's folder
if (!dir.exists(sprintf("G:\\Actuarial\\commercial lines division\\Resources\\ISO\\%s\\LDF\\Exhibits", nextYear))) {
  dir.create(sprintf("G:\\Actuarial\\commercial lines division\\Resources\\ISO\\%s\\LDF\\Exhibits", nextYear), recursive = TRUE)
}


# 7. Duration log
duration_secs <- as.numeric(difftime(Sys.time(), start_time, units = "secs"))

hours <- floor(duration_secs / 3600)
minutes <- floor((duration_secs %% 3600) / 60)
seconds <- round(duration_secs %% 60)

formatted_duration <- sprintf("%02d:%02d:%02d", hours, minutes, seconds)
cat(
  "End time:",
  format(Sys.time(), "%H:%M:%S"),
  "\n",
  "Duration:",
  formatted_duration,
  "\n"
)
