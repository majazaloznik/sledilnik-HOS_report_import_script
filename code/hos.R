###############################################################################
###############################################################################
#'
#' INPUTS
#' 1. download new HOS report from google drive
#'
#' checks:
#'     - file name with today's date
#'     - if multiple files are found with the same name, one of them will be
#'       downloaded, but cannot pick which one... You'll get a warning tho :)
#'     * no checks on column names (variables) or row names (hospitals)
#' 
#' 2. from the target googleSheet get the rownumber for the new appended row
#' 
#' checks:
#'     - that correct number of columns is present
#'     * no ckecks that of colum names (variables)
#' 
#' MAIN
#' 3. write vector for new appended row which combines
#'     - excel formulas with correct cell references
#'     - data from HOS tables
#'     
#' checks:
#'     - check if any of the H.i fields calculated are negative. 
#'     
#' OUTPUTS:
#' 4. append vector as new row in GoogleSheet
#' 
#' CLEANUP:
#' 5. delete local file
#' 
###############################################################################
###############################################################################



###############################################################################
## pre-PRELIMINARIES - only run first time
###############################################################################
# install.packages("googledrive")
# install.packages("readxl")
# install.packages("googlesheets4")
drive_auth()
gs4_auth(token = drive_token())

###############################################################################
## PRELIMINARIES
###############################################################################
library(googledrive)
library(readxl)
library(googlesheets4)


###############################################################################
## IMPORT AND PREPARE HOS REPORT
###############################################################################

# get filename
today <- as.Date("2021-01-07") # Sys.Date()
today <- gsub("\\b0(\\d)\\b", "\\1", format(today, "%d.%m.%Y"))
fname <- paste0("Hospitalizacije COVID-19_", today, ".xlsx")

# check file exists and download it
file.info <- drive_get(fname)

if (nrow(file.info) == 0) {
  print("File does not exist on Gdrive. Script aborted.") } else {
    
    file.id <- file.info$id[1]
    file.name <- file.info$name[1]
    file.path <- paste0("data/", file.name)
    
    # download file
    drive_download(as_id(file.id), path = file.path, overwrite = TRUE)
    if (nrow(file.info) > 1) {
      print(paste("Multiple files with the name", fname,
                   "were found. One of them was downloaded, but I can't tell", 
                   "you which one. They are probably the same anyway :shrug:"))}
  }

# define cell ranges to read from the xlsx file
hos.range <- "A2:O15"
care.range <- "A19:K21"
psi.range <- "A25:N30"

hos <- read_xlsx(file.path, range = hos.range)
care <- read_xlsx(file.path, range = care.range)
psi <- read_xlsx(file.path, range = psi.range)


###############################################################################
## READ NEW ROW NUMBER FROM GSHEET
###############################################################################

rowz <- range_read(as_id("1kiXh4SUnFqp_Xe6gU6Be-Mrob4bkq7jUOohIbKRt7eM"),
            sheet = "HOS",
            range = "A:A",
            col_names = FALSE)

N <- nrow(rowz)














