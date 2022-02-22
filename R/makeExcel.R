#' Output a set of data tables to an accessible xlsx file
#'
#' This function outputs a .xlsx file made up of a series of R data frames.
#'
#' @param tables A list of data frames. Each list element should be a data.frame.
#' @param titles A list of table titles. Each list element should be a character vector of length 1.
#' @param subtitles A list of subtitles. Each list element should be a character vector of any length.
#' @param tabnames A list of worksheet tab names. Each list element should be a character vector of length 1.
#' @param tablenames A list of table names. Each list element should be a character vector of length 1.
#' @param filename A string specifying the name of the output file.
#' @return A .xlsx workbook
#' @export

makeExcel <- function(tables,titles,subtitles,tabnames,tablenames,filename) {

  headerStyle <- openxlsx::createStyle(fontName = "Arial",
                                       fontSize = 12,
                                       textDecoration = "BOLD")

  titleStyle <- openxlsx::createStyle(fontName = "Arial",
                                      fontSize = 14,
                                      textDecoration = "BOLD")

  wb <- openxlsx::createWorkbook()

  openxlsx::modifyBaseFont(wb,
                           fontSize = 12,
                           fontName = "Arial")

  for (i in 1:length(tables)) {

    tabrows <- nrow(tables[[i]])
    tabcols <- ncol(tables[[i]])

    titledf <- data.frame(subtitles[[i]])
    names(titledf) <- titles[[i]]

    openxlsx::addWorksheet(wb,
                           tabnames[[i]])

    openxlsx::writeDataTable(wb,
                             sheet = tabnames[[i]],
                             x = data.frame(tables[[i]]),
                             startRow = length(subtitles[[i]]) + 2,
                             tableStyle = "none",
                             headerStyle = headerStyle,
                             withFilter = FALSE,
                             bandedRows = FALSE,
                             tableName = names(tables[i]))

    openxlsx::setColWidths(wb,
                           sheet = tabnames[[i]],
                           cols = 1:tabcols,
                           widths = "auto")

    openxlsx::writeData(wb,
                        sheet =  tabnames[[i]],
                        x = titledf,
                        headerStyle = titleStyle)

  }

    openxlsx::saveWorkbook(wb,filename,overwrite=TRUE)

}
