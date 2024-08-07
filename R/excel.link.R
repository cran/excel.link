#' excel.link: convenient data exchange with Microsoft Excel
#' 
#' The excel.link package mainly consists of two rather independent parts: one
#' is for transferring data/graphics to running instance of Excel, another part
#' - work with data table in Excel in similar way as with usual data.frame.
#' 
#' @section Transferring data: Package provided family of objects: 
#'   \code{\link{xl}}, \code{\link{xlc}}, \code{\link{xlr}} and 
#'   \code{\link{xlrc}}. You don't need to initialize these objects or to do any
#'   other preliminary actions. Just after execution \code{library(excel.link)} 
#'   you can transfer data to Excel active sheet by simple assignment, for 
#'   example: \code{xlrc[a1]  = iris}. In this notation 'iris' dataset will be 
#'   written with column and row names. If you doesn't need column/row names 
#'   just remove 'r'/'c' letters (\code{xlc[a1]  = iris} - with column names but
#'   without row names). To read Excel data just type something like this: 
#'   \code{xl[a1:b5]}. You will get data.frame with values from range a1:a5 
#'   without column and row names. It is possible to use named ranges (e. g. 
#'   \code{xl[MyNamedRange]}). To transfer graphics use \code{xl[a1] = 
#'   current.graphics()}.
#'   You can make active binding to Excel range:
#'   \preformatted{
#'   xl.workbook.add()
#'   xl_iris \%=crc\% a1 # bind variable to current region around cell A1 on Excel active sheet
#'   xl_iris = iris # put iris data set 
#'   identical(xl_iris$Sepal.Width, iris$Sepal.Width)
#'   xl_iris$test = "Hello, world!" # add new column on Excel sheet
#'   xl_iris = within(xl_iris, {
#'      new_col = Sepal.Width * Sepal.Length # add new column on Excel sheet
#'      }) 
#'   }
#'   
#' @section Live connection: For example we put iris datasset to Excel sheet:
#'   \code{xlc[a1] = iris}. After that we connect Excel range with R object:
#'   \code{xl_iris = xl.connect.table("a1",row.names = FALSE, col.names =
#'   TRUE)}. So we can: 
#'   \itemize{ 
#'   \item get data from this Excel range: \code{xl_iris$Species} 
#'   \item add new data to this Excel range: \code{xl_iris$new_column = 42} 
#'   \item sort this range: \code{sort(xl_iris,column = "Sepal.Length")} 
#'   \item and more...
#'   }
#'   Live connection is faster than active binding to range but is less universal 
#'   (for example, you can't use \code{within} statement with it).
#' @seealso \code{\link{xl}}, \code{\link{current.graphics}},
#'   \code{\link{xl.connect.table}}
"_PACKAGE"


#' @useDynLib "excel.link",.registration = TRUE
#' @import methods grDevices utils
NULL



