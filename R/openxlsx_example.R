# make accessible spreadsheets with the openXLSX package

# Data wrangling ---------------------------------------------------------------

# load tidyverse package for data wrangling (ignore 'Warning message: package 
# ‘tidyverse’ was built under R version 4.0.3 )' - this may pop up because you 
# have R 3.6 installed)
library(tidyverse)

# look at the built-in example dataset 'iris'
head(iris)

# and its documentation
?iris

# The iris dataset gives the measurements in centimeters of the variables sepal 
# length and width and petal length and width, respectively, for 50 flowers from 
# each of 3 species of iris. The species are Iris setosa, versicolor, and 
# virginica.

# create table

table <- iris %>% 
  group_by(Species) %>% 
  summarise(Sample = n(),
            Petal_length_mean = mean(Petal.Length),
            Sepal_length_mean = mean(Sepal.Length)) %>% 
  mutate(Ratio_S_P = Sepal_length_mean/Petal_length_mean,
         Ratio_S_P = round(Ratio_S_P, 2),
         Petal_length_mean = round(Petal_length_mean, 1),
         Sepal_length_mean = round(Sepal_length_mean, 1) ) 

# inspect


table


# load openxlsx package for creating spreadsheets (ignore 'Warning message: 
# package ‘openxlsx’ was built under R version 4.0.3 )' - this may pop up 
# because you have R 3.6 installed)
library(openxlsx)

# Quick way --------------------------------------------------------------------

# quick way of creating a spreadsheet - little customisation available, no title
write.xlsx(table, "output/iris_table_quick.xlsx", asTable = TRUE)


# Long way ---------------------------------------------------------------------

# verbose but highly customisable way of creating a spreadsheet

wb <- createWorkbook()
addWorksheet(wb, "Iris_table")

# add title
writeData(wb, "Iris_table", "Mean petal and sepal lengths by iris species")

# add dataset
writeDataTable(wb, "Iris_table", table, startRow = 2)

# Save as xlsx in output folder
saveWorkbook(wb, 'output/iris_table_long.xlsx', overwrite = TRUE)


# Formatting -------------------------------------------------------------------

wb <- createWorkbook()

# add worksheet with no grid lines
addWorksheet(wb, "Iris_table", gridLines = FALSE)

# Add title and make it bold
writeData(wb, "Iris_table", "Mean petal and sepal lengths by iris species")
addStyle(wb, "Iris_table", createStyle(textDecoration = "bold"), rows = 1, cols = 1)

# Add source and make it italic
writeData(wb, "Iris_table", "Source: Base R built-in 'iris' dataset",  startRow = 2)
addStyle(wb, "Iris_table", createStyle(textDecoration = "italic"), rows = 2, cols = 1)

# Add table and simplify style with bold headers
writeDataTable(wb, "Iris_table", table, startRow = 3, 
               tableStyle = "none",
               withFilter = FALSE,
               headerStyle = createStyle(textDecoration = "bold"))

# format ratio column as percentage
addStyle(wb, "Iris_table", 
         rows = 4:(dim(table)[1] + 3), cols = 5, 
         style = createStyle(numFmt = "0%", halign = "right"))

# right-align headers of numeric columns
addStyle(wb, "Iris_table",
         rows = 3, cols = 2:5,
         style = createStyle(halign = "right", textDecoration = "bold"))

# adjust column width
setColWidths(wb, "Iris_table", cols = 2:ncol(table), widths = "auto")

# adjust row height
setRowHeights(wb, "Iris_table", heights = 25, rows = 3)

# Save as xlsx in output folder
saveWorkbook(wb, 'output/iris_table_long_formatted.xlsx', overwrite = TRUE)



# Add Readme worksheet ---------------------------------------------------------

wb <- createWorkbook()

# Readme worksheet

addWorksheet(wb, "Readme", gridLines = FALSE)
writeData(wb, "Readme", 
          c("Readme",
            "This document contains a simple example of an accessible table.",
            "There isn't a lot more to say",
            "Add as many lines as you want."))
# format title
addStyle(wb, "Readme", rows = 1, cols = 1, createStyle(textDecoration = "bold"))



# Iris_table worksheet

# add worksheet with no grid lines
addWorksheet(wb, "Iris_table", gridLines = FALSE)

# Add title and make it bold
writeData(wb, "Iris_table", "Mean petal and sepal lengths by iris species")
addStyle(wb, "Iris_table", createStyle(textDecoration = "bold"), rows = 1, cols = 1)

# Add source and make it italic
writeData(wb, "Iris_table", "Source: Base R built-in 'iris' dataset",  startRow = 2)
addStyle(wb, "Iris_table", createStyle(textDecoration = "italic"), rows = 2, cols = 1)

# Add table and simplify style with bold headers
writeDataTable(wb, "Iris_table", table, startRow = 3, 
               tableStyle = "TableStyleLight3",
               withFilter = FALSE,
               headerStyle = createStyle(textDecoration = "bold"))

# format ratio column as percentage
addStyle(wb, "Iris_table", 
         rows = 4:(dim(table)[1] + 3), cols = 5, 
         style = createStyle(numFmt = "0%", halign = "right"))

# right-align headers of numeric columns
addStyle(wb, "Iris_table",
         rows = 3, cols = 2:5,
         style = createStyle(halign = "right", textDecoration = "bold"))

# adjust column width
setColWidths(wb, "Iris_table", cols = 2:ncol(table), widths = "auto")

# adjust row height
setRowHeights(wb, "Iris_table", heights = 25, rows = 3)

# Save as xlsx in output folder
saveWorkbook(wb, 'output/iris_table_long_formatted_with_readme.xlsx', overwrite = TRUE)





