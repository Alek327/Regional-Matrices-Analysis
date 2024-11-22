# Regional Matrices Analysis (MUST4Water Project)

# Overview

This repository contains R code developed as part of the MUST4Water project, focused on regional economic analysis and water management. The code automates the process of extracting, analyzing, and saving sectoral trade matrices from MRIO (Multi-Regional Input-Output) data. The matrices are organized by region and exported to a single Excel file for further analysis.

# Features

Efficient extraction and transformation of MRIO data into 43x6 matrices for each region.
Automated export of the extracted matrices into a single Excel file with separate sheets for each region.
Designed to support regional socio-economic analysis within the MUST4Water project.
Installation

Clone this repository:
bash
Copy code
git clone https://github.com/Alek327/Regional-Matrices-Analysis.git
Ensure that you have R installed on your system. If not, download it from CRAN.
Install the necessary R packages by running:
r
Copy code
install.packages("openxlsx")
Usage

Open the R script in your preferred R environment (e.g., RStudio).
Adjust the file paths in the script as needed, depending on the location of your input MRIO data files.
Run the script to extract the matrices and save them to an Excel file.
The output Excel file will contain 6 sheets corresponding to 6 regions: NW, NE, Toscana, Centro, Sud, and Ext. Each sheet will have a 43x6 matrix representing sectoral trade data.
Example Code
Here’s a snippet of the main code:

r
Copy code
# Extract 6 matrices of size 43x6
extracted_matrices <- list()

for (i in 1:6) {
  extracted_matrices[[i]] <- t(mrio6reg[[2]][i, , ])
}

names(extracted_matrices) <- c("NW", "NE", "Toscana", "Centro", "Sud", "Ext")

# Save matrices to Excel
library(openxlsx)
wb <- createWorkbook()

for (i in 1:length(extracted_matrices)) {
  addWorksheet(wb, names(extracted_matrices)[i])
  writeData(wb, sheet = names(extracted_matrices)[i], extracted_matrices[[i]])
}

saveWorkbook(wb, file = "regions_matrices.xlsx", overwrite = TRUE)
Output

A single Excel file named regions_matrices.xlsx with 6 sheets, each representing a region:
NW
NE
Toscana
Centro
Sud
Ext
Contributing

Contributions to improve this code are welcome! Feel free to submit a pull request or open an issue.

# License

This project is licensed under the GNU General Public License – see the LICENSE file for details.

# Contact

For questions or further information, please contact:

Oleksandr Galychyn
ogalychyn@yahoo.com
