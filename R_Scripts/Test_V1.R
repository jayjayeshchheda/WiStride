# loading the packages required to run the code
library(readxl)
library(tidyverse)
library(datasets)

# loading the excel file along with the range of cells to read from the file.
Trial_Work <- read_excel("Trial work.xlsx",
                         range = "A1:S39")
head(Trial_Work) # code to show first 6 rows worth of data in the excel file.

# Method of loading selected rows and filtering them
Trial_Work %>% 
  select(Molecule_List, Country, Product) %>% 
  filter(Molecule_List == "EMPAGLIFLOZIN", Product == "COSPIAQ")

data() # code to view all the data sets present in the library(datasets)

view(starwars) # code to view the data set starwars
# Here the selected data from the data set starwars is assigned to variable sw
sw <- starwars %>% 
  select(gender, mass, height, species) %>% 
  filter(species=="Human") %>% 
  na.omit() %>% # code to remove na values from the data
  mutate(height = height/100) %>% # code to change the value within the column height
  mutate(BMI = mass/height^2) # code to add a new column BMI in the sw variable
head(sw)

plot(sw$mass, sw$BMI) #code to plot sw
plot(sw)


# Method to install multiple packages at once from default R library (CRAN)
install.packages(c("devtools", "lme4"))
install.packages("KernSmooth")
library(KernSmooth)
# Method to install packages from bioconductor
source("https://bioconductor.org/biocLite.R")
biocLite()
biocLite("GenomicFeatures")

# Method to install packages from github
library(devtools)
install_github("author/package")

update.packages() #To update all packages

help(package = "devtools") #loads documentation for a package

browseVignettes("ggplot2") #To see the vignettes included in a package
