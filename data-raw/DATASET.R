## code to prepare `DATASET` dataset goes here

library(haven)

sad_code <- read_sas("Q:/50-radgivning/02-mynd/SAS Library/Arter/art.sas7bdat",
                     NULL)
species <- sad_code[-1, c("start", "art", "engkode")]

more_species <- data.frame(
  start = c("USO", "MZZ"),
  art = c("Usorterbart", "Uspecificeret"),
  engkode = c("USO", "MZZ")
)

species  <- rbind(species , more_species)

usethis::use_data(species, overwrite = TRUE)


