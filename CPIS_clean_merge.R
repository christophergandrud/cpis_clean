################################################################################
# Clean IMF CPIS data
# Christopher Gandrud
# 23 December 2014
################################################################################

# Set working directory
setwd('~/Desktop/CPIS_Raw/')

# Load required packages
library(xlsx)
library(dplyr)
library(reshape2)
library(countrycode)

files <- list.files()

# files = c('CPIS_2012.xls', 'CPIS_2013.xls')

for (i in files){
    # Extract the data's year
    year <- i %>% gsub('CPIS_', '', .) %>% gsub('.xls', '', .) %>% as.numeric()
    message(year)

    # Extract sheet names and clean
    sheet_names_dirty <- loadWorkbook(i) %>% getSheets() %>% names()

    # Extract sheets and clean
    for (u in sheet_names_dirty){
        temp <- read.xlsx(i, u, stringsAsFactors = FALSE)

        if (u == 'Sheet1'){
            sheet_name_clean <- 'LongTermDebtSecurities'
        }
        else if (u == 'Sheet2'){
            sheet_name_clean <- 'ShortTermDebtSecurities'
        }
        else {
            sheet_name_clean <- u %>%
                                 gsub('Table 11.*-', '', .) %>%
                                 gsub(' ', '', .)
        }

        # Clean sheet
        message(paste('    ', sheet_name_clean))
        temp_sub <- temp %>% slice(5:248) %>% select(-1, -(3:4))

        # Use iso2c codes
        country_from <- as.character(temp_sub[1, -1]) %>% countrycode(
                                                        origin = 'country.name',
                                                        destination = 'iso2c')
        names(temp_sub) <- c('country_to', country_from)
        temp_sub$country_to <- countrycode(temp_sub$country_to,
                                           origin = 'country.name',
                                           destination = 'iso2c')

        temp_sub <- temp_sub[, 1:ncol(temp_sub) - 1]
        temp_sub <- rename(temp_sub, IntOrg = NA.1)

        temp_sub <- temp_sub %>% slice(-1) %>%
                    melt(id.vars = 'country_to')
        names(temp_sub) <- c('country_to', 'country_from', sheet_name_clean)
        temp_sub$country_from <- as.character(temp_sub$country_from)

        # Standardise missing
        temp_sub[, sheet_name_clean] <- temp_sub[, sheet_name_clean] %>%
                                        gsub('-', '0', . ) %>%
                                        gsub('...', '', .) %>% # ... missing
                                        gsub('c', '', . ) %>% # c = confidential
                                        as.numeric()
        # Remove NA countries
        temp_sub <- temp_sub %>% subset(!is.na(country_to)) %>% 
                    subset(!is.na(country_from))
        temp_sub$year <- year

        # Merge together within year
        if (grep(u, sheet_names_dirty) == 1){
            temp_merged <- temp_sub
        }
        else if (grep(u, sheet_names_dirty) != 1){
            temp_merged <- left_join(temp_merged, temp_sub,
                                by = c('country_to', 'country_from', 'year'))
        }
    }
    if (grep(i, files) == 1){
        final_merged <- temp_merged
    }
    else if (grep(i, files) != 1){
        final_merged <- rbind(final_merged, temp_merged)
    }
}


write.csv(final_merged, file = '~/Desktop/CPIS_dyadic_data.csv', 
          row.names = FALSE)