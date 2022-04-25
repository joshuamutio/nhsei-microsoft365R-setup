# Instructions for connecting R to NHSEI Microsoft 365
# See also the vignettes here: https://cran.r-project.org/web/packages/Microsoft365R/index.html

# Install

install.packages('Microsoft365R')

# Load

library(Microsoft365R)

# Invoke connection - this will open a browser window for authenticating
# Login with your @england.nhs.uk username and password
# If encountering issues authenticating, make sure the default browser is Chrome or similar

list_sharepoint_sites()

# Connect with the site name as retrieved in the previous step

site <- get_sharepoint_site('Example Sharepoint Site Name')

# Alternatively connect using site url

site <- get_sharepoint_site(site_url = 'https://nhsengland.sharepoint.com/examplesharepointsiteurl')

# Writing a csv e.g. to a Restricted Library in the Collaboration Drive

site$list_drives()
reslib <- site$get_drive('Restricted Library')
write.csv(mtcars, file.path(tempdir(), "mtcars"))
mtcars_csv <- file.path(tempdir(), "mtcars.csv")
reslib$upload_file(mtcars_csv, dest = "mtcars.csv")

# Posting a message to a Teams channel

list_teams()
team <- get_team("Example Team")
channel <- team$get_channel("General")
channel$send_message("Hello world")

