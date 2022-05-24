# Instructions for connecting R to NHSEI Microsoft 365
# See also the vignettes here: https://cran.r-project.org/web/packages/Microsoft365R/index.html
# This uses the here and tidyverse packages for ease, but not necessary

library(here)
library(tidyverse)

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

reslib <- site$get_drive('Restricted Library')
write.csv(mtcars, here('mtcars_upload.csv'))
reslib$upload_file(here('mtcars_upload.csv'), dest = "mtcars_upload.csv")

# Reading the csv back

reslib$download_file('mtcars_upload.csv', dest = here('mtcars_download.csv'))

# Posting a message to a Teams channel

list_teams()
team <- get_team("Example Team")
channel <- team$get_channel("General")
channel$send_message("Hello world")

# Posting to Teams chat
# Chat metadata is nested; the following code will extract and combine the relevant properties for ease

chats <- list_chats()

chats_properties <-
  lapply(chats, function(x) {
    x$properties %>% unlist() %>% as.data.frame() %>% t() %>% as.tibble()
  }) %>% bind_rows() %>%
  relocate(topic)

chats_properties$createdDateTime <- strptime(chats_properties$createdDateTime, "%Y-%m-%dT%H:%M:%OSZ")
chats_properties$lastUpdatedDateTime <- strptime(chats_properties$lastUpdatedDateTime, "%Y-%m-%dT%H:%M:%OSZ")
chats_properties$viewpoint.lastMessageReadDateTime <- strptime(chats_properties$viewpoint.lastMessageReadDateTime, "%Y-%m-%dT%H:%M:%OSZ")

chats_properties

chat <- get_chat("ExampleChatID")
chat$list_members()
chat$send_message("Hello world")
