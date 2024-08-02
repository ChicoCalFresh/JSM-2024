# Import script ----
# This goes in a separate script file

library(qualtRics)
library(tidyverse)
library(janitor)
library(readxl)


# get survey 
qualtrics.data.raw <- fetch_survey(surveyID='XXXXXXXXXXXX', force_request = TRUE) %>% # Update surveyID
  filter(Finished == "TRUE", # filter for participants who completed the survey
         consent == "Yes, I consent to participate") #%>% # filter for participants who consented to the survey


## Split username and email extension into two columns
qualtrics.data.raw$email_name <- str_split(qualtrics.data.raw$campus_email, "@", simplify=TRUE)[,1] %>% str_trim() # Uses @ to split the username/extension

### check for duplicate usernames/responses ----
dups <- qualtrics.data.raw %>% group_by(email_name) %>% 
  tally() %>% # counts the number of responses by username
  filter(n>1) %>% # filters for usernames with more than one response
  left_join(qualtrics.data.raw) # joins survey data back onto the list (to get RecordedDate)

dups_count <- dups %>% select(email_name, RecordedDate) %>% 
  filter(!is.na(email_name)) %>% # filters out blank responses
  group_by(email_name) %>% 
  slice(2) %>% # finds the second (duplicate) response
  mutate(multipleresponses = "drop") # adds a flag to the duplicate responses

qualtrics.data <- qualtrics.data.raw %>% left_join(dups_count) # joins the flagged data back onto the survey data

## Fix student emails -----
# Add lines for other email extensions as needed
qualtrics.data <- qualtrics.data.raw %>% mutate(campus_email=case_when(
  grepl("chic", campus_email, ignore.case=TRUE) ~ paste0(email_name, "@csuchico.edu"), # fixes CSU Chico emails
  TRUE ~ campus_email)) %>% # this is the 'else' condition
  filter(!email_name %in% c("test"), # filter any test responses
         is.na(multipleresponses)) # filter the duplicate responses

# Data export
save("qualtrics.data",
     file = "survey_data_raw.Rdata") # update file name and/or path if needed



# Gift card script ----
# This goes in its own script file
# Includes instructions for sending emails through either Outlook or Gmail 

# Setup ----
library(tidyverse)
library(readxl)
# library(Microsoft365R)  # for sending emails through Outlook
library(gmailr)           # for sending emails through Gmail
library(janitor)

load("survey_data_raw.Rdata")

# Set up connection to Gmail
gm_auth_configure(path=here::here("scripts/research_credentials.json")) # needed to connect with a gmail account

# Clean and select only emails/names/student ids
qualtrics.data <- qualtrics.data %>%
  filter(Finished == "TRUE", # filter for participants who completed the survey
         consent == "Yes, I consent to participate", # filter for participants who consented to the survey
         !campus_email %in% c("testemail@qemailserver.com", "test@gmail.com")) %>% # filter any remaining test responses
  # convert emails to lowercase 
  mutate(campus_email = tolower(campus_email),
         second_email = tolower(second_email),
         # convert names to characters 
         first_name = as.character(first_name),
         last_name = as.character(last_name)) %>%
  select(campus_email:last_name) %>% clean_names() # select only the emails, names, and ID columns

# Load in gift card list ----
## This list contains gift card links and the name/email address of the participant who had been sent each gift card (if applicable)

## Load in most recent Amazon gift card file ----
gc.path.amazon <- "data/gift_cards/Amazon" # update gift card file path

# Find most recently created file
data_files.amazon <- as.data.frame(file.info(list.files(gc.path.amazon, full.names = TRUE))) 
newest.csv.amazon.fullpath <- row.names(data_files.amazon)[which.max(data_files.amazon[["ctime"]])] # finds most recently created file
newest.csv.amazon.filename <- gsub(pattern=('/'), replacement="", 
                                   x=strsplit(newest.csv.amazon.fullpath, "Amazon")[[1]][2], fixed=T)

## Prompt user to double check that correct gift card csv is being read in
print(paste0("Please check that - ", newest.csv.amazon.filename , " - is the most recent csv created."))
cont.script <- ifelse(readline(prompt = "Is this the correct Amazon gift card file (Y/N):") == 'Y', TRUE, FALSE)

##>>>>> -------FIRST SEGMENT ------<<<<<#### 

if(!cont.script) { # if wrong file, then exit script
  rm(list = ls(all.names = TRUE)) # clear workspace
  stop("Issue reading in gift card data, please rerun script or manually import csv file.", call. = FALSE)
}
rm(cont.script, data_files.amazon, newest.csv.amazon.filename)

## Import most recently created Amazon file
gift.card.data <- read.csv(newest.csv.amazon.fullpath)


## Create a list of prior gift card recipients ----
prior.recipients <- gift.card.data %>%
  mutate(campus_email = as.character(Email),
         Recipient.Name = as.character(Recipient.Name)) %>%
  select(Recipient.Name, campus_email)

# Find survey participants who need a gift card ----
need_gc <- anti_join(qualtrics.data, prior.recipients, by = "campus_email")

rm(prior.recipients) # cleanup

# Gift card email setup ----
# add a rownumber index to gc data
gift.card.data <- gift.card.data %>% mutate(rnum = seq.int(nrow(gift.card.data)))
  
gc.available <- gift.card.data %>% 
  filter(is.na(Email)) %>%
  slice(1:nrow(need_gc)) 
  
if(nrow(gc.available) < nrow(need_gc)) { # make sure enough gift cards for eligible students
  rm(list = ls(all.names = TRUE))
  stop("Not enough gift cards for eligble students, stopping script.", call. = FALSE)
}
  
need_gc <- need_gc %>% mutate(gc.link = gc.available$eGift.Redemption.Link)
  
  
# Email message setup - change email text as needed
gc.email.txt <- need_gc %>%
  mutate(sent_gift_card_date = Sys.Date(), 
          msg_body = str_c("Dear ", first_name, ",<p> Thank you for your participation in our survey about CalFresh. 
                          Here is the Amazon Gift Card link as a token of appreciation for your 
                          time and responses: <p>", gc.link, "<p> If you are experiencing issues 
                          opening your link, try opening the link on another device such as your phone/another 
                          computer, or try opening the link in a different browser. Thank you again!
                          <p><br><p> Sincerely,<br> Center for Healthy Communities Research Team")) 


##>>>>> ------- SECOND SEGMENT ------<<<<<#### 

# Sets a limit on the number of emails sent at a time (in this case, 30 at a time)
link.email.idx <- round(seq(0, nrow(gc.email.txt), length=ceiling(nrow(gc.email.txt)/30)+1))

# Gmail email process ----
## Create email drafts----
# Have user check to make sure in correct format, then delete the drafts
for(l in 1:3){ # create a few sample email drafts
  if(!is.na(gc.email.txt$campus_email[l])){
    
    draft_link_email <- gm_mime() %>%
      gm_to(gc.email.txt$campus_email[l]) %>%
      gm_from("sample_email@gmail.com") %>% # update the email address
      gm_subject("Gift Card for CalFresh Food Survey") %>% 
      gm_html_body(gc.email.txt$msg_body[l]) %>%
      gm_attach_file("scripts/gmail_sig.PNG", id="sig")
    
    gm_create_draft(draft_link_email)
  }
}

print("Sample drafts created, please check Gmail to ensure format is correct.")
send.emails <- ifelse(readline(prompt = "Are emails in correct format (Y/N):") == 'Y', TRUE, FALSE)

## Send out emails ----
if(send.emails) {
  print("Sending gift cards to eligible students...")
  for (i in 1:(length(link.email.idx)-1)) {
    start <- link.email.idx[i]+1
    end <- link.email.idx[i+1]
    for(l in start:end){ # only up to 50 at a time
      if(!is.na(gc.email.txt$campus_email[l])){
        
        draft_link_email <- gm_mime() %>%
          gm_to(gc.email.txt$campus_email[l]) %>%
          gm_from("sample_email@gmail.com") %>%   # update the email address
          gm_subject("Gift Card for Survey") %>%  # update subject line
          gm_html_body(gc.email.txt$msg_body[l]) %>%
          gm_attach_file("scripts/gmail_sig.PNG", id="sig")
        
        gm_send_message(draft_link_email)
      }
    }
    Sys.sleep(8) # add short delay in case high number of new students to write
  }
  rm(end, start, link.email.idx, i)
} else {
  stop("Ending script, please step through manually to see where draft error occured.", call. = FALSE)
}

rm(send.emails)



# Outlook email process ----
## Create email drafts ----
# outl <- get_business_outlook(shared_mbox_email = "sample_email@csuchico.edu") # update email address
#
# for(l in 1:3){ # create a few sample email drafts
#   if(!is.na(gc.email.txt$campus_email[l])){
#     em2 <- outl$
#       create_email(content_type="html")$
#       set_body(gc.email.txt$msg_body[l])$
#       set_subject("Gift Card Link")$
#       set_recipients(to=gc.email.txt$campus_email[l], cc=gc.email.txt$second_email[l])$
#       add_attachment("scripts/gmail_sig.PNG")
#     
#   }
# }
# rm(l, draft_link_email)
# 
# print("Sample drafts created, please check Outlook to ensure format is correct.")
# send.emails <- ifelse(readline(prompt = "Are emails in correct format (Y/N):") == 'Y', TRUE, FALSE)
# 
# 
# ## Send out emails ----
# if(send.emails) {
#   print("Sending gift cards to eligible students...")
#   for (i in 1:(length(link.email.idx)-1)) {
#     start <- link.email.idx[i]+1
#     end <- link.email.idx[i+1]
#     for(l in start:end){ # only up to 50 at a time
#       if(!is.na(gc.email.txt$email[l])){
# 
#         draft_link_email <- gm_mime() %>%
#           gm_to(gc.email.txt$email[l]) %>%
#           gm_from("research.for.chc@gmail.com") %>%
#           gm_subject("Gift Card Link - Center for Healthy Communities") %>%
#           gm_html_body(gc.email.txt$msg_body[l]) %>%
#           gm_attach_file("scripts/gmail_sig.PNG", id="sig")
# 
#         gm_send_message(draft_link_email)
#       }
#     }
#     Sys.sleep(8) # add short delay in case high number of new students to write
#   }
#   rm(end, start, link.email.idx, i)
# } else {
#   stop("Ending script, please step through manually to see where draft error occured.", call. = FALSE)
# }
# rm(send.emails)


# Update gift card file ----
## check with your grant and/or research office for any gift card tracking requirements
gc.data.tmp <- gc.available %>% mutate(Recipient.Name = paste(need_gc$first_name, need_gc$last_name),
                                       Email = need_gc$campus_email,
                                       Distributed.Date = format(Sys.Date(), "%m-%d-%Y"))
  
  tmp.idx <- match(gc.data.tmp$rnum, gift.card.data$rnum)
  
  gift.card.data[tmp.idx,] <- gc.data.tmp
  gift.card.data <- gift.card.data %>% select(-rnum)
  
  write.csv(gift.card.data,
            paste0(gc.path.amazon, "Gift cards as of ", format(Sys.time(), "%Y-%m-%d-%H%M"), ".csv"),
            row.names = FALSE)  