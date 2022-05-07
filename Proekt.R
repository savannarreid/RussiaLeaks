
library(utf8)
Sys.setlocale("LC_ALL", "Russian")
library(tidyverse)
library(readr)
library(dplyr)
library(stringr)
library(tidyr)
library(purrr)
library(data.table)
library(openxlsx)
### To import and parse emails from a folder of .txt files ###



training_folder <- "C:/Users/Savanna/Desktop/Ros_proekt/sample_emails"

# Define a function to read all files from a folder into a data frame
read_folder <- function(infolder) {
  tibble(file = dir(infolder, full.names = TRUE)) %>%
    mutate(text = map(file, read_lines)) %>%
    transmute(id = basename(file), text) %>%
    unnest(text)
}

raw_text <- tibble(folder = dir(training_folder, full.names = TRUE)) %>%
  mutate(folder_out = map(folder, read_folder)) %>%
  unnest(cols = c(folder_out)) %>%
  transmute(newsgroup = basename(folder), id, text)

### First extract data from each row of the header (rows 1-7):

text_df <- raw_text %>%
  group_by(id) %>% 
  mutate(From = first(text))

### Remove outlook business cards

text_df <- text_df %>%
  group_by(id) %>%
  filter(!any(str_detect(text, "Birthday")))

### Create columns for header fields

body <- text_df %>%
  group_by(id) %>%
  filter(!any(str_detect(text, 
                         "Birthday")),
         !nchar(text) == 0,
         !str_detect(text, 
                     "^SentOn"),
         !str_detect(text,
                     "^From"),
         !str_detect(text,
                     "^To"),
         !str_detect(text,
                     "^Subject"),
         !str_detect(text,
                     "Importance"),
         !str_detect(text,
                     "Attachments"))

body_condensed <- body %>%
  group_by(id) %>%
  mutate(Bodytext = paste0(text, " ", collapse = "")) %>%
  select(-text) %>%
  distinct()

sent <- text_df %>% group_by(id) %>% slice(2)
to <- text_df  %>% group_by(id) %>% slice(3)
subject <- text_df  %>% group_by(id) %>% slice(4)
importance <- text_df  %>% group_by(id) %>% slice(5)
attachments <- text_df  %>% group_by(id) %>% slice(6)

text_df_short <- subset(text_df, select = -c(text, newsgroup))
sent_short <- subset(sent, select = -c(newsgroup, From))
to_short <- subset(to, select = -c(newsgroup, From))
subject_short <- subset(subject, select = -c(newsgroup, From))
importance_short <- subset(importance, select = -c(newsgroup, From))
attachments_short <- subset(attachments, select = -c(newsgroup, From))
body_short <- subset(body_condensed, select = -c(newsgroup, From))
  
new_df <- text_df_short %>%
  inner_join(., sent_short,
             by = c("id" = "id")) %>%
  inner_join(., to,
             by = c("id" = "id")) %>%
  inner_join(., subject_short,
             by = c("id" = "id")) %>%
  inner_join(., importance_short,
             by = c("id" = "id")) %>%
  inner_join(., attachments_short,
             by = c("id" = "id")) %>%
  inner_join(., body_short,
             by = c("id" = "id"))

new_df2 <- subset(new_df, select = -c(From.y, newsgroup))
names(new_df2)[names(new_df2) == 'From.x'] <- "From"
names(new_df2)[names(new_df2) == 'text.x'] <- "Sent_on_Date"
names(new_df2)[names(new_df2) == 'text.y'] <- "To"
names(new_df2)[names(new_df2) == 'text.x.x'] <- "Subject"
names(new_df2)[names(new_df2) == 'text.y.y'] <- "Importance"
names(new_df2)[names(new_df2) == 'text'] <- "Attachments"
new_df2

new_df2$From <- new_df2$From %>%
  str_replace("From: ", "") %>% str_trim()
new_df2$Sent_on_Date <- new_df2$Sent_on_Date %>%
  str_replace("SentOn: ", "") %>% str_trim()
new_df2$Sent_on_Date <- new_df2$Sent_on_Date %>%
  str_replace("SentOn: ", "") %>% str_trim()
new_df2$To <- new_df2$To %>%
  str_replace("To: ", "") %>% str_trim()
new_df2$Subject <- new_df2$Subject %>%
  str_replace("Subject: ", "") %>% str_trim()
new_df2$Importance <- new_df2$Importance %>%
  str_replace("Importance: ", "") %>% str_trim()
new_df2$Attachments <- new_df2$Attachments %>%
  str_replace("Attachments: ", "") %>% str_trim()
new_df2$Bodytext <- new_df2$Bodytext %>% str_trim()

ros_proekt_df <- new_df2[!duplicated(new_df2), ]
ros_proekt_df

write.xlsx(ros_proekt_df, "ros_proekt_df.xlsx")


### Filtering by keyword ####


ru_finance <- read.xlsx("C:/Users/Savanna/Desktop/Code_projects/R_projects/Text_mining_with_R/russian_finance_shorter.xlsx", colNames = TRUE)
finance_vector <- ru_finance$value
finance <- str_trim(finance_vector, side = "both")
filter_finance <- str_c(finance, collapse = "|")

proekt_finance_df <- ros_proekt_df %>% filter(str_detect(Bodytext, regex(filter_finance, ignore_case = FALSE)))
write.xlsx(proekt_finance_df, "proekt_finance.xlsx")


ru_corrupt <- read.xlsx("C:/Users/Savanna/Desktop/Code_projects/R_projects/Text_mining_with_R/ru_corruption.xlsx", colNames = TRUE)
corrupt_vector <- ru_corrupt$value
corruption <- str_trim(corrupt_vector, side = "both")
filter_corrupt <- str_c(corruption, collapse = "|")

proekt_corruption_df <- ros_proekt_df %>% filter(str_detect(Bodytext, regex(filter_corrupt, ignore_case = FALSE)))
write.xlsx(proekt_corruption_df, "proekt_corruption.xlsx")

proekt_corruption_df2 <- ros_proekt_df %>% filter(str_detect(To, regex("ekaterina.kukarskaya@mail.ru", ignore_case = TRUE)))
write.xlsx(proekt_corruption_df2, "proekt_corruption2.xlsx")

german_names <- read.xlsx("C:/Users/Savanna/Desktop/Code_projects/R_projects/Text_mining_with_R/ru_german_names.xlsx", colNames = TRUE)
german_name <- german_names$Name
filter_german_names <- str_c(german_name, collapse = "|")

proekt_german_names <- ros_proekt_df %>% filter(str_detect(Bodytext, regex(filter_german_names, ignore_case = TRUE)))
write.xlsx(proekt_german_names, "proekt_german_names.xlsx")


ru_ukraine <- read.xlsx("C:/Users/Savanna/Desktop/Code_projects/R_projects/Text_mining_with_R/ru_ukraine.xlsx", colNames = TRUE)
ukraine_vector <- ru_ukraine$value
ukraine <- str_trim(ukraine_vector, side = "both")
filter_ukraine <- str_c(ukraine, collapse = "|")

proekt_ukraine <- ros_proekt_df %>% filter(str_detect(Bodytext, regex(filter_ukraine, ignore_case = TRUE)))
write.xlsx(proekt_ukraine, "proekt_ukraine.xlsx")


ru_fsb_malware <- read.xlsx("C:/Users/Savanna/Desktop/Code_projects/R_projects/Text_mining_with_R/ru_fsb_malware.xlsx", colNames = TRUE)
fsb_malware_vector <- ru_fsb_malware$value
fsb_malware <- str_trim(fsb_malware_vector, side = "both")
filter_fsb_malware <- str_c(fsb_malware, collapse = "|")

proekt_fsb_malware <- ros_proekt_df %>% filter(str_detect(Bodytext, regex(filter_fsb_malware, ignore_case = TRUE)))



proekt_ibans <- ros_proekt_df %>% filter(str_detect(Bodytext, regex("[A-Z]{2}[0-9]{2} ?[0-9]{4} ?[0-9]{4} ?[0-9]{4} ?[0-9]{4} ?[0-9]{0,2}|IBAN", ignore_case = TRUE)))

politicians <- read.xlsx("C:/Users/Savanna/Desktop/Code_projects/R_projects/Text_mining_with_R/American_officials.xlsx", colNames = TRUE)
politicians_vector <- politicians$Name
politicians <- str_trim(politicians_vector, side = "both")
filter_politicians <- str_c(politicians, collapse = "|")

proekt_politicians <- ros_proekt_df %>% filter(str_detect(Bodytext, regex(filter_politicians, ignore_case = TRUE)))

proekt_de_to <- ros_proekt_df %>% filter(str_detect(To, regex(".+@.+de$", ignore_case = TRUE)))
proekt_de_from <- ros_proekt_df %>% filter(str_detect(From, regex(".+@.+de$", ignore_case = TRUE)))

### Tokenizing and visualizing sentiment and bigrams

ru_stopwords <- read.xlsx("C:/Users/Savanna/Desktop/Code_projects/R_projects/Text_mining_with_R/ru_stopwords.xlsx", colNames = TRUE)
ru_sentiment <- read.xlsx("C:/Users/Savanna/Desktop/Code_projects/R_projects/Text_mining_with_R/ru_sentiment.xlsx", colNames = TRUE)


ros_proekt_df <- as.data.frame(ros_proekt_df)
proekt_bigrams <- ros_proekt_df %>% 
  unnest_tokens(bigram, Bodytext, token = "ngrams", n = 2)

proekt_bigrams_sep <- proekt_bigrams %>% 
  separate(bigram, c("word1", "word2"), sep = " ")
proekt_keybigrams <- proekt_bigrams_sep %>% 
  anti_join(ru_stopwords, c("word1" = "words")) %>% 
  anti_join(ru_stopwords, c("word2" = "words"))
proekt_bigrams_key <- proekt_keybigrams %>% 
  unite(bigram, word1, word2, sep = " ")

proekt_bigram_counts <- proekt_bigrams %>% 
  count(From, bigram, sort = TRUE)

proekt_bigram_dtm <- proekt_bigram_counts %>% 
  cast_dtm(From, bigram, n)

proekt_bigram_lda <- LDA(
  proekt_bigram_dtm, k=6, control = list(seed = 1234))

proekt_bigram_topics <- tidy(proekt_bigram_lda, matrix = "beta")

top_bigrams_proekt <- proekt_bigram_topics %>% 
  group_by(topic) %>% 
  slice_max(beta, n=5) %>% ungroup() %>% arrange(topic, -beta)

top_bigrams_proekt %>% 
  mutate(term = reorder_within(term, beta, topic)) %>%
  ggplot(aes(beta, term, fill = factor(topic))) +
  geom_col(show.legend = FALSE) +
  facet_wrap(~topic, scales = "free") +
  scale_y_reordered()

write.xlsx(top_bigrams_proekt, "top_bigrams_proekt.xlsx")
top_bigrams_proekt <- read.xlsx("top_bigrams_proekt.xlsx")


top_bigrams_proekt %>% 
  mutate(term = reorder_within(term, beta, topic)) %>%
  ggplot(aes(beta, translated, fill = factor(topic))) +
  geom_col(show.legend = FALSE) +
  facet_wrap(~topic, scales = "free") +
  scale_y_reordered()


proekt_bigrams_gamma <- tidy(proekt_bigram_lda, matrix = "gamma")


proekt_bigrams_gamma %>% 
  mutate(From = reorder(document, gamma * topic)) %>%
  ggplot(aes(factor(topic), gamma)) +
  geom_boxplot() +
  facet_wrap(~ document) +
  labs(x = "topic", y = expression(gamma))



proekt_words <- ros_proekt_df %>% 
  unnest_tokens(word, Bodytext, "words")

proekt_keywords <- proekt_words %>% 
  stringdist_anti_join(ru_stopwords, c("word" = "words"))


proekt_sentiment_counts <- proekt_keywords %>%
  full_join(ru_sentiment, by = c("word" = "word"), copy = TRUE) %>%
  count(word, affect, sort = TRUE) %>%
  ungroup()

proekt_sentiment <- proekt_sentiment_counts %>%
  group_by(affect) %>%
  slice_max(n, n = 15) %>% 
  ungroup() %>%
  mutate(word = reorder(word, n))


write.xlsx(proekt_sentiment, "proekt_sentiment.xlsx")
proekt_sentiment <- read.xlsx("proekt_sentiment.xlsx")

proekt_sentiment %>%
  ggplot(aes(n, translated, fill = affect)) +
  geom_col(show.legend = FALSE) +
  facet_wrap(~affect, scales = "free_y") +
  labs(x = "Contribution to sentiment",
       y = NULL)


### Bigrams and sentiment for finance terms emails ###


finance_proekt_df <- as.data.frame(proekt_finance_df)
proekt_finance_bigrams <- finance_proekt_df %>% 
  unnest_tokens(bigram, Bodytext, token = "ngrams", n = 2)

proekt_finance_bigrams_sep <- proekt_finance_bigrams %>% 
  separate(bigram, c("word1", "word2"), sep = " ")
proekt_finance_keybigrams <- proekt_finance_bigrams_sep %>% 
  anti_join(ru_stopwords, c("word1" = "words")) %>% 
  anti_join(ru_stopwords, c("word2" = "words"))
proekt_finance_bigrams_key <- proekt_finance_keybigrams %>% 
  unite(bigram, word1, word2, sep = " ")

proekt_finance_bigram_counts <- proekt_finance_bigrams %>% 
  count(From, bigram, sort = TRUE)

proekt_finance_bigram_dtm <- proekt_finance_bigram_counts %>% 
  cast_dtm(From, bigram, n)

proekt_finance_bigram_lda <- LDA(
  proekt_finance_bigram_dtm, k=6, control = list(seed = 1234))

proekt_finance_bigram_topics <- tidy(proekt_finance_bigram_lda, matrix = "beta")

top_bigrams_proekt_finance <- proekt_finance_bigram_topics %>% 
  group_by(topic) %>% 
  slice_max(beta, n=5) %>% ungroup() %>% arrange(topic, -beta)

top_bigrams_proekt_finance %>% 
  mutate(term = reorder_within(term, beta, topic)) %>%
  ggplot(aes(beta, term, fill = factor(topic))) +
  geom_col(show.legend = FALSE) +
  facet_wrap(~topic, scales = "free") +
  scale_y_reordered()

write.xlsx(top_bigrams_proekt_finance, "top_bigrams_proekt_finance.xlsx")
top_bigrams_proekt_finance <- read.xlsx("top_bigrams_proekt_finance.xlsx")


top_bigrams_proekt_finance %>% 
  mutate(term = reorder_within(term, beta, topic)) %>%
  ggplot(aes(beta, translated, fill = factor(topic))) +
  geom_col(show.legend = FALSE) +
  facet_wrap(~topic, scales = "free") +
  scale_y_reordered()


proekt_finance_bigrams_gamma <- tidy(proekt_finance_bigram_lda, matrix = "gamma")


proekt_finance_bigrams_gamma %>% 
  mutate(From = reorder(document, gamma * topic)) %>%
  ggplot(aes(factor(topic), gamma)) +
  geom_boxplot() +
  facet_wrap(~ document) +
  labs(x = "topic", y = expression(gamma))



proekt_finance_words <- proekt_finance_df %>% 
  unnest_tokens(word, Bodytext, "words")

proekt_finance_keywords <- proekt_finance_words %>% 
  stringdist_anti_join(ru_stopwords, c("word" = "words"))


proekt_finance_sentiment_counts <- proekt_finance_keywords %>%
  full_join(ru_sentiment, by = c("word" = "word"), copy = TRUE) %>%
  count(word, affect, sort = TRUE) %>%
  ungroup()

proekt_finance_sentiment <- proekt_finance_sentiment_counts %>%
  group_by(affect) %>%
  slice_max(n, n = 15) %>% 
  ungroup() %>%
  mutate(word = reorder(word, n))


write.xlsx(proekt_finance_sentiment, "proekt_finance_sentiment.xlsx")
proekt_finance_sentiment <- read.xlsx("proekt_finance_sentiment.xlsx")

proekt_finance_sentiment %>%
  ggplot(aes(n, translated, fill = affect)) +
  geom_col(show.legend = FALSE) +
  facet_wrap(~affect, scales = "free_y") +
  labs(x = "Contribution to sentiment",
       y = NULL)

#### Searching for keywords from sentiment analysis ####


sentiment_extremes_df <- ros_proekt_df %>% filter(str_detect(Bodytext, regex("любов.*|€ростн.*|€рост.*", ignore_case = FALSE)))
write.xlsx(proekt_finance_df, "proekt_finance.xlsx")



