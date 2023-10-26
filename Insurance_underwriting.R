# Load necessary libraries
library(dplyr)
library(ggplot2)
library(readxl)
library(tidyverse)
library(sampling)
library(moments)
library(e1071)
library(benford.analysis)
library(rmarkdown)


# Load dataset
df <-
  read_excel(
    "L06101基础表保全清单.xlsx",
    # nolint
    skip = 1
  )

# take glimpse of dataframe
df %>% glimpse

# Filter only relevant columns and rows
df_filtered <- df %>%
  filter(保全类型 %in% c("退保", "减保", "保单贷款", "账户部分领取")) %>%
  mutate(补退费金额 = as.numeric(补退费金额))

# Remove rows with missing values if necessary
df_clean <- df_filtered %>% filter(!is.na(补退费金额))

# Function to detect outliers using the IQR method
detect_outliers_iqr <- function(data) {
  Q1 <- quantile(data, 0.25) # nolint: object_name_linter.
  Q3 <- quantile(data, 0.75) # nolint
  IQR <- Q3 - Q1 # nolint: object_name_linter.
  lower_bound <- Q1 - 1.5 * IQR
  upper_bound <- Q3 + 1.5 * IQR
  outliers <- data[data < lower_bound | data > upper_bound]
  return(outliers)
}

# Detect outliers for each category & summaries outliers using irq function
outliers <- df_clean %>%
  group_by(保全类型) %>%
  summarise(Outliers = list(detect_outliers_iqr(补退费金额)))
print(outliers) # output results

# retrieve the summary and set relevant values
summary(df_clean$补退费金额)

# set up values
min_value <- 0
max_value <- 500000
interval <- 50000

# select varies programs
df_clean %>% glimpse

# Box plot with adjusted y-axis
Q1 <- quantile(df_clean$补退费金额, 0.25)
Q3 <- quantile(df_clean$补退费金额, 0.75)
IQR <- Q3 - Q1

# set up outliers with iqr
df_clean_no_outliers <-
  df_clean[df_clean$补退费金额 >= (Q1 - 1.5 * IQR) &
             df_clean$补退费金额 <= (Q3 + 1.5 * IQR),]

# set up ggplot
ggplot(df_clean_no_outliers, aes(x = 保全类型, y = 补退费金额)) +
  geom_boxplot() +
  labs(title = "Boxplot of 补退费金额 by 保全类型 (No Outliers)", x = "保全类型", y = "补退费金额")

# Scatter plot with adjusted y-axis
ggplot(df_clean, aes(x = 保全类型, y = 补退费金额)) +
  geom_point() +
  labs(title = "Scatterplot of 补退费金额 by 保全类型", x = "保全类型", y = "补退费金额") +
  scale_y_continuous(
    limits = c(min_value, max_value),
    breaks = seq(min_value, max_value, interval)
  )

# Function to detect outliers using the IQR method (returns a logical vector)
detect_outliers_iqr <- function(data) {
  Q1 <- quantile(data, 0.25)
  Q3 <- quantile(data, 0.75)
  IQR <- Q3 - Q1
  lower_bound <- Q1 - 1.5 * IQR
  upper_bound <- Q3 + 1.5 * IQR
  outliers <- data < lower_bound | data > upper_bound
  return(outliers)
}

# Add an index column to the cleaned data frame
df_clean <- df_clean %>% mutate(Index = row_number())

# Find the 保单号 of outliers for each category
outlier_indices_and_保单号 <- df_clean %>%
  nest_by(保全类型) %>%
  mutate(Outliers = list(detect_outliers_iqr(data$补退费金额))) %>%
  unnest(cols = c(data, Outliers)) %>%
  filter(Outliers) %>%
  select(保全类型, Index, 保单号)

# print the results
print(outlier_indices_and_保单号)

# view the whole list
view(outlier_indices_and_保单号)

# set up sample size
sample_size <- c(退保 = 4,
                 保单贷款 = 4,
                 联系方式变更 = 4)

# convert to dataframe
df_sampling <- as.data.frame(outlier_indices_and_保单号)

# perform stratified sampling with non_replacement sampling
set.seed(1234)

# using stratified sampling
stratified_sample <- strata(
  data = df_sampling,
  stratanames = "保全类型",
  size = sample_size,
  method = "srswor" # non-replacement sampling
)


# retrieve results
sampled_indices <- stratified_sample$ID_unit

# retrieve samples
sampled_data <- df_sampling[sampled_indices, ]

# output results
sampled_data

# Monte Carlo simulation
n_simulations <- 10000
simulated_means <- numeric(n_simulations)
simulated_sd <- sd(df_clean$补退费金额)

# loop through the simulation using 补退费金额
for (i in 1:n_simulations) {
  simulated_sample <-
    rnorm(nrow(df_clean),
          mean = mean(df_clean$补退费金额),
          sd = simulated_sd)
  simulated_means[i] <- mean(simulated_sample)
}

# Compare the mean of the original data to the simulated means
hist(simulated_means, main = "Monte Carlo Simulation Randomness Detection", xlab = "Mean")
abline(v = mean(df_clean$补退费金额),
       col = "red",
       lwd = 2)

# take glimpse of dataframe
df %>% glimpse

# Filter only relevant columns and rows
df_filtered <- df %>%
  filter(保全类型 %in% c("退保", "联系方式变更", "保单贷款", "签名变更")) %>%
  mutate(补退费金额 = as.numeric(补退费金额))

# Remove rows with missing values if necessary
df_clean <- df_filtered %>% filter(!is.na(补退费金额))

# Function to detect outliers using the IQR method
detect_outliers_iqr <- function(data) {
  Q1 <- quantile(data, 0.25)
  Q3 <- quantile(data, 0.75)
  IQR <- Q3 - Q1
  lower_bound <- Q1 - 1.5 * IQR
  upper_bound <- Q3 + 1.5 * IQR
  outliers <- data[data < lower_bound | data > upper_bound]
  return(outliers)
}

# Detect outliers for each category:
outliers <- df_clean %>%
  group_by(保全类型) %>%
  summarise(Outliers = list(detect_outliers_iqr(补退费金额)))
print(outliers)

# retrieve the summary and set relevant values
summary(df_clean$补退费金额)
min_value <- 0
max_value <- 500000
interval <- 50000

# select varies programs
df_clean %>% head()

# Box plot with adjusted y-axis
Q1 <- quantile(df_clean$补退费金额, 0.25)
Q3 <- quantile(df_clean$补退费金额, 0.75)
IQR <- Q3 - Q1

# set up outliers with iqr
df_clean_no_outliers <-
  df_clean[df_clean$补退费金额 >= (Q1 - 1.5 * IQR) &
             df_clean$补退费金额 <= (Q3 + 1.5 * IQR),]

# set up ggplot
ggplot(df_clean_no_outliers, aes(x = 保全类型, y = 补退费金额)) +
  geom_boxplot() +
  labs(title = "Boxplot of 补退费金额 by 保全类型 (No Outliers)", x = "保全类型", y = "补退费金额")

# Scatter plot with adjusted y-axis
ggplot(df_clean, aes(x = 保全类型, y = 补退费金额)) +
  geom_point() +
  labs(title = "Scatterplot of 补退费金额 by 保全类型", x = "保全类型", y = "补退费金额") +
  scale_y_continuous(
    limits = c(min_value, max_value),
    breaks = seq(min_value, max_value, interval)
  )

# Function to detect outliers using the IQR method (returns a logical vector)
detect_outliers_iqr <- function(data) {
  Q1 <- quantile(data, 0.25)
  Q3 <- quantile(data, 0.75)
  IQR <- Q3 - Q1
  lower_bound <- Q1 - 1.5 * IQR
  upper_bound <- Q3 + 1.5 * IQR
  outliers <- data < lower_bound | data > upper_bound
  return(outliers)
}

# Add an index column to the cleaned data frame
df_clean <- df_clean %>% mutate(Index = row_number())

# Find the 保单号 of outliers for each category
outlier_indices_and_保单号 <- df_clean %>%
  nest_by(保全类型) %>%
  mutate(Outliers = list(detect_outliers_iqr(data$补退费金额))) %>%
  unnest(cols = c(data, Outliers)) %>%
  filter(Outliers) %>%
  select(保全类型, Index, 保单号)

# print the results
print(outlier_indices_and_保单号)

# view the whole list
view(outlier_indices_and_保单号)

# set up sample size
sample_size <- c(退保 = 4,
                 保单贷款 = 4,
                 联系方式变更 = 4)

# convert to dataframe
df_sampling <- as.data.frame(outlier_indices_and_保单号)

# perform stratified sampling with non_replacement sampling
set.seed(1234)

# stratified sampling
stratified_sample <- strata(
  data = df_sampling,
  stratanames = "保全类型",
  size = sample_size,
  method = "srswor"                 # non-replacement sampling
)


# retrieve results
sampled_indices <- stratified_sample$ID_unit

# retrieve samples
sampled_data <- df_sampling[sampled_indices, ]

# output results
sampled_data

# Monte Carlo simulation
n_simulations <- 20000
simulated_means <- numeric(n_simulations)

# loop through the simulations
for (i in 1:n_simulations) {
  simulated_sample <-
    sample(df_clean$补退费金额,
           size = nrow(df_clean),
           replace = TRUE)
  simulated_means[i] <- mean(simulated_sample)
}

# Compare the mean of the original data to the simulated means
hist(simulated_means, main = "Monte Carlo Simulation Randomness Detection", xlab = "Mean")
abline(v = mean(df_clean$补退费金额),
       col = "red",
       lwd = 2)


##############################################################################
# Group by 被保人姓名 and calculate the sum of 补退费金额
df_sum_by_name <- df_clean %>%
  group_by(被保人姓名) %>% # group by 被保险人姓名
  summarise(补退费金额总和 = sum(补退费金额))

# Calculate the total 补退费金额
total_补退费金额 <- sum(df_sum_by_name$补退费金额总和)

# Calculate the proportion of 补退费金额 for each 被保人姓名
df_sum_by_name <- df_sum_by_name %>%
  mutate(补退费金额占比 = 补退费金额总和 / total_补退费金额) # create a new column with mutate

# Calculate the Herfindahl Index
herfindahl_index <-
  sum(df_sum_by_name$补退费金额占比 ^ 2) # sum of all squares

# Print the result
print(herfindahl_index)

# Group by 业务员姓名 and calculate the sum of 补退费金额
df_sum_by_sales <- df_clean %>%
  group_by(业务员姓名) %>%
  summarise(补退费金额总和 = sum(补退费金额))

# Calculate the total 补退费金额
total_补退费金额 <- sum(df_sum_by_sales$补退费金额总和) # nolint

# Calculate the proportion of 补退费金额 for each 业务员姓名
df_sum_by_sales <- df_sum_by_sales %>%
  mutate(补退费金额占比 = 补退费金额总和 / total_补退费金额)

# Calculate the Herfindahl Index
herfindahl_index_sales <- sum(df_sum_by_sales$补退费金额占比 ^ 2)

# Print the result
print(herfindahl_index_sales)

# Group by 险种名称 and calculate the sum of 补退费金额
df_sum_by_items <- df_clean %>%
  group_by(险种名称) %>%
  summarise(补退费金额总和 = sum(补退费金额))


# Calculate the total 补退费金额
total_补退费金额 <- sum(df_sum_by_items$补退费金额总和)

# Calculate the proportion of 补退费金额 for each 险种名称
df_sum_by_items <- df_sum_by_items %>%
  mutate(补退费金额占比 = 补退费金额总和 / total_补退费金额)

# Calculate the Herfindahl Index
herfindahl_index_items <- sum(df_sum_by_items$补退费金额占比 ^ 2)

# Print the result
print(herfindahl_index_items)

# the red line falls within the range of the simulated means, it suggests that
# the outliers in the original dataset might be random, as they don't significantly
# affect the mean value.
# Sort the df_sum_by_items dataframe by 补退费金额占比 in descending order
df_sum_by_items_sorted <- df_sum_by_items %>%
  arrange(desc(补退费金额占比))

# Print the top few rows
head(df_sum_by_items_sorted)

# Group by 险种名称 and 保全类型, and calculate the sum of 补退费金额
df_sum_by_items_and_保全类型 <- df_clean %>%
  group_by(险种名称, 保全类型) %>%
  summarise(补退费金额总和 = sum(补退费金额), .groups = "drop")

# Calculate the total 补退费金额
total_补退费金额 <- sum(df_sum_by_items_and_保全类型$补退费金额总和)

# Calculate the proportion of 补退费金额 for each combination of 险种名称 and 保全类型
df_sum_by_items_and_保全类型 <- df_sum_by_items_and_保全类型 %>%
  mutate(补退费金额占比 = 补退费金额总和 / total_补退费金额)

# Sort the df_sum_by_items_and_保全类型 dataframe by 补退费金额占比 in descending order
df_sum_by_items_and_保全类型_sorted <- df_sum_by_items_and_保全类型 %>%
  arrange(desc(补退费金额占比))

# Print the top few rows
head(df_sum_by_items_and_保全类型_sorted)

