# Load necessary libraries
library(data.table)
library(dplyr)
library(ggplot2)
library(tidyr)
library(readxl)
library(openxlsx)

# Importing data
online_retail <- read_excel("Downloads/archive/online_retail_II.xlsx")

# Generate a 'Money' column and summarize NA values
online_retail$Money <- online_retail$Quantity * online_retail$Price
sum(is.na(online_retail))

# Data Cleaning
online_retail <- online_retail %>%
    mutate(
        Quantity = replace(Quantity, Quantity <= 0, NA),
        Price = replace(Price, Price <= 0, NA)
    ) %>%
    drop_na()

# Renaming column for consistency
names(online_retail)[names(online_retail) == "Customer ID"] <- "CustomerID"

# Data Type Conversion
online_retail <- online_retail %>%
    mutate(
        Invoice = as.factor(Invoice),
        StockCode = as.factor(StockCode),
        InvoiceDate = as.Date(InvoiceDate, '%m/%d/%Y %H:%M'),
        CustomerID = as.factor(CustomerID),
        Country = as.factor(Country)
    )

# RFM Analysis
reference_date <- max(online_retail$InvoiceDate) + 1
rfm_scores <- online_retail %>%
    group_by(CustomerID) %>%
    summarise(
        Recency = as.numeric(reference_date - max(as.Date(InvoiceDate))),
        Frequency = n_distinct(Invoice),
        Monetary = sum(Money)
    ) %>%
    mutate(
        RecencyScore = ntile(-Recency, 5),
        FrequencyScore = ntile(Frequency, 5),
        MonetaryScore = ntile(Monetary, 5)
    )

# Create segments based on RFM scores
rfm_scores <- rfm_scores %>%
    mutate(
        Segment = case_when(
            RecencyScore == 5 & FrequencyScore %in% 4:5 ~ "Champions",
            RecencyScore %in% 3:4 & FrequencyScore %in% 4:5 ~ "Loyal Customers",
            RecencyScore %in% 4:5 & FrequencyScore %in% 2:3 ~ "Potential Loyalists",
            RecencyScore == 5 & FrequencyScore == 1 ~ "New Customers",
            RecencyScore == 4 & FrequencyScore == 1 ~ "Promising",
            RecencyScore %in% 1:2 & FrequencyScore == 5 ~ "Can’t Lose Them",
            RecencyScore == 3 & FrequencyScore == 3 ~ "Need Attention",
            RecencyScore == 3 & FrequencyScore %in% 1:2 ~ "About To Sleep",
            RecencyScore %in% 1:2 & FrequencyScore %in% 3:4 ~ "At Risk",
            RecencyScore %in% 1:2 & FrequencyScore %in% 1:2 ~ "Hibernating"
        )
    )

# Customer count and percentage per segment
rfm_summary <- rfm_scores %>%
    group_by(RecencyScore, FrequencyScore, MonetaryScore) %>%
    summarise(Count = n(), .groups = 'drop') %>%
    mutate(Percentage = round((Count / sum(Count)) * 100, 1))

# Define color scheme for plotting
color_scheme <- c(
    "Champions" = "#1aa3ff",
    "Loyal Customers" = "#ccccff",
    "Potential Loyalists" = "#1affc6",
    "New Customers" = "#d2ff4d",
    "Promising" = "#cccc00",
    "Need Attention" = "#ffe680",
    "About To Sleep" = "#00cccc",
    "At Risk" = "#ff9933",
    "Can’t Lose Them" = "#ff5050",
    "Hibernating" = "#cce6ff"
)

# Plotting the RFM heatmap
ggplot(rfm_summary, aes(x = RecencyScore, y = FrequencyScore, fill = Segment)) +
    geom_tile(color = "white", size = 0.1) +
    geom_text(aes(label = paste0(Count, "\n(", Percentage, "%)")), size = 3) +
    scale_fill_manual(values = color_scheme) +
    labs(x = "Recency Score", y = "Frequency Score", fill = "Segment") +
    theme_minimal()

# Exporting RFM data
write.xlsx(rfm_summary, file = "rfm_count.xlsx")
write.xlsx(rfm_scores, file = "rfm_scores.xlsx")
