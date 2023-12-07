#' @import readxl
#' @import dplyr
#' @import tidyr
#' @import stringr

# Load required libraries
library(readxl)
library(dplyr)
library(tidyr)
library(stringr)

# Function to calculate Pert_Duration, Standard_Deviation, and Probability
calculate_metrics <- function(optimistic, expected, pessimistic) {
  pert_duration <- (4 * expected + pessimistic + optimistic) / 6
  variance <- ((pessimistic - optimistic) / 6)^2
  std_deviation <- sqrt(variance)

  # Calculate probabilities
  probability_68 <- pert_duration + std_deviation
  probability_95 <- pert_duration + 2 * std_deviation
  probability_99 <- pert_duration + 3 * std_deviation

  # Calculate lower and upper boundaries for each probability
  lower_boundary_68 <- pert_duration - std_deviation
  upper_boundary_68 <- pert_duration + std_deviation

  lower_boundary_95 <- pert_duration - 2 * std_deviation
  upper_boundary_95 <- pert_duration + 2 * std_deviation

  lower_boundary_99 <- pert_duration - 3 * std_deviation
  upper_boundary_99 <- pert_duration + 3 * std_deviation

  # Calculate cumulative probability
  cumulative_probability_90 <- pert_duration + 1.2816 * std_deviation

  return(c(pert_duration, std_deviation,
           lower_boundary_68, upper_boundary_68,
           lower_boundary_95, upper_boundary_95,
           lower_boundary_99, upper_boundary_99,
           cumulative_probability_90))
}


# Function to calculate PERT metrics for the entire project
calculate_pert_metrics <- function(project_data) {

  #gsub function to replace all occurrences of whitespace (\\s) with an empty string in the Predecessor column
  project_data$Predecessor <- gsub("\\s", "", project_data$Predecessor)

  # Apply the function to each row of the data frame
  metrics <- apply(project_data[, c("Optimistic_Duration", "Expected_Duration", "Pessimistic_Duration")], 1,
                   function(x) calculate_metrics(x[1], x[2], x[3]))

  # Create a new data frame with the calculated metrics
  result <- cbind(project_data, t(metrics))
  colnames(result) <- c(colnames(project_data), "Pert_Duration", "Standard_Deviation",
                        "68.3%_Probability_Lower_Boundary", "68.3%_Probability_Upper_Boundary",
                        "95.5%_Probability_Lower_Boundary", "95.5%_Probability_Upper_Boundary",
                        "99.7%_Probability_Lower_Boundary", "99.7%_Probability_Upper_Boundary",
                        "90%_Cumulative_Probability")

  return(result)
}

create_cpm_data <- function(pert_data) {
  # Subset the table to create cpm_data
  # the end result will be Activity, Predecessor, Duration, Successor EarlyStart, EarlyFinish which will be accessible in [[3]]
  # EarlyStart and EarlyFinish will be all 0's and will be calculated with values by next function i.e., run_forward_pass

  cpm_data <- pert_data[, c("Activity", "Predecessor", "90%_Cumulative_Probability")]
  colnames(cpm_data) <- c("Activity", "Predecessor", "Duration")
  #gsub function to replace all occurrences of whitespace (\\s) with an empty string in the Predecessor column
  cpm_data$Predecessor <- gsub("\\s", "", cpm_data$Predecessor)

  # Unpack rows with multiple predecessors
  unpacked_data <- cpm_data %>%
    tidyr::separate_rows(Predecessor, sep = ",")

  # Create a data frame for Predecessors
  predecessors_data <- unpacked_data %>%
    dplyr::select(Activity, Predecessor, Duration) %>%
    dplyr::distinct()

  # Use the value in the Predecessor column to determine that the Activity in same row is the Successor
  successors_data <- unpacked_data %>%
    dplyr::filter(!is.na(Predecessor)) %>%
    dplyr::select(Activity = Predecessor, Successor = Activity) %>%
    dplyr::distinct()

  # Merge Predecessors and Successors data frames based on the "Activity" column
   relationships_data <- merge(successors_data, predecessors_data, by = "Activity", all = TRUE)

  return(relationships_data)
}


calculate_forward_pass <- function(cpm_data) {
  forward_pass_raw_data <- cpm_data

  # Assuming forward_pass_raw_data is your initial dataset

  # Remove leading and trailing whitespaces from the "Activity" column and filter out rows with empty "Activity" values
  forward_pass_raw_data <- forward_pass_raw_data %>%
    filter(!is.na(Activity) & str_trim(Activity) != "")

  # Replace empty values in the Predecessor column with NA
  forward_pass_raw_data$Predecessor[forward_pass_raw_data$Predecessor == ""] <- NA

  # Add columns for EarlyStart and EarlyFinish
  forward_pass_raw_data <- forward_pass_raw_data %>%
    mutate(
      EarlyStart = ifelse(is.na(Predecessor), 0, NA),  # Initialize EarlyStart to 0 for the first activity
      EarlyFinish = ifelse(is.na(Predecessor), Duration, NA) # Calculate EarlyFinish for the first activity
    )

  # Identify tasks with no predecessors
  start_tasks <- forward_pass_raw_data[is.na(forward_pass_raw_data$Predecessor), ]

  # Set EarlyStart and EarlyFinish for start tasks
  start_tasks$EarlyStart <- 0
  start_tasks$EarlyFinish <- start_tasks$Duration

  # Loop through the start tasks and update successors
  for (i in seq_along(start_tasks$Activity)) {
    current_activity <- start_tasks[i, ]

    # Find successors for the current activity
    successors <- forward_pass_raw_data[forward_pass_raw_data$Predecessor == current_activity$Activity, ]

    # Update EarlyStart and EarlyFinish for the successors
    forward_pass_raw_data$EarlyStart[forward_pass_raw_data$Activity %in% successors$Activity] <- current_activity$EarlyFinish + 1
    forward_pass_raw_data$EarlyFinish[forward_pass_raw_data$Activity %in% successors$Activity] <- forward_pass_raw_data$EarlyStart[forward_pass_raw_data$Activity %in% successors$Activity] + forward_pass_raw_data$Duration[forward_pass_raw_data$Activity %in% successors$Activity]
  }

  # Update the remaining tasks that haven't been processed
  while (any(is.na(forward_pass_raw_data$EarlyStart) | is.na(forward_pass_raw_data$EarlyFinish))) {
    for (i in seq_along(forward_pass_raw_data$Activity)) {
      current_activity <- forward_pass_raw_data[i, ]

      # Check if EarlyStart and EarlyFinish are still NA
      if (is.na(current_activity$EarlyStart) && is.na(current_activity$EarlyFinish)) {
        # Find predecessors for the current activity
        predecessors <- forward_pass_raw_data[forward_pass_raw_data$Activity %in% current_activity$Predecessor, ]

        # Update EarlyStart and EarlyFinish for the current activity
        if (nrow(predecessors) > 0) {
          forward_pass_raw_data$EarlyStart[i] <- max(predecessors$EarlyFinish) + 1
          forward_pass_raw_data$EarlyFinish[i] <- forward_pass_raw_data$EarlyStart[i] + current_activity$Duration
        }
      }
    }
  }



  # Order activities based on Early Start and Early Finish
  forward_pass_processed_data <- forward_pass_raw_data %>%
    arrange(EarlyStart, EarlyFinish) %>%
    group_by(Activity) %>%
    slice(1) %>%
    ungroup()


  return(list(forward_pass_raw_data = forward_pass_raw_data, forward_pass_processed_data = forward_pass_processed_data))
}





calculate_backward_pass <- function(forward_pass_processed_data) {
  backward_pass_data <- forward_pass_processed_data

  # Initialize LateStart and LateFinish columns
  backward_pass_data$LateStart <- 0
  backward_pass_data$LateFinish <- 0

  # Set LateFinish for the tasks that have no Successors
  tasks_with_no_successors <- backward_pass_data$Activity[is.na(backward_pass_data$Successor)]
  backward_pass_data$LateFinish[backward_pass_data$Activity %in% tasks_with_no_successors] <- max(backward_pass_data$EarlyFinish)

  # Continue the backward pass calculation
  for (i in seq(nrow(backward_pass_data), 1, -1)) {
    current_activity <- backward_pass_data[i, ]  # Fix: Use backward_pass_data[i, ] instead of backward_pass_data[i]

    # Update LateFinish for the current activity
    if (!is.na(current_activity$Successor)) {  # Fix: Use current_activity$Activity instead of backward_pass_data$Activity[i]
      successors <- backward_pass_data$Successor[which(backward_pass_data$Activity == current_activity$Successor)]

      # Exclude missing values when calculating the minimum
      valid_successors <- successors[!is.na(successors)]
      if (length(valid_successors) > 0) {
        backward_pass_data$LateFinish[i] <- min(backward_pass_data$LateStart[backward_pass_data$Activity %in% valid_successors]) - 1
      } else {
        # If all successors have missing LateStart values, set LateFinish to maximum
        backward_pass_data$LateFinish[i] <- max(backward_pass_data$LateFinish)
      }
    } else {
      # Task with no successors
      backward_pass_data$LateFinish[i] <- max(backward_pass_data$LateFinish)
    }

    # Update LateStart based on LateFinish
    backward_pass_data$LateStart[i] <- backward_pass_data$LateFinish[i] - backward_pass_data$Duration[i] + 1
  }

  return(backward_pass_data)
}



determine_critical_path <- function(backward_pass_data) {
  # Calculate Total Slack
  backward_pass_data$TotalSlack <- round(backward_pass_data$LateFinish - backward_pass_data$EarlyFinish, 4)

  # Create a data frame to store the EarlyStart values of Successors
  successor_earlystart <- data.frame(Activity = backward_pass_data$Activity, EarlyStart_Successor = NA)

  # Loop through rows in data
  for (i in seq_len(nrow(backward_pass_data))) {
    # Check if the Successor column is not NA
    if (!is.na(backward_pass_data$Successor[i])) {
      # Find the matching row in successor_earlystart
      matching_row <- which(successor_earlystart$Activity == backward_pass_data$Successor[i])

      # Update the EarlyStart_Successor value in the matching row
      successor_earlystart$EarlyStart_Successor[matching_row] <- round(backward_pass_data$EarlyStart[i], 4)
    }
  }

  # Calculate Free Slack considering negative values
  backward_pass_data$FreeSlack <- ifelse(is.na(backward_pass_data$Successor),
                                         round(backward_pass_data$TotalSlack, 4),
                                         ifelse(is.na(lead(successor_earlystart$EarlyStart_Successor)),
                                                max(backward_pass_data$LateFinish) - backward_pass_data$LateFinish,
                                                round(lead(successor_earlystart$EarlyStart_Successor) - backward_pass_data$LateFinish, 4)))

  # Identify OnCriticalPath and AlreadyDelayed
  backward_pass_data$OnCriticalPath <- ifelse(backward_pass_data$TotalSlack == 0, "Yes", "No")
  backward_pass_data$AlreadyDelayed <- ifelse(backward_pass_data$FreeSlack < 0, "Yes", "No")

  # Round numeric columns to 4 decimal places
  numeric_columns <- sapply(backward_pass_data, is.numeric)
  backward_pass_data[numeric_columns] <- round(backward_pass_data[numeric_columns], 4)

  return(backward_pass_data)
}


# Parent function to execute PERT, CPM, and Forward Pass
pertcpm <- function(project_data) {
  # Step 1: PERT Calculation
  pert_data <- calculate_pert_metrics(project_data)

  # Step 2: CPM Data Preparation
  cpm_data <- create_cpm_data(pert_data)

  # Step 3: Forward Pass
  forward_pass_all_data <- calculate_forward_pass(cpm_data)
  forward_pass_raw_data<- forward_pass_all_data$forward_pass_raw_data
  forward_pass_processed_data<- forward_pass_all_data$forward_pass_processed_data

  # Step 4: Backward Pass
  backward_pass_data<- calculate_backward_pass(forward_pass_processed_data)

  # Step 5: Critical Path
  critical_path <- determine_critical_path(backward_pass_data)

  # Step 6:
  library(dplyr)

  # Drop the "Predecessor" column from project_data
  project_data <- project_data %>%
    select(-Predecessor)

  # Merge the modified project_data with critical_path based on the "Activity" column
  merged_data <- project_data %>%
    left_join(critical_path, by = "Activity")

  return(merged_data)

}


# # Specify the data path, filename, and sheetname
# # datapath <- "C:/Users/abstephe/OneDrive - Microsoft/Documents/R Code/CriticalPath/Data/"
#  datapath <- "/Users/abeljstephen/Documents/R Code/R/CriticalPath/Data"
#  filename <- "CriticalPath.xlsx"
#  sheetname <- "Sheet5"
# # Combine the path and filename
#  file_path <- file.path(datapath, filename)
# # Read data from Excel file
#  original_data <- readxl::read_excel(file_path, sheet = sheetname)
#
# # Example usage:
#  project_data<- original_data
#  critical_path_final <- pertcpm(project_data)
#  print(critical_path_final, n = Inf)



