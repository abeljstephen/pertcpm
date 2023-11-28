#' @import readxl
#' @import dplyr
#' @import tidyr

# Load required libraries
library(readxl)
library(dplyr)
library(tidyr)



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
  cpm_data <- pert_data[, c("Activity", "Predecessor", "90%_Cumulative_Probability")]
  colnames(cpm_data) <- c("Activity", "Predecessor", "Duration")

  # Unpack rows with multiple predecessors
  unpacked_data <- cpm_data %>%
    separate_rows(Predecessor, sep = ",")

  # Create a data frame for Predecessors
  predecessors_data <- unpacked_data %>%
    select(Activity, Predecessor, Duration) %>%
    distinct()

  # Rename Predecessor to Successor in the unpacked data
  successors_data <- unpacked_data %>%
    filter(!is.na(Predecessor)) %>%
    select(Activity = Predecessor, Successor = Activity) %>%
    distinct()

  # Combine Predecessors and Successors data frames
  relationships_data <- bind_rows(predecessors_data, successors_data)

  # Add columns for EarlyStart and EarlyFinish
  relationships_data <- relationships_data %>%
    mutate(
      EarlyStart = 0,  # Initialize EarlyStart to 0
      EarlyFinish = 0  # Initialize EarlyFinish to 0
    )

  # Iterate through rows in relationships_data
  for (i in seq_len(nrow(relationships_data))) {
    # Check if Duration is missing
    if (is.na(relationships_data$Duration[i])) {
      # Find the corresponding Activity in unpacked_data
      matching_row <- unpacked_data[unpacked_data$Activity == relationships_data$Activity[i], ]

      # Check if a matching row is found
      if (nrow(matching_row) > 0) {
        # Use the Duration from matching row to fill in the missing Duration in relationships_data
        relationships_data$Duration[i] <- matching_row$Duration[1]
      }
    }
  }

  return(list(predecessors_data, successors_data, relationships_data))
}

# Function for forward pass
run_forward_pass <- function(predecessors_data) {
  # Assuming 'predecessors_data' is the data frame with columns Activity, Predecessor, and Duration
  data_forward <- predecessors_data %>%
    mutate(
      EarlyStart = ifelse(is.na(Predecessor), 0, NA),  # Initialize EarlyStart to 0 for the first activity
      EarlyFinish = ifelse(is.na(Predecessor), Duration, NA)  # Calculate EarlyFinish for the first activity
    )

  # Initialize indices for the loop
  preceding_activity_index <- 1
  next_activity_index <- preceding_activity_index + 1

  # Loop for forward pass
  while (next_activity_index <= nrow(data_forward)) {
    # Check if the current next_activity has a predecessor
    if (!is.na(data_forward$Predecessor[next_activity_index])) {
      # Activity has predecessor, get EarlyFinish of predecessor
      predecessor_finish <- data_forward$EarlyFinish[which(data_forward$Activity == data_forward$Predecessor[next_activity_index])]

      # Set EarlyStart for the current next_activity
      data_forward$EarlyStart[next_activity_index] <- max(predecessor_finish) + 1
    } else {
      # Activity has no predecessor, use the existing logic
      data_forward$EarlyStart[next_activity_index] <- max(data_forward$EarlyFinish) + 1
    }

    # Calculate EarlyFinish for the current next_activity
    data_forward$EarlyFinish[next_activity_index] <- data_forward$EarlyStart[next_activity_index] + data_forward$Duration[next_activity_index] - 1

    # Update indices for the next iteration
    preceding_activity_index <- next_activity_index
    next_activity_index <- preceding_activity_index + 1
  }

  # Assuming 'data_forward' is your original data frame
  # Replace 'YourColumnName' with the actual column name containing EarlyFinish values

  data_forward_final <- data_forward %>%
    group_by(Activity) %>%
    filter(EarlyFinish == max(EarlyFinish)) %>%
    ungroup()

  # Return the final result
  return(data_forward_final)
}



run_backward_pass <- function(forward_pass_result, relationships_data) {
  # Subset relationships_data where Successor exists
  backward_pass <- relationships_data[!is.na(relationships_data$Successor), ]

  # Drop the Predecessor column
  backward_pass <- backward_pass[, !(names(backward_pass) %in% "Predecessor")]

  # Find the last row of forward_pass_result
  last_row <- tail(forward_pass_result, 1)

  # Drop the Predecessor column
  last_row <- last_row[, !(names(last_row) %in% "Predecessor")]

  # Add column "Successor" with value NA after column Activity
  last_row <- cbind(last_row[, "Activity", drop = FALSE], Successor = NA, last_row[, -1])

  # Add the last row to backward_pass
  updated_backward_pass <- rbind(backward_pass, last_row)

  # Copy over the EarlyStart and EarlyFinish values to updated_backward_pass
  # Loop through backward_pass
  for (i in 1:nrow(backward_pass)) {
    # Match backward_pass$Activity with forward_pass_result$Activity
    match_index <- match(backward_pass$Activity[i], forward_pass_result$Activity)

    # Check if there's a match
    if (!is.na(match_index)) {
      # Copy forward_pass_result$EarlyStart and forward_pass_result$EarlyFinish to backward_pass
      updated_backward_pass$EarlyStart[i] <- forward_pass_result$EarlyStart[match_index]
      updated_backward_pass$EarlyFinish[i] <- forward_pass_result$EarlyFinish[match_index]
    }
  }

  # Add columns LateStart and LateFinish with initial values of 0
  updated_backward_pass$LateStart <- 0
  updated_backward_pass$LateFinish <- 0

  # Calculate last activity LateStart and LateFinish values
  # Determine the last row as the last activity
  last_activity_row <- updated_backward_pass[nrow(updated_backward_pass), ]

  # Set LateFinish value to be the same as EarlyFinish
  last_activity_row$LateFinish <- last_activity_row$EarlyFinish

  # Subtract EarlyFinish from Duration to set LateStart value
  last_activity_row$LateStart <- last_activity_row$LateFinish - last_activity_row$Duration

  # Update the last row in updated_backward_pass
  updated_backward_pass[nrow(updated_backward_pass), ] <- last_activity_row

  # Loop through rows starting from the second-to-last row
  for (index_underneath in (nrow(updated_backward_pass) - 1):1) {
    # Get the Successor value of the row underneath
    successor_value <- updated_backward_pass$Successor[index_underneath]

    # Find the first match of Successor value in the Activity column
    matching_row <- which(updated_backward_pass$Activity == successor_value)[1]

    # Take the LateStart of the matching row as LateFinish for the row underneath
    updated_backward_pass$LateFinish[index_underneath] <- updated_backward_pass$LateStart[matching_row]

    # Subtract Duration from LateFinish to get LateStart for the row underneath
    updated_backward_pass$LateStart[index_underneath] <- updated_backward_pass$LateFinish[index_underneath] - updated_backward_pass$Duration[index_underneath]
  }

  # Find the row with the minimum LateFinish for each Activity
  min_late_finish_rows <- updated_backward_pass %>%
    group_by(Activity) %>%
    slice(which.min(LateFinish))

  # Print the rows with minimum LateFinish for each Activity
  print(min_late_finish_rows)

  #return(updated_backward_pass)
  return(min_late_finish_rows)
}


calculate_slack <- function(backward_pass_result) {
  # Calculate Total Slack
  backward_pass_result$TotalSlack <- backward_pass_result$LateFinish - backward_pass_result$EarlyFinish

  # Calculate Free Slack
  backward_pass_result$FreeSlack <- ifelse(is.na(backward_pass_result$Successor),
                                           backward_pass_result$TotalSlack,
                                           lead(backward_pass_result$EarlyStart, default = max(backward_pass_result$LateFinish)) - backward_pass_result$LateFinish)

  # Identify Critical activities
  backward_pass_result$Critical <- ifelse(backward_pass_result$TotalSlack == 0, "Yes", "No")

  return(backward_pass_result)
}



clean_up_and_merge <- function(min_late_finish_rows, pert_data, round_digits = 4) {
  # Convert min_late_finish_rows to a data frame
  min_late_finish_table <- as.data.frame(min_late_finish_rows)

  # Print the resulting table
  print(min_late_finish_table)

  # Round numeric values in the data frame
  round_numeric <- function(x, digits = round_digits) {
    is_numeric <- sapply(x, is.numeric)
    x[, is_numeric] <- round(x[, is_numeric], digits)
    return(x)
  }

  # Apply the rounding function
  min_late_finish_table <- round_numeric(min_late_finish_table)

  # Assign min_late_finish_table to cpm_table
  cpm_table <- min_late_finish_table

  # Print the resulting table
  print(cpm_table)

  # Convert pert_data to a data frame
  pert_table <- as.data.frame(pert_data)

  # Merge data frames based on the "Activity" column
  pert_cpm_table <- merge(pert_table, cpm_table, by = "Activity", all = TRUE)

  # Print the resulting merged table
  print(pert_cpm_table)

  # Return the merged table
  return(pert_cpm_table)
}






# Parent function to execute PERT, CPM, and Forward Pass
pertcpm <- function(user_data) {
  # Step 1: PERT Calculation
  pert_data <- calculate_pert_metrics(user_data)

  # Step 2: CPM Calculation
  cpm_results <- create_cpm_data(pert_data)

  # Step 3: Forward Pass
  forward_pass_result <- run_forward_pass(cpm_results[[1]])

  # Step 4: Backward Pass
  relationships_data <- cpm_results[[3]]
  backward_pass_result <- run_backward_pass(forward_pass_result, relationships_data)


  # Step 5: Slack Calculation
  slack_result <- calculate_slack(backward_pass_result)

  # Step 6: Final Table
  # Assuming min_late_finish_rows and pert_data are available in your environment
  result_table <- clean_up_and_merge(backward_pass_result, pert_data)

  # Return the results
  # return(list(pert_data, cpm_results, forward_pass_result))
    return(result_table)
}


#libraries needed
#library(readxl)
#library(tidyr)
#library(dplyr)
# # Specify the data path, filename, and sheetname
# #datapath <- "C:/Users/abstephe/OneDrive - Microsoft/Documents/R Code/CriticalPath/Data/"
# datapath <- "/Users/abeljstephen/Documents/R Code/R/CriticalPath/Data"
# filename <- "CriticalPath.xlsx"
# sheetname <- "Sheet5"
# # Combine the path and filename
# file_path <- file.path(datapath, filename)
# # Read data from Excel file
# original_data <- readxl::read_excel(file_path, sheet = sheetname)
#
# # Example usage:
# user_data <- original_data
# result <- pertcpm(user_data)
# print(result)

