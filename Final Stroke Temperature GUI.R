library(shiny)
library(shinyjs)
library(shinythemes)    # Using the cerulean theme
library(readr)
library(dplyr)
library(plotly)
library(ggplot2)
library(tools)          # for file_path_sans_ext
library(openxlsx)       # For exporting Excel workbooks
library(data.table)
# ------------------------------
# Data Loading & Preprocessing (Static)
# ------------------------------

file_path <- "/Users/druschel/Documents/R/Temp App/VG 10-22-25 - Fs RENAMED.csv"
preferred_min <- 36.6   # Preferred minimum temperature
preferred_max <- 36.8   # Preferred maximum temperature
mouse_min <- 30.0       # Minimum temperature for mouse (rectal probe connected)
mouse_max <- 39.0       # Maximum temperature for mouse (rectal probe connected)

# Parse those lines into a data frame
data <- read_csv(file_path, show_col_types = FALSE)

data <- data %>%
  rename(Time = `...1`) %>%
  rename_with(~ "Channel_B", grep("Channel B", names(data))) %>%
  rename_with(~ "Channel_C", grep("Channel C", names(data)))
# Detect format based on the first time value


# Check first time value and format accordingly
first_time_value <- as.character(data$Time[1])

if (startsWith(first_time_value, "00:")) {
  # Relative time format (HH:MM:SS) → assume midnight
  data$Time <- as.POSIXct(data$Time, format = "%H:%M:%S", tz = "UTC")
  midnight <- as.POSIXct("1970-01-01 00:00:00", tz = "UTC")
  data <- data %>% mutate(data_secs = as.numeric(difftime(Time, midnight, units = "secs")))
} else {
  # Absolute time (e.g. 2025-02-19T09:37:48-08:00)
  data$Time <- as.POSIXct(data$Time, format = "%Y-%m-%dT%H:%M:%S", tz = "America/Los_Angeles")
  data <- data %>% mutate(data_secs = as.numeric(difftime(Time, min(Time, na.rm = TRUE), units = "secs")))
  midnight <- min(data$Time, na.rm = TRUE)  # used for shifting later
}


filtered_data_B <- data %>% filter(Channel_B >= mouse_min & Channel_B <= mouse_max)
filtered_data_C <- data %>% filter(Channel_C >= mouse_min & Channel_C <= mouse_max)

# ------------------------------
# UI
# ------------------------------
ui <- fluidPage(
  theme = shinytheme("cerulean"),
  useShinyjs(),
  titlePanel("Temperature Data Review"),
  
  # Row 1: Shift controls and a toggle for Actual Start Time.
  fluidRow(
    column(4,
           actionButton("toggle_shift", "Shift"),
           hidden(
             div(id = "shift_controls",
                 sliderInput("shift_time", "Shift Time (seconds):",
                             min = -3600, max = 3600, value = 0, step = 1),
                 sliderInput("shift_temp", "Shift Temperature (°C):",
                             min = -3, max = 3, value = 0, step = 0.01)
             )
           )
    ),
    column(4,
           actionButton("set_actual_time", "Set Actual Time"),
           hidden(
             div(id = "actual_time_controls",
                 textInput("actual_start_time", "Enter Actual Start Time (e.g., 9:00 AM):", value = "")
             )
           )
    )
  ),
  
  # Row 2: Action buttons.
  fluidRow(
    column(12,
           actionButton("remove_region", "Remove Selected Region"),
           actionButton("run_stats", "Run Stats on Selected Region"),
           downloadButton("download_filtered", "Download Filtered Data"),
           downloadButton("download_stats", "Download Stats Excel")
    )
  ),
  
  # Row 3: Time-range slider.
  fluidRow(
    column(12,
           sliderInput("time_range", "Select Time Range:",
                       min = min(data$Time, na.rm = TRUE),
                       max = max(data$Time, na.rm = TRUE),
                       value = c(min(data$Time, na.rm = TRUE), max(data$Time, na.rm = TRUE)),
                       timeFormat = "%H:%M:%S",
                       width = "100%", step = 1)
    )
  ),
  
  # Row 4: Tabset panel for plots and stats.
  fluidRow(
    column(12,
           tabsetPanel(
             id = "channel_tab",
             tabPanel("Channel B",
                      plotlyOutput("temp_plot_B", height = "400px"),
                      uiOutput("stats_B")  # Rendered as formatted HTML.
             ),
             tabPanel("Channel C",
                      plotlyOutput("temp_plot_C", height = "400px"),
                      uiOutput("stats_C")
             )
           )
    )
  )
)

# ------------------------------
# Server
# ------------------------------
server <- function(input, output, session) {
  
  # ReactiveValues to store graph snapshots for each animal.
  graphSnapshots <- reactiveValues(full = list(), animal = list())
  
  # Save Animal ID for use in plot titles.
  last_animal <- reactiveVal("")
  
  # Toggle the actual time input when the "Set Actual Time" button is clicked.
  observeEvent(input$set_actual_time, {
    toggle("actual_time_controls")
  })
  
  # Reactive values for each channel.
  reactive_data_B <- reactiveVal(filtered_data_B)
  reactive_data_C <- reactiveVal(filtered_data_C)
  
  # Compute base_offset (in seconds) from the optional Actual Start Time.
  base_offset <- reactive({
    offset <- 0
    if (nzchar(input$actual_start_time)) {
      parsed_time <- as.POSIXct(input$actual_start_time, format = "%I:%M %p", tz = "UTC")
      if (!is.na(parsed_time)) {
        offset <- as.numeric(difftime(parsed_time, midnight, units = "secs"))
      }
    }
    offset
  })
  
  # Effective offset = base_offset + manual shift_time.
  effective_offset <- reactive({ base_offset() + input$shift_time })
  
  # Function to compute new_time = midnight + (data_secs + effective_offset)
  transform_time <- function(df) {
    df %>% mutate(new_time = midnight + data_secs + effective_offset())
  }
  
  # Update reactive data for the active channel.
  observe({
    if (input$channel_tab == "Channel B") {
      reactive_data_B(
        transform_time(filtered_data_B) %>% mutate(Channel_B = Channel_B + input$shift_temp)
      )
    } else {
      reactive_data_C(
        transform_time(filtered_data_C) %>% mutate(Channel_C = Channel_C + input$shift_temp)
      )
    }
  })
  
  # Update the time_range slider based on new_time.
  observe({
    new_min <- min(transform_time(data)$new_time, na.rm = TRUE)
    new_max <- max(transform_time(data)$new_time, na.rm = TRUE)
    updateSliderInput(session, "time_range", min = new_min, max = new_max,
                      value = c(new_min, new_max))
  })
  
  observeEvent(input$toggle_shift, {
    toggle("shift_controls")
  })
  
  # Create a reactive static plot (full graph) with Animal ID in the title.
  static_plot <- reactive({
    range <- input$time_range
    title_suffix <- ""
    if(nzchar(last_animal())) {
      title_suffix <- paste(" (Animal:", last_animal(), ")")
    }
    if (input$channel_tab == "Channel B") {
      ggplot(reactive_data_B(), aes(x = new_time)) +
        geom_line(aes(y = Channel_B, color = "Channel B")) +
        annotate("rect",
                 xmin = min(transform_time(data)$new_time, na.rm = TRUE),
                 xmax = max(transform_time(data)$new_time, na.rm = TRUE),
                 ymin = preferred_min, ymax = preferred_max,
                 fill = "green", alpha = 0.1) +
        geom_vline(xintercept = as.numeric(range[1]), color = "blue", linetype = "dashed") +
        geom_vline(xintercept = as.numeric(range[2]), color = "blue", linetype = "dashed") +
        annotate("rect", xmin = range[1], xmax = range[2],
                 ymin = -Inf, ymax = Inf, fill = "blue", alpha = 0.2) +
        labs(title = paste("Channel B Temperature Analysis", title_suffix),
             x = "Time (HH:MM:SS)",
             y = "Temperature (°C)",
             color = "Channels") +
        scale_color_manual(values = c("Channel B" = "red")) +
        theme_minimal()
    } else {
      ggplot(reactive_data_C(), aes(x = new_time)) +
        geom_line(aes(y = Channel_C, color = "Channel C")) +
        annotate("rect",
                 xmin = min(transform_time(data)$new_time, na.rm = TRUE),
                 xmax = max(transform_time(data)$new_time, na.rm = TRUE),
                 ymin = preferred_min, ymax = preferred_max,
                 fill = "green", alpha = 0.1) +
        geom_vline(xintercept = as.numeric(range[1]), color = "blue", linetype = "dashed") +
        geom_vline(xintercept = as.numeric(range[2]), color = "blue", linetype = "dashed") +
        annotate("rect", xmin = range[1], xmax = range[2],
                 ymin = -Inf, ymax = Inf, fill = "blue", alpha = 0.2) +
        labs(title = paste("Channel C Temperature Analysis", title_suffix),
             x = "Time (HH:MM:SS)",
             y = "Temperature (°C)",
             color = "Channels") +
        scale_color_manual(values = c("Channel C" = "blue")) +
        theme_minimal()
    }
  })
  
  # Render the Plotly plots.
  output$temp_plot_B <- renderPlotly({
    ggplotly(static_plot())
  })
  output$temp_plot_C <- renderPlotly({
    ggplotly(static_plot())
  })
  
  # Remove selected region.
  observeEvent(input$remove_region, {
    range <- input$time_range
    if (input$channel_tab == "Channel B") {
      reactive_data_B(
        reactive_data_B() %>% filter(new_time < range[1] | new_time > range[2])
      )
    } else {
      reactive_data_C(
        reactive_data_C() %>% filter(new_time < range[1] | new_time > range[2])
      )
    }
  })
  
  # ----- Stats and Integration Calculations -----
  integrate_area <- function(time, values) {
    dt <- diff(time)
    sum(((head(values, -1) + tail(values, -1)) / 2) * dt)
  }
  
  compute_stats <- function(temp_values, time_values) {
    in_range <- (temp_values >= preferred_min) & (temp_values <= preferred_max)
    perc_in_range <- mean(in_range, na.rm = TRUE) * 100
    perc_out_total <- (1 - mean(in_range, na.rm = TRUE)) * 100
    
    # Overall integration (entire selection)
    if (length(time_values) > 1) {
      overall_excess_above <- pmax(0, temp_values - preferred_max)
      overall_excess_below <- pmax(0, preferred_min - temp_values)
      overall_area_above <- integrate_area(time_values, overall_excess_above)
      overall_area_below <- integrate_area(time_values, overall_excess_below)
      overall_area_total <- overall_area_above + overall_area_below
    } else {
      overall_area_above <- NA
      overall_area_below <- NA
      overall_area_total <- NA
    }
    
    # Post-entry integration (from first entry)
    first_entry <- which(in_range)[1]
    if (!is.na(first_entry)) {
      post_in_range <- in_range[first_entry:length(in_range)]
      perc_out_post <- (1 - mean(post_in_range, na.rm = TRUE)) * 100
      
      post_time <- time_values[first_entry:length(time_values)]
      post_temp <- temp_values[first_entry:length(temp_values)]
      post_excess_above <- pmax(0, post_temp - preferred_max)
      post_excess_below <- pmax(0, preferred_min - post_temp)
      if (length(post_time) > 1) {
        post_area_above <- integrate_area(post_time, post_excess_above)
        post_area_below <- integrate_area(post_time, post_excess_below)
      } else {
        post_area_above <- NA
        post_area_below <- NA
      }
      post_area_total <- post_area_above + post_area_below
    } else {
      perc_out_post <- NA
      post_area_above <- NA
      post_area_below <- NA
      post_area_total <- NA
    }
    
    time_above <- sum(temp_values > preferred_max, na.rm = TRUE)
    time_below <- sum(temp_values < preferred_min, na.rm = TRUE)
    rate_of_change <- diff(temp_values) / diff(time_values)
    avg_abs_rate <- mean(abs(rate_of_change), na.rm = TRUE)
    if (!is.na(first_entry)) {
      ttr <- time_values[first_entry] - time_values[1]
    } else {
      ttr <- NA
    }
    
    list(perc_in_range = perc_in_range,
         perc_out_total = perc_out_total,
         perc_out_post = perc_out_post,
         time_above = time_above,
         time_below = time_below,
         avg_abs_rate = avg_abs_rate,
         overall_area_above = overall_area_above,
         overall_area_below = overall_area_below,
         overall_area_total = overall_area_total,
         post_area_above = post_area_above,
         post_area_below = post_area_below,
         post_area_total = post_area_total,
         time_to_reach = ttr)
  }
  
  collected_stats <- reactiveVal(data.frame(
    Animal_ID = character(),
    Channel = character(),
    Max_Temp = numeric(),
    Avg_Temp = numeric(),
    Min_Temp = numeric(),
    Time_to_Reached = character(),
    Perc_In_Range = numeric(),
    Perc_Out_Total = numeric(),
    Perc_Out_Post = numeric(),
    Time_Above = numeric(),
    Time_Below = numeric(),
    Avg_Abs_Change = numeric(),
    Overall_Area_Above = numeric(),
    Overall_Area_Below = numeric(),
    Overall_Area_Total = numeric(),
    Post_Area_Above = numeric(),
    Post_Area_Below = numeric(),
    Post_Area_Total = numeric(),
    Start_Time_PST = character(),
    Start_Time_Since_Surgery = character(),
    Total_Time = character(),
    stringsAsFactors = FALSE
  ))
  
  seconds_to_hms <- function(s) {
    h <- s %/% 3600
    m <- (s %% 3600) %/% 60
    s_rem <- s %% 60
    sprintf("%02d:%02d:%02d", h, m, s_rem)
  }
  
  output$stats_B <- renderUI({ req(input$run_stats) })
  output$stats_C <- renderUI({ req(input$run_stats) })
  
  observeEvent(input$run_stats, {
    showModal(modalDialog(
      title = "Enter Animal ID",
      textInput("modal_animal_id", "Animal ID:"),
      footer = tagList(
        modalButton("Cancel"),
        actionButton("submit_animal", "Submit")
      )
    ))
  })
  
  observeEvent(input$submit_animal, {
    removeModal()
    last_animal(input$modal_animal_id)
    animal_id <- input$modal_animal_id
    range <- input$time_range
    if (input$channel_tab == "Channel B") {
      selected_data <- reactive_data_B() %>% filter(new_time >= range[1] & new_time <= range[2])
      if (nrow(selected_data) > 1) {
        temp_values <- selected_data$Channel_B
        time_values <- as.numeric(selected_data$new_time)
        stats_calc <- compute_stats(temp_values, time_values)
        start_actual <- format(min(selected_data$new_time, na.rm = TRUE), "%H:%M:%S")
        start_since_secs <- min(selected_data$data_secs, na.rm = TRUE)
        start_since_hms <- seconds_to_hms(start_since_secs)
        
        output$stats_B <- renderUI({
          HTML(paste0(
            "<hr><strong>Channel B:</strong><br>",
            "<strong>Temperature Stats:</strong><br>",
            "Max Temp: ", round(max(temp_values, na.rm = TRUE), 2), " °C<br>",
            "Avg Temp: ", round(mean(temp_values, na.rm = TRUE), 2), " °C<br>",
            "Min Temp: ", round(min(temp_values, na.rm = TRUE), 2), " °C<br><br>",
            
            "<strong>Time Stats:</strong><br>",
            "Time to Reach Preferred Range: ", 
            ifelse(is.na(stats_calc$time_to_reach), "Never Reached", paste0(round(stats_calc$time_to_reach, 1), " sec")), "<br>",
            "% Time In Preferred Range: ", round(stats_calc$perc_in_range, 1), " %<br>",
            "% Time Out-of-Range (Overall): ", round(stats_calc$perc_out_total, 1), " %<br>",
            "% Time Out-of-Range (Post-Entry): ", 
            ifelse(is.na(stats_calc$perc_out_post), "Never Entered", paste0(round(stats_calc$perc_out_post, 1), " %")), "<br>",
            "Time Above Preferred Range: ", stats_calc$time_above, " sec<br>",
            "Time Below Preferred Range: ", stats_calc$time_below, " sec<br>",
            "Avg Absolute Rate of Change: ", round(stats_calc$avg_abs_rate, 3), " °C/sec<br><br>",
            
            "<strong>Overall Integration:</strong><br>",
            "Area Above: ", round(stats_calc$overall_area_above, 3), " °C·sec<br>",
            "Area Below: ", round(stats_calc$overall_area_below, 3), " °C·sec<br>",
            "Total: ", round(stats_calc$overall_area_total, 3), " °C·sec<br><br>",
            
            "<strong>Post-Entry Integration:</strong><br>",
            "Area Above: ", round(stats_calc$post_area_above, 3), " °C·sec<br>",
            "Area Below: ", round(stats_calc$post_area_below, 3), " °C·sec<br>",
            "Total: ", round(stats_calc$post_area_total, 3), " °C·sec<br><br>",
            
            "<strong>Start Times:</strong><br>",
            "Start Time (PST): ", start_actual, "<br>",
            "Since Start of Surgery: ", start_since_hms, "<br><hr>"
          ))
        })
        new_row <- data.frame(
          Animal_ID = animal_id,
          Channel = "Channel B",
          Max_Temp = round(max(temp_values, na.rm = TRUE), 2),
          Avg_Temp = round(mean(temp_values, na.rm = TRUE), 2),
          Min_Temp = round(min(temp_values, na.rm = TRUE), 2),
          Time_to_Reached = ifelse(is.na(stats_calc$time_to_reach), "Never Reached", paste0(round(stats_calc$time_to_reach, 1), " sec")),
          Perc_In_Range = round(stats_calc$perc_in_range, 1),
          Perc_Out_Total = round(stats_calc$perc_out_total, 1),
          Perc_Out_Post = ifelse(is.na(stats_calc$perc_out_post), NA, round(stats_calc$perc_out_post, 1)),
          Time_Above = stats_calc$time_above,
          Time_Below = stats_calc$time_below,
          Avg_Abs_Change = round(stats_calc$avg_abs_rate, 3),
          Overall_Area_Above = round(stats_calc$overall_area_above, 3),
          Overall_Area_Below = round(stats_calc$overall_area_below, 3),
          Overall_Area_Total = round(stats_calc$overall_area_total, 3),
          Post_Area_Above = round(stats_calc$post_area_above, 3),
          Post_Area_Below = round(stats_calc$post_area_below, 3),
          Post_Area_Total = round(stats_calc$post_area_total, 3),
          Start_Time_PST = start_actual,
          Start_Time_Since_Surgery = start_since_hms,
          Total_Time = seconds_to_hms(max(time_values) - min(time_values)),
          stringsAsFactors = FALSE
        )
        # Save snapshot for this animal.
        snapshotFull <- tempfile(fileext = ".png")
        ggsave(snapshotFull, plot = static_plot(), width = 8, height = 4, dpi = 300)
        graphSnapshots$full[[length(graphSnapshots$full) + 1]] <- snapshotFull
        graphSnapshots$animal[[length(graphSnapshots$animal) + 1]] <- animal_id
        
        collected_stats(rbind(collected_stats(), new_row))
      } else {
        output$stats_B <- renderUI({ HTML("<hr><p>No data in the selected range for Channel B.</p>") })
      }
    }
    
    if (input$channel_tab == "Channel C") {
      selected_data <- reactive_data_C() %>% filter(new_time >= range[1] & new_time <= range[2])
      if (nrow(selected_data) > 1) {
        temp_values <- selected_data$Channel_C
        time_values <- as.numeric(selected_data$new_time)
        stats_calc <- compute_stats(temp_values, time_values)
        start_actual <- format(min(selected_data$new_time, na.rm = TRUE), "%H:%M:%S")
        start_since_secs <- min(selected_data$data_secs, na.rm = TRUE)
        start_since_hms <- seconds_to_hms(start_since_secs)
        
        output$stats_C <- renderUI({
          HTML(paste0(
            "<hr><strong>Channel C:</strong><br>",
            "<strong>Temperature Stats:</strong><br>",
            "Max Temp: ", round(max(temp_values, na.rm = TRUE), 2), " °C<br>",
            "Avg Temp: ", round(mean(temp_values, na.rm = TRUE), 2), " °C<br>",
            "Min Temp: ", round(min(temp_values, na.rm = TRUE), 2), " °C<br><br>",
            
            "<strong>Time Stats:</strong><br>",
            "Time to Reach Preferred Range: ", 
            ifelse(is.na(stats_calc$time_to_reach), "Never Reached", paste0(round(stats_calc$time_to_reach, 1), " sec")), "<br>",
            "% Time In Preferred Range: ", round(stats_calc$perc_in_range, 1), " %<br>",
            "% Time Out-of-Range (Overall): ", round(stats_calc$perc_out_total, 1), " %<br>",
            "% Time Out-of-Range (Post-Entry): ", 
            ifelse(is.na(stats_calc$perc_out_post), "Never Entered", paste0(round(stats_calc$perc_out_post, 1), " %")), "<br>",
            "Time Above Preferred Range: ", stats_calc$time_above, " sec<br>",
            "Time Below Preferred Range: ", stats_calc$time_below, " sec<br>",
            "Avg Absolute Rate of Change: ", round(stats_calc$avg_abs_rate, 3), " °C/sec<br><br>",
            
            "<strong>Overall Integration:</strong><br>",
            "Area Above: ", round(stats_calc$overall_area_above, 3), " °C·sec<br>",
            "Area Below: ", round(stats_calc$overall_area_below, 3), " °C·sec<br>",
            "Total: ", round(stats_calc$overall_area_total, 3), " °C·sec<br><br>",
            
            "<strong>Post-Entry Integration:</strong><br>",
            "Area Above: ", round(stats_calc$post_area_above, 3), " °C·sec<br>",
            "Area Below: ", round(stats_calc$post_area_below, 3), " °C·sec<br>",
            "Total: ", round(stats_calc$post_area_total, 3), " °C·sec<br><br>",
            
            "<strong>Start Times:</strong><br>",
            "Start Time (PST): ", start_actual, "<br>",
            "Since Start of Surgery: ", start_since_hms, "<br><hr>"
          ))
        })
        new_row <- data.frame(
          Animal_ID = animal_id,
          Channel = "Channel C",
          Max_Temp = round(max(temp_values, na.rm = TRUE), 2),
          Avg_Temp = round(mean(temp_values, na.rm = TRUE), 2),
          Min_Temp = round(min(temp_values, na.rm = TRUE), 2),
          Time_to_Reached = ifelse(is.na(stats_calc$time_to_reach), "Never Reached", paste0(round(stats_calc$time_to_reach, 1), " sec")),
          Perc_In_Range = round(stats_calc$perc_in_range, 1),
          Perc_Out_Total = round(stats_calc$perc_out_total, 1),
          Perc_Out_Post = ifelse(is.na(stats_calc$perc_out_post), NA, round(stats_calc$perc_out_post, 1)),
          Time_Above = stats_calc$time_above,
          Time_Below = stats_calc$time_below,
          Avg_Abs_Change = round(stats_calc$avg_abs_rate, 3),
          Overall_Area_Above = round(stats_calc$overall_area_above, 3),
          Overall_Area_Below = round(stats_calc$overall_area_below, 3),
          Overall_Area_Total = round(stats_calc$overall_area_total, 3),
          Post_Area_Above = round(stats_calc$post_area_above, 3),
          Post_Area_Below = round(stats_calc$post_area_below, 3),
          Post_Area_Total = round(stats_calc$post_area_total, 3),
          Start_Time_PST = start_actual,
          Start_Time_Since_Surgery = start_since_hms,
          Total_Time = seconds_to_hms(max(time_values) - min(time_values)),
          stringsAsFactors = FALSE
        )
        graphSnapshots$full[[length(graphSnapshots$full) + 1]] <- tempfile(fileext = ".png")
        ggsave(graphSnapshots$full[[length(graphSnapshots$full)]], plot = static_plot(), width = 8, height = 4, dpi = 300)
        graphSnapshots$animal[[length(graphSnapshots$animal) + 1]] <- animal_id
        
        collected_stats(rbind(collected_stats(), new_row))
      } else {
        output$stats_C <- renderUI({ HTML("<hr><p>No data in the selected range for Channel C.</p>") })
      }
    }
  })
  
  # Download handler for filtered data (CSV export).
  output$download_filtered <- downloadHandler(
    filename = function() {
      paste("filtered_data_", tools::file_path_sans_ext(basename(file_path)), ".csv", sep = "")
    },
    content = function(file) {
      filtered_data_to_download <- if (input$channel_tab == "Channel B") {
        reactive_data_B()
      } else {
        reactive_data_C()
      }
      write.csv(filtered_data_to_download, file, row.names = FALSE)
    }
  )
  
  # Download handler for stats with graph snapshots (Excel export).
  output$download_stats <- downloadHandler(
    filename = function() {
      paste("surgery_stats_", tools::file_path_sans_ext(basename(file_path)), ".xlsx", sep = "")
    },
    content = function(file) {
      wb <- createWorkbook()
      
      # Add Stats worksheet.
      addWorksheet(wb, "Stats")
      writeDataTable(wb, sheet = "Stats", x = collected_stats())
      
      # Add Graph worksheet.
      addWorksheet(wb, "Graph")
      
      # Iterate through saved snapshots for each animal.
      start_row <- 1
      n <- length(graphSnapshots$full)
      if(n > 0) {
        for(i in seq_len(n)) {
          animal_title <- graphSnapshots$animal[[i]]
          writeData(wb, sheet = "Graph", x = paste("Full Graph for Animal:", animal_title), startRow = start_row, startCol = 1)
          insertImage(wb, sheet = "Graph", file = graphSnapshots$full[[i]], startRow = start_row + 2, startCol = 1, width = 8, height = 4, units = "in")
          start_row <- start_row + 8   # Increment start row for the next animal.
        }
      }
      
      saveWorkbook(wb, file, overwrite = TRUE)
    }
  )
}

shinyApp(ui, server)
