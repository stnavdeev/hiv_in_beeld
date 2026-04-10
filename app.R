cat("\n========== HIV DASHBOARD APP STARTUP ==========\n")
cat(paste0("Time: ", Sys.time(), "\n"))
cat(paste0("Working directory: ", getwd(), "\n"))
cat("===============================================\n\n")
flush.console()

cat("[STARTUP] Installing and loading required packages...\n")
flush.console()

packages <- c(
  "shiny",
  "readxl",
  "dplyr",
  "tidyr",
  "ggplot2",
  "plotly",
  "purrr",
  "stringr",
  "tibble",
  "DT",
  "scales"
)

new_packages <- packages[!(packages %in% installed.packages()[, "Package"])]
if (length(new_packages) > 0) {
  cat(sprintf("[STARTUP] Installing missing packages: %s\n", paste(new_packages, collapse = ", ")))
  flush.console()
  install.packages(
    new_packages,
    lib = Sys.getenv("R_LIBS_USER"),
    repos = "https://cran.r-project.org"
  )
}

invisible(lapply(packages, library, character.only = TRUE))
cat("[STARTUP] Packages ready.\n")
flush.console()

if (getRversion() >= "2.15.1") {
  utils::globalVariables(c(
    "sheet", "dataset", "population", "group_var", "outcome",
    "years_since_diagnosis", "year", "has_hiv", "value", "type",
    "coef", "lo", "hi", "group", "factor", "category", "level", "share",
    "N", "series", "x_value", "has_hiv_label", "group_value",
    "dataset_label", "pretty_variable", "outcome_label", "factor_label",
    "avg_effect_label", "display_name", "value_nominal",
    "value_adjusted_2023", "value_plot", "cpi"
  ))
}

data_path <- dplyr::case_when(
  file.exists("data/output.xlsx") ~ "data/output.xlsx",
  file.exists("output.xlsx") ~ "output.xlsx",
  TRUE ~ "data/output.xlsx"
)
log_file <- "hiv_dashboard_log.txt"
unlink(log_file)

log_msg <- function(msg) {
  line <- paste0("[", Sys.time(), "] ", msg)
  cat(line, "\n", file = log_file, append = TRUE)
  cat(line, "\n")
  flush.console()
}

safe_read_sheet <- function(path, sheet) {
  tryCatch(
    readxl::read_excel(path, sheet = sheet, guess_max = 100000),
    error = function(e) {
      log_msg(sprintf("[safe_read_sheet] Failed to read %s: %s", sheet, e$message))
      tibble::tibble()
    }
  )
}

infer_dataset <- function(sheet_name) {
  if (grepl("_es_group$", sheet_name)) {
    sub("_es_group$", "", sheet_name)
  } else if (grepl("_es$", sheet_name)) {
    sub("_es$", "", sheet_name)
  } else if (grepl("_mean", sheet_name)) {
    sub("_mean.*$", "", sheet_name)
  } else {
    NA_character_
  }
}

prettify_label <- function(x) {
  x |>
    stringr::str_replace_all("_", " ") |>
    stringr::str_replace_all("\\s+", " ") |>
    stringr::str_trim() |>
    stringr::str_to_title()
}

prettify_factor <- function(x) {
  dplyr::recode(
    x,
    geslacht = "Sex",
    migratie_achtergrond = "Migration background",
    hiv_stage = "HIV stage",
    leeftijd_cat = "Age category",
    burgstaat = "Marital status",
    typehh = "Household type",
    ggd = "GGD",
    hgopl = "Education",
    .default = prettify_label(x)
  )
}

prettify_hiv <- function(x) {
  dplyr::recode(
    as.character(x),
    `0` = "Matched controls / no HIV",
    `1` = "People with HIV",
    .default = as.character(x)
  )
}

cpi_index <- tibble::tibble(
  year = 2014:2024,
  cpi = c(99.40, 100.00, 100.32, 101.70, 103.44, 106.16,
          107.51, 110.39, 121.43, 126.09, 130.31)
)
base_2023 <- cpi_index$cpi[cpi_index$year == 2023]
cpi_index <- cpi_index |>
  dplyr::mutate(cpi = cpi / base_2023)

is_cost_variable <- function(x) {
  x_chr <- as.character(x)
  !is.na(x_chr) &
    (
      stringr::str_detect(x_chr, "^zvwk") |
        stringr::str_detect(x_chr, "^costs_")
    ) &
    !stringr::str_detect(x_chr, "^used_")
}

sheet_names <- readxl::excel_sheets(data_path)
log_msg(sprintf("[startup] Found %d sheets", length(sheet_names)))

sheet_preview <- purrr::map_dfr(sheet_names, function(s) {
  df <- safe_read_sheet(data_path, s)
  cols_sorted <- sort(names(df))
  tibble::tibble(
    sheet = s,
    dataset = infer_dataset(s),
    n_rows = nrow(df),
    cols = paste(names(df), collapse = ", "),
    has_value = all(c("value", "variable", "type") %in% names(df)),
    has_es = all(c("outcome", "coef", "lo", "hi", "years_since_diagnosis") %in% names(df)),
    has_es_group = all(c("outcome", "factor", "group", "coef", "lo", "hi", "years_since_diagnosis") %in% names(df)),
    has_profile = identical(cols_sorted, sort(c("has_hiv", "level", "N", "category", "share"))),
    has_shm_total = identical(s, "shm_total") || ("year" %in% names(df) && sum(grepl("^N_|^n_", names(df))) >= 1)
  )
})

means_index <- sheet_preview |>
  dplyr::filter(has_value, !sheet %in% c("codes", "matching_stats")) |>
  dplyr::filter(
    !stringr::str_detect(sheet, "_mean_yr$") |
      stringr::str_detect(sheet, "_mean_yr.*_all$")
  ) |>
  dplyr::mutate(
    dataset_label = prettify_label(dataset),
    time_scale = dplyr::case_when(
      stringr::str_detect(sheet, "_mean_yr") ~ "Calendar year",
      TRUE ~ "Years since diagnosis"
    ),
    group_var = dplyr::case_when(
      stringr::str_detect(sheet, "_mean_yr_gesl_all$|_mean_gesl(_all)?$") ~ "geslacht",
      stringr::str_detect(sheet, "_mean_yr_migr_all$|_mean_migr(_all)?$") ~ "migratie_achtergrond",
      stringr::str_detect(sheet, "_mean_yr_hiv.*_all$|_mean_hiv.*(_all)?$") ~ "hiv_stage",
      stringr::str_detect(sheet, "_mean_yr_leef_all$|_mean_leef(_all)?$") ~ "leeftijd_cat",
      stringr::str_detect(cols, "geslacht") ~ "geslacht",
      stringr::str_detect(cols, "migratie_achtergrond") ~ "migratie_achtergrond",
      stringr::str_detect(cols, "hiv_stage") ~ "hiv_stage",
      stringr::str_detect(cols, "leeftijd_cat") ~ "leeftijd_cat",
      TRUE ~ "none"
    ),
    group_label = dplyr::case_when(
      group_var == "none" ~ "No subgroup split",
      TRUE ~ prettify_factor(group_var)
    )
  )

es_index <- sheet_preview |>
  dplyr::filter(has_es, !has_es_group, !sheet %in% c("codes", "matching_stats")) |>
  dplyr::mutate(dataset_label = prettify_label(dataset))

es_group_index <- sheet_preview |>
  dplyr::filter(has_es_group, !sheet %in% c("codes", "matching_stats")) |>
  dplyr::mutate(dataset_label = prettify_label(dataset))

profile_sheet_name <- dplyr::first(sheet_preview$sheet[sheet_preview$has_profile])
shm_sheet_name <- dplyr::first(sheet_preview$sheet[sheet_preview$has_shm_total & sheet_preview$sheet != "codes" & sheet_preview$sheet != "matching_stats"])

cache_env <- new.env(parent = emptyenv())
get_sheet <- function(sheet_name) {
  req(!is.na(sheet_name), nzchar(sheet_name))
  key <- paste0("sheet__", sheet_name)
  if (!exists(key, envir = cache_env, inherits = FALSE)) {
    df <- safe_read_sheet(data_path, sheet_name)
    assign(key, df, envir = cache_env)
    log_msg(sprintf("[cache] Loaded %s (%d rows)", sheet_name, nrow(df)))
  }
  get(key, envir = cache_env, inherits = FALSE)
}

ui <- navbarPage(
  title = "Understanding HIV care in the Netherlands",
  id = "main_nav",
  header = tags$div(
    style = "padding: 12px 18px 4px 18px; color: #4b5563; font-size: 14px;",
    "Explore healthcare costs, medication use, profiles, and event-study estimates for people with HIV and matched controls."
  ),
  
  tabPanel(
    "Overview",
    fluidPage(
      br(),
      fluidRow(
        column(
          12,
          h4("Totals"),
          plotlyOutput("plot_shm_total", height = "460px"),
          DTOutput("tbl_shm_total")
        )
      )
    )
  ),
  
  tabPanel(
    "Profiles",
    sidebarLayout(
      sidebarPanel(
        selectInput("profile_category", "Category", choices = NULL),
        downloadButton("dl_profile", "Download filtered profile data")
      ),
      mainPanel(
        plotlyOutput("plot_profile", height = "820px"),
        DTOutput("tbl_profile")
      )
    )
  ),
  
  tabPanel(
    "Aggregated means",
    sidebarLayout(
      sidebarPanel(
        selectInput("mean_dataset", "Dataset", choices = NULL),
        radioButtons("mean_time_scale", "Time scale", choices = c("Years since diagnosis", "Calendar year")),
        selectInput("mean_group_var", "Subgroup split", choices = NULL),
        selectInput("mean_variable", "Outcome variable", choices = NULL),
        selectInput("mean_type", "Statistic type", choices = NULL),
        uiOutput("mean_inflation_ui"),
        uiOutput("mean_group_filter_ui"),
        downloadButton("dl_mean", "Download filtered mean data")
      ),
      mainPanel(
        plotlyOutput("plot_mean", height = "660px"),
        DTOutput("tbl_mean")
      )
    )
  ),
  
  tabPanel(
    "Event study",
    sidebarLayout(
      sidebarPanel(
        selectInput("es_dataset", "Dataset", choices = NULL),
        selectInput("es_outcome", "Outcome", choices = NULL),
        downloadButton("dl_es", "Download filtered event-study data")
      ),
      mainPanel(
        uiOutput("es_avg_effect"),
        plotlyOutput("plot_es", height = "660px"),
        DTOutput("tbl_es_summary"),
        DTOutput("tbl_es")
      )
    )
  ),
  
  tabPanel(
    "Event study by group",
    sidebarLayout(
      sidebarPanel(
        selectInput("esg_dataset", "Dataset", choices = NULL),
        selectInput("esg_factor", "Grouping factor", choices = NULL),
        selectInput("esg_outcome", "Outcome", choices = NULL),
        selectizeInput("esg_groups", "Visible groups", choices = NULL, multiple = TRUE),
        downloadButton("dl_esg", "Download filtered grouped event-study data")
      ),
      mainPanel(
        uiOutput("esg_avg_effect"),
        plotlyOutput("plot_esg", height = "660px"),
        DTOutput("tbl_esg_summary"),
        DTOutput("tbl_esg")
      )
    )
  )
)

server <- function(input, output, session) {
  error_log <- reactiveVal(character())
  add_error <- function(msg) {
    log_msg(msg)
    error_log(c(error_log(), msg))
  }
  
  shm_total_df <- reactive({
    req(!is.na(shm_sheet_name))
    get_sheet(shm_sheet_name)
  })
  
  profile_df <- reactive({
    req(!is.na(profile_sheet_name))
    get_sheet(profile_sheet_name)
  })
  
  observe({
    prof <- profile_df()
    if (nrow(prof) > 0) {
      cats <- sort(unique(prof$category))
      updateSelectInput(session, "profile_category", choices = cats, selected = cats[1])
    }
  })
  
  output$tbl_shm_total <- renderDT({
    DT::datatable(shm_total_df(), options = list(pageLength = 10, scrollX = TRUE))
  })
  
  output$plot_shm_total <- renderPlotly({
    df <- shm_total_df()
    req(nrow(df) > 0, "year" %in% names(df))
    value_cols <- setdiff(names(df), "year")
    df_long <- df |>
      tidyr::pivot_longer(dplyr::all_of(value_cols), names_to = "series", values_to = "value_raw") |>
      dplyr::mutate(
        year = suppressWarnings(as.numeric(year)),
        value = suppressWarnings(as.numeric(value_raw))
      ) |>
      dplyr::filter(!is.na(value), !is.na(year)) |>
      dplyr::mutate(
        tooltip = paste0(
          "Series: ", series, "<br>",
          "Year: ", year, "<br>",
          "Value: ", scales::comma(value)
        )
      )
    
    req(nrow(df_long) > 0)
    
    p <- plotly::plot_ly()
    for (s in unique(df_long$series)) {
      trace_df <- df_long[df_long$series == s, , drop = FALSE]
      p <- p |>
        plotly::add_trace(
          data = trace_df,
          x = ~year,
          y = ~value,
          type = "scatter",
          mode = "lines+markers",
          name = s,
          text = ~tooltip,
          hoverinfo = "text"
        )
    }
    
    p |>
      plotly::layout(
        title = list(text = "Totals over time"),
        xaxis = list(title = ""),
        yaxis = list(title = "Count"),
        legend = list(title = list(text = ""))
      )
  })
  
  filtered_profile <- reactive({
    df <- profile_df()
    req(nrow(df) > 0, input$profile_category)
    df |>
      dplyr::filter(category == input$profile_category) |>
      dplyr::mutate(
        has_hiv_label = prettify_hiv(has_hiv),
        level = as.character(level)
      )
  })
  
  output$plot_profile <- renderPlotly({
    df <- filtered_profile()
    req(nrow(df) > 0)
    
    df_plot <- df |>
      dplyr::filter(!is.na(share), share <= 1, !tolower(level) %in% c("all", "unknown"))
    if (nrow(df_plot) == 0) {
      df_plot <- df |>
        dplyr::filter(!is.na(share), share <= 1)
    }
    metric_col <- "share"
    metric_lab <- "Share"
    df_plot <- df_plot |>
      dplyr::mutate(
        metric_value = share,
        value_label = scales::percent(share, accuracy = 1)
      )
    
    req(nrow(df_plot) > 0)
    
    level_order <- df_plot |>
      dplyr::group_by(level) |>
      dplyr::summarise(order_value = max(metric_value, na.rm = TRUE), .groups = "drop") |>
      dplyr::arrange(order_value) |>
      dplyr::pull(level)
    
    df_plot <- df_plot |>
      dplyr::mutate(
        level = factor(stringr::str_wrap(level, width = 28),
                       levels = stringr::str_wrap(level_order, width = 28)),
        tooltip = paste0(
          "Level: ", gsub("\n", " ", as.character(level)), "<br>",
          "Population: ", has_hiv_label, "<br>",
          metric_lab, ": ", value_label
        )
      )
    
    p <- plotly::plot_ly()
    for (grp in unique(df_plot$has_hiv_label)) {
      trace_df <- df_plot[df_plot$has_hiv_label == grp, , drop = FALSE]
      p <- p |>
        plotly::add_trace(
          data = trace_df,
          x = ~metric_value,
          y = ~level,
          type = "bar",
          orientation = "h",
          name = grp,
          text = ~value_label,
          textposition = "auto",
          customdata = ~tooltip,
          hovertemplate = "%{customdata}<extra></extra>"
        )
    }
    
    p |>
      plotly::layout(
        barmode = "group",
        title = list(text = paste("Profile:", prettify_factor(input$profile_category))),
        xaxis = list(title = metric_lab, tickformat = if (metric_col == "share") ",.0%" else NULL),
        yaxis = list(title = ""),
        legend = list(orientation = "h", x = 0, y = -0.12)
      )
  })
  
  output$tbl_profile <- renderDT({
    DT::datatable(filtered_profile(), options = list(pageLength = 15, scrollX = TRUE))
  })
  
  observe({
    ds_choices <- sort(unique(means_index$dataset_label))
    if (length(ds_choices) > 0) {
      selected <- isolate(input$mean_dataset)
      if (is.null(selected) || !(selected %in% ds_choices)) selected <- ds_choices[1]
      updateSelectInput(session, "mean_dataset", choices = ds_choices, selected = selected)
    }
  })
  
  mean_sheet_choice <- reactive({
    req(input$mean_dataset, input$mean_time_scale)
    
    candidates <- means_index |>
      dplyr::filter(
        dataset_label == input$mean_dataset,
        time_scale == input$mean_time_scale
      )
    
    req(nrow(candidates) > 0)
    candidates
  })
  
  observe({
    candidates <- mean_sheet_choice()
    choices <- unique(candidates$group_label)
    
    choices <- c(
      "No subgroup split",
      sort(setdiff(choices, "No subgroup split"))
    )
    choices <- unique(choices[choices %in% candidates$group_label])
    
    selected <- isolate(input$mean_group_var)
    if (is.null(selected) || !(selected %in% choices)) selected <- choices[1]
    
    updateSelectInput(session, "mean_group_var", choices = choices, selected = selected)
  })
  
  mean_sheet_selected <- reactive({
    candidates <- mean_sheet_choice()
    req(input$mean_group_var)
    
    rows <- candidates |>
      dplyr::filter(group_label == input$mean_group_var)
    
    if (nrow(rows) == 0) {
      rows <- candidates
    }
    
    if (identical(input$mean_time_scale, "Calendar year")) {
      if (identical(input$mean_group_var, "No subgroup split")) {
        preferred <- rows |>
          dplyr::filter(stringr::str_detect(sheet, "_mean_yr_all$"))
      } else {
        preferred <- rows |>
          dplyr::filter(stringr::str_detect(sheet, "_mean_yr_.*_all$"))
      }
      
      if (nrow(preferred) > 0) {
        rows <- preferred
      }
    }
    
    row <- rows |>
      dplyr::arrange(sheet) |>
      dplyr::slice(1)
    
    req(nrow(row) == 1)
    row
  })
  
  mean_data_raw <- reactive({
    get_sheet(mean_sheet_selected()$sheet[[1]])
  })
  
  observe({
    df <- mean_data_raw()
    req(nrow(df) > 0)
    var_choices <- sort(unique(df$name))
    type_choices <- sort(unique(df$type))
    
    selected_var <- isolate(input$mean_variable)
    if (is.null(selected_var) || !(selected_var %in% var_choices)) selected_var <- var_choices[1]
    
    selected_type <- isolate(input$mean_type)
    default_type <- dplyr::coalesce(type_choices[type_choices == "gemiddelde_per_persoon"][1], type_choices[1])
    if (is.null(selected_type) || !(selected_type %in% type_choices)) selected_type <- default_type
    
    updateSelectInput(session, "mean_variable", choices = var_choices, selected = selected_var)
    updateSelectInput(
      session,
      "mean_type",
      choices = type_choices,
      selected = selected_type
    )
  })

  output$mean_inflation_ui <- renderUI({
    req(input$mean_time_scale, input$mean_variable)

    if (!identical(input$mean_time_scale, "Calendar year") || !isTRUE(is_cost_variable(input$mean_variable))) {
      return(NULL)
    }

    checkboxInput(
      "mean_adjust_inflation",
      "Adjust costs to 2023 EUR using CPI",
      value = FALSE
    )
  })
  
  output$mean_group_filter_ui <- renderUI({
    df <- mean_data_raw()
    req(nrow(df) > 0)
    group_var <- mean_sheet_selected()$group_var[[1]]
    if (group_var == "none" || !(group_var %in% names(df))) {
      return(NULL)
    }
    choices <- sort(unique(as.character(df[[group_var]])))
    selected <- isolate(input$mean_group_values)
    if (is.null(selected) || length(selected) == 0) {
      selected <- choices
    } else {
      selected <- intersect(selected, choices)
      if (length(selected) == 0) selected <- choices
    }
    selectizeInput("mean_group_values", prettify_factor(group_var), choices = choices, selected = selected, multiple = TRUE)
  })
  
  filtered_mean <- reactive({
    df <- mean_data_raw()
    req(nrow(df) > 0, input$mean_variable, input$mean_type)
    
    df <- df |>
      dplyr::filter(name == input$mean_variable, type == input$mean_type) |>
      dplyr::mutate(
        has_hiv_label = prettify_hiv(has_hiv),
        pretty_variable = prettify_label(name),
        display_name = as.character(name)
      )
    
    group_var <- mean_sheet_selected()$group_var[[1]]
    x_var <- if ("years_since_diagnosis" %in% names(df)) "years_since_diagnosis" else "year"
    
    if (group_var != "none" && group_var %in% names(df)) {
      selected_groups <- input$mean_group_values
      if (!is.null(selected_groups) && length(selected_groups) > 0) {
        df <- df |>
          dplyr::filter(as.character(.data[[group_var]]) %in% selected_groups)
      }
      df <- df |>
        dplyr::mutate(group_value = as.character(.data[[group_var]]))
    } else {
      df <- df |>
        dplyr::mutate(group_value = "All")
    }
    
    df <- df |>
      dplyr::mutate(
        x_value = suppressWarnings(as.numeric(.data[[x_var]]))
      ) |>
      dplyr::filter(!is.na(x_value), !is.na(value))

    if (identical(x_var, "year")) {
      df <- df |>
        dplyr::left_join(cpi_index, by = "year")
    } else {
      df <- df |>
        dplyr::mutate(cpi = NA_real_)
    }

    adjust_for_inflation <- identical(input$mean_time_scale, "Calendar year") &&
      isTRUE(is_cost_variable(input$mean_variable)) &&
      isTRUE(input$mean_adjust_inflation)

    df <- df |>
      dplyr::mutate(
        value_nominal = value,
        value_adjusted_2023 = dplyr::if_else(
          !is.na(cpi) & cpi > 0,
          value / cpi,
          value
        )
      )

    if (adjust_for_inflation) {
      df <- df |>
        dplyr::mutate(
          value_plot = value_adjusted_2023,
          value_tooltip = paste0(
            "Value (2023 EUR): ", scales::comma(value_plot),
            "<br>Nominal value: ", scales::comma(value_nominal)
          )
        )
    } else {
      df <- df |>
        dplyr::mutate(
          value_plot = value_nominal,
          value_tooltip = paste0("Value: ", scales::comma(value_plot))
        )
    }

    df
  })
  
  output$plot_mean <- renderPlotly({
    df <- filtered_mean()
    req(nrow(df) > 0)
    
    x_var <- if ("years_since_diagnosis" %in% names(df)) "years_since_diagnosis" else "year"
    group_var <- mean_sheet_selected()$group_var[[1]]
    facet_formula <- if (group_var != "none") ~ group_value else NULL
    adjust_for_inflation <- identical(input$mean_time_scale, "Calendar year") &&
      isTRUE(is_cost_variable(input$mean_variable)) &&
      isTRUE(input$mean_adjust_inflation)
    palette_vals <- c(
      "Matched controls / no HIV" = "gray50",
      "People with HIV" = "steelblue4"
    )
    
    p <- ggplot(
      df,
      aes(
        x = x_value,
        y = value_plot,
        color = has_hiv_label,
        group = has_hiv_label,
        text = paste0(
          "Outcome: ", pretty_variable, "<br>",
          "Series: ", has_hiv_label, "<br>",
          ifelse(x_var == "year", "Year: ", "Years since diagnosis: "), x_value, "<br>",
          value_tooltip
        )
      )
    ) +
      geom_line(linewidth = 1, linetype = "solid") +
      geom_point(size = 2) +
      scale_color_manual(values = palette_vals, drop = FALSE) +
      theme_minimal(base_size = 13) +
      labs(
        title = paste(
          "Aggregated means:",
          unique(df$pretty_variable),
          if (adjust_for_inflation) "(2023 EUR)" else ""
        ),
        subtitle = paste(input$mean_dataset, "-", input$mean_group_var),
        x = ifelse(x_var == "year", "Calendar year", "Years since diagnosis"),
        y = if (adjust_for_inflation) "Value (2023 EUR)" else "Value",
        color = NULL
      )
    
    if (!is.null(facet_formula)) {
      p <- p + facet_wrap(facet_formula, scales = "fixed", ncol = 1)
    }
    
    if (x_var == "years_since_diagnosis") {
      p <- p + geom_vline(xintercept = 0, linetype = "dashed", color = "red")
    }
    
    plotly::ggplotly(p, tooltip = "text")
  })
  
  output$tbl_mean <- renderDT({
    DT::datatable(filtered_mean(), options = list(pageLength = 15, scrollX = TRUE))
  })
  
  observe({
    ds_choices <- sort(unique(es_index$dataset_label))
    if (length(ds_choices) > 0) {
      updateSelectInput(session, "es_dataset", choices = ds_choices, selected = ds_choices[1])
    }
  })
  
  es_sheet_selected <- reactive({
    req(input$es_dataset)
    row <- es_index |>
      dplyr::filter(dataset_label == input$es_dataset) |>
      dplyr::slice(1)
    req(nrow(row) == 1)
    row
  })
  
  es_data_raw <- reactive({
    get_sheet(es_sheet_selected()$sheet[[1]])
  })
  
  observe({
    df <- es_data_raw()
    req(nrow(df) > 0)
    outc <- sort(unique(df$outcome))
    updateSelectInput(session, "es_outcome", choices = outc, selected = outc[1])
  })
  
  filtered_es <- reactive({
    es_data_raw() |>
      dplyr::filter(outcome == input$es_outcome) |>
      dplyr::mutate(
        outcome_label = prettify_label(outcome),
        avg_effect_label = dplyr::case_when(
          !is.na(total_effect) ~ sprintf("%.3f", total_effect),
          TRUE ~ "NA"
        )
      ) |>
      dplyr::arrange(years_since_diagnosis)
  })
  
  es_summary <- reactive({
    df <- filtered_es()
    req(nrow(df) > 0)
    df |>
      dplyr::summarise(
        outcome = dplyr::first(outcome_label),
        average_post_treatment_effect = dplyr::first(total_effect),
        average_post_treatment_se = dplyr::first(total_se),
        n = dplyr::first(n)
      )
  })
  
  output$es_avg_effect <- renderUI({
    sm <- es_summary()
    req(nrow(sm) > 0)
    effect_txt <- ifelse(is.na(sm$average_post_treatment_effect[[1]]), "NA", sprintf("%.3f", sm$average_post_treatment_effect[[1]]))
    se_txt <- ifelse(is.na(sm$average_post_treatment_se[[1]]), "NA", sprintf("%.3f", sm$average_post_treatment_se[[1]]))
    HTML(sprintf("<div style='margin-bottom:10px;'><b>Average post-treatment effect:</b> %s &nbsp;&nbsp; <b>SE:</b> %s</div>", effect_txt, se_txt))
  })
  
  output$plot_es <- renderPlotly({
    df <- filtered_es()
    sm <- es_summary()
    req(nrow(df) > 0, nrow(sm) > 0)
    
    subtitle_txt <- sprintf(
      "Average post-treatment effect: %s | SE: %s",
      ifelse(is.na(sm$average_post_treatment_effect[[1]]), "NA", sprintf("%.3f", sm$average_post_treatment_effect[[1]])),
      ifelse(is.na(sm$average_post_treatment_se[[1]]), "NA", sprintf("%.3f", sm$average_post_treatment_se[[1]]))
    )
    
    p <- ggplot(
      df,
      aes(
        x = years_since_diagnosis,
        y = coef,
        text = paste0(
          "Year: ", years_since_diagnosis, "<br>",
          "Estimate: ", round(coef, 3), "<br>",
          "95% CI: [", round(lo, 3), ", ", round(hi, 3), "]<br>",
          "Average post-treatment effect: ",
          ifelse(is.na(total_effect), "NA", sprintf("%.3f", total_effect)),
          "<br>Average post-treatment SE: ",
          ifelse(is.na(total_se), "NA", sprintf("%.3f", total_se))
        )
      )
    ) +
      geom_errorbar(aes(ymin = lo, ymax = hi), width = 0.1, color = "steelblue4") +
      geom_line(linewidth = 1, linetype = "dashed", color = "steelblue4") +
      geom_point(size = 2, color = "steelblue4") +
      geom_hline(yintercept = 0, linetype = "dashed", color = "gray40") +
      geom_vline(xintercept = 0, linetype = "dashed", color = "red") +
      scale_x_continuous(breaks = sort(unique(df$years_since_diagnosis))) +
      theme_minimal(base_size = 13) +
      labs(
        title = paste("Event study:", unique(df$outcome_label)),
        subtitle = subtitle_txt,
        x = "Years since diagnosis",
        y = "Change relative to pre-diagnosis"
      )
    
    plotly::ggplotly(p, tooltip = "text")
  })
  
  output$tbl_es_summary <- renderDT({
    sm <- es_summary() |>
      dplyr::mutate(
        dplyr::across(c(average_post_treatment_effect, average_post_treatment_se), ~ round(.x, 3))
      )
    DT::datatable(sm, options = list(dom = 't', scrollX = TRUE), rownames = FALSE)
  })
  
  output$tbl_es <- renderDT({
    DT::datatable(filtered_es(), options = list(pageLength = 15, scrollX = TRUE))
  })
  
  observe({
    ds_choices <- sort(unique(es_group_index$dataset_label))
    if (length(ds_choices) > 0) {
      updateSelectInput(session, "esg_dataset", choices = ds_choices, selected = ds_choices[1])
    }
  })
  
  esg_data_raw <- reactive({
    req(input$esg_dataset)
    sheet <- es_group_index |>
      dplyr::filter(dataset_label == input$esg_dataset) |>
      dplyr::slice(1) |>
      dplyr::pull(sheet)
    req(length(sheet) == 1)
    get_sheet(sheet)
  })
  
  observe({
    df <- esg_data_raw()
    req(nrow(df) > 0)
    
    facs <- sort(unique(df$factor))
    selected <- isolate(input$esg_factor)
    if (is.null(selected) || !(selected %in% facs)) selected <- facs[1]
    
    updateSelectInput(
      session,
      "esg_factor",
      choices = stats::setNames(facs, prettify_factor(facs)),
      selected = selected
    )
  })
  
  observe({
    df <- esg_data_raw()
    req(nrow(df) > 0, input$esg_factor)
    outc <- df |>
      dplyr::filter(factor == input$esg_factor) |>
      dplyr::pull(outcome) |>
      unique() |>
      sort()
    updateSelectInput(session, "esg_outcome", choices = outc, selected = outc[1])
  })
  
  observe({
    df <- esg_data_raw()
    req(nrow(df) > 0, input$esg_factor, input$esg_outcome)
    groups <- df |>
      dplyr::filter(factor == input$esg_factor, outcome == input$esg_outcome) |>
      dplyr::pull(group) |>
      unique() |>
      sort()
    updateSelectizeInput(session, "esg_groups", choices = groups, selected = groups, server = TRUE)
  })
  
  filtered_esg <- reactive({
    df <- esg_data_raw()
    req(nrow(df) > 0, input$esg_factor, input$esg_outcome)
    df <- df |>
      dplyr::filter(factor == input$esg_factor, outcome == input$esg_outcome)
    if (!is.null(input$esg_groups) && length(input$esg_groups) > 0) {
      df <- df |>
        dplyr::filter(group %in% input$esg_groups)
    }
    df |>
      dplyr::arrange(group, years_since_diagnosis) |>
      dplyr::mutate(
        factor_label = prettify_factor(factor),
        outcome_label = prettify_label(outcome)
      )
  })
  
  esg_summary <- reactive({
    df <- filtered_esg()
    req(nrow(df) > 0)
    df |>
      dplyr::group_by(group) |>
      dplyr::summarise(
        average_post_treatment_effect = dplyr::first(total_effect),
        average_post_treatment_se = dplyr::first(total_se),
        n = dplyr::first(n),
        .groups = "drop"
      ) |>
      dplyr::arrange(group)
  })
  
  output$esg_avg_effect <- renderUI({
    sm <- esg_summary()
    req(nrow(sm) > 0)
    summary_txt <- paste(
      paste0(
        sm$group, ": ",
        ifelse(is.na(sm$average_post_treatment_effect), "NA", sprintf("%.3f", sm$average_post_treatment_effect)),
        " (SE ",
        ifelse(is.na(sm$average_post_treatment_se), "NA", sprintf("%.3f", sm$average_post_treatment_se)),
        ")"
      ),
      collapse = " | "
    )
    HTML(sprintf("<div style='margin-bottom:10px;'><b>Average post-treatment effect by group:</b> %s</div>", summary_txt))
  })
  
  output$plot_esg <- renderPlotly({
    df <- filtered_esg()
    sm <- esg_summary()
    req(nrow(df) > 0, nrow(sm) > 0)
    
    subtitle_txt <- paste0(
      unique(df$factor_label),
      " | Avg post-treatment effect: ",
      paste(
        paste0(
          sm$group, "=",
          ifelse(is.na(sm$average_post_treatment_effect), "NA", sprintf("%.3f", sm$average_post_treatment_effect)),
          " (SE ",
          ifelse(is.na(sm$average_post_treatment_se), "NA", sprintf("%.3f", sm$average_post_treatment_se)),
          ")"
        ),
        collapse = "; "
      )
    )
    
    p <- ggplot(
      df,
      aes(
        x = years_since_diagnosis,
        y = coef,
        color = group,
        group = group,
        text = paste0(
          "Group: ", group, "<br>",
          "Year: ", years_since_diagnosis, "<br>",
          "Estimate: ", round(coef, 3), "<br>",
          "95% CI: [", round(lo, 3), ", ", round(hi, 3), "]<br>",
          "Average post-treatment effect: ",
          ifelse(is.na(total_effect), "NA", sprintf("%.3f", total_effect)),
          "<br>Average post-treatment SE: ",
          ifelse(is.na(total_se), "NA", sprintf("%.3f", total_se))
        )
      )
    ) +
      geom_errorbar(aes(ymin = lo, ymax = hi), width = 0.1) +
      geom_line(linewidth = 1, linetype = "dashed") +
      geom_point(size = 2) +
      geom_hline(yintercept = 0, linetype = "dashed", color = "gray40") +
      geom_vline(xintercept = 0, linetype = "dashed", color = "red") +
      scale_x_continuous(breaks = sort(unique(df$years_since_diagnosis))) +
      theme_minimal(base_size = 13) +
      labs(
        title = paste("Event study by group:", unique(df$outcome_label)),
        subtitle = subtitle_txt,
        x = "Years since diagnosis",
        y = "Change relative to pre-diagnosis",
        color = NULL
      )
    
    plotly::ggplotly(p, tooltip = "text")
  })
  
  output$tbl_esg_summary <- renderDT({
    sm <- esg_summary() |>
      dplyr::mutate(
        dplyr::across(c(average_post_treatment_effect, average_post_treatment_se), ~ round(.x, 3))
      )
    DT::datatable(sm, options = list(dom = 't', scrollX = TRUE), rownames = FALSE)
  })
  
  output$tbl_esg <- renderDT({
    DT::datatable(filtered_esg(), options = list(pageLength = 15, scrollX = TRUE))
  })
  
  output$dl_profile <- downloadHandler(
    filename = function() paste0("profile_", input$profile_category, "_", Sys.Date(), ".csv"),
    content = function(file) write.csv(filtered_profile(), file, row.names = FALSE)
  )
  
  output$dl_mean <- downloadHandler(
    filename = function() paste0("aggregated_means_", Sys.Date(), ".csv"),
    content = function(file) write.csv(filtered_mean(), file, row.names = FALSE)
  )
  
  output$dl_es <- downloadHandler(
    filename = function() paste0("event_study_", Sys.Date(), ".csv"),
    content = function(file) write.csv(filtered_es(), file, row.names = FALSE)
  )
  
  output$dl_esg <- downloadHandler(
    filename = function() paste0("event_study_by_group_", Sys.Date(), ".csv"),
    content = function(file) write.csv(filtered_esg(), file, row.names = FALSE)
  )
}

options(shiny.error = function() {
  err <- geterrmessage()
  message(sprintf("[shiny.error] %s", err))
  writeLines(sprintf("[shiny.error] %s", err), con = "shiny_error.log")
})

if (exists("secure_app", mode = "function") && exists("secure_server", mode = "function")) {
  shinyApp(ui = secure_app(ui), server = server)
} else {
  shinyApp(ui, server)
}
