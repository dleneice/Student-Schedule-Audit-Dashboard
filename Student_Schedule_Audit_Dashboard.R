# Student Schedule Audit Dashboard
# Standalone Shiny application: all required dashboard code is in this file.

library(shiny)
library(DT)
library(dplyr)
library(openxlsx)

options(shiny.maxRequestSize = 75 * 1024^2)

# =============================================================================
# HELPER FUNCTIONS
# =============================================================================

clean_text <- function(x) {
  x <- trimws(as.character(x))
  x[x == ""] <- NA_character_
  x
}

normalize_key <- function(x) toupper(clean_text(x))

has_choice <- function(x) {
  !is.null(x) && length(x) == 1 && !is.na(x) && nzchar(x)
}

first_nonblank <- function(x) {
  x <- clean_text(x)
  x <- x[!is.na(x)]
  if (length(x) == 0) NA_character_ else x[[1]]
}

paste_nonblank <- function(x) {
  x <- sort(unique(clean_text(x)))
  x <- x[!is.na(x)]
  paste(x, collapse = ", ")
}

read_csv_as_text <- function(path) {
  read.csv(
    path,
    header = TRUE,
    sep = ",",
    quote = "\"",
    colClasses = "character",
    na.strings = c("", "NA"),
    check.names = FALSE,
    strip.white = FALSE,
    fileEncoding = "UTF-8-BOM"
  )
}

guess_column <- function(columns, patterns, fallback = columns[[1]]) {
  normalized <- tolower(gsub("[^a-z0-9]", "", columns))
  for (pattern in patterns) {
    hit <- which(grepl(pattern, normalized, fixed = TRUE))
    if (length(hit) > 0) return(columns[[hit[[1]]]])
  }
  fallback
}

parse_list <- function(value) {
  items <- trimws(unlist(strsplit(value, ",", fixed = TRUE)))
  unique(items[nzchar(items)])
}

normalize_subject <- function(x) {
  x <- tolower(trimws(x))
  if (x %in% c("english", "ela")) return("English")
  if (x %in% c("math", "mathematics")) return("Math")
  if (x %in% c("science", "sci")) return("Science")
  if (x %in% c("social", "social science", "social studies")) return("Social Science")
  if (x %in% c("non-core", "noncore", "other")) return("Non-Core")
  NA_character_
}

parse_subject_overrides <- function(value) {
  lines <- trimws(unlist(strsplit(value, "\n", fixed = TRUE)))
  lines <- lines[nzchar(lines)]
  if (length(lines) == 0) {
    return(data.frame(.course_code = character(), .subject_override = character()))
  }

  parsed <- lapply(lines, function(line) {
    pieces <- trimws(unlist(strsplit(line, "=", fixed = TRUE)))
    if (length(pieces) < 2) return(NULL)
    subject <- normalize_subject(pieces[[2]])
    if (is.na(subject)) return(NULL)
    data.frame(
      .course_code = normalize_key(pieces[[1]]),
      .subject_override = subject,
      stringsAsFactors = FALSE
    )
  })
  bind_rows(parsed) |> distinct(.course_code, .keep_all = TRUE)
}

classify_subject <- function(description) {
  value <- tolower(ifelse(is.na(description), "", description))
  result <- rep(NA_character_, length(value))

  social <- grepl(
    "social[ -]?(science|studies)|world studies|history|civics|government|economics|geography|psychology|sociology|political|law",
    value
  )
  math <- grepl(
    "math|algebra|geometry|trigonometry|calculus|statistics|quantitative|precalculus",
    value
  )
  science <- grepl(
    "science|biology|chemistry|physics|environmental|anatomy|physiology|astronomy|earth science|forensic",
    value
  )
  english <- grepl(
    "english|language & literacy|language and literacy|literature|creative writing|composition|reading|journalism|communications",
    value
  )

  result[social] <- "Social Science"
  result[math] <- "Math"
  result[science] <- "Science"
  result[english] <- "English"
  result
}

expand_period_label <- function(value) {
  value <- clean_text(value)
  if (length(value) == 0 || is.na(value[[1]])) return(character())
  value <- value[[1]]
  suffix <- if (grepl("\\([^)]*\\)$", value)) sub("^.*?(\\([^)]*\\))$", "\\1", value) else ""
  core <- trimws(sub("\\([^)]*\\)$", "", value))
  pieces <- trimws(unlist(strsplit(core, ",", fixed = TRUE)))

  expanded <- unlist(lapply(pieces, function(piece) {
    if (grepl("^[0-9]+-[0-9]+$", piece)) {
      text_bounds <- unlist(strsplit(piece, "-", fixed = TRUE))
      bounds <- as.integer(text_bounds)
      width <- max(nchar(text_bounds))
      paste0(sprintf(paste0("%0", width, "d"), seq(bounds[[1]], bounds[[2]])), suffix)
    } else {
      paste0(piece, suffix)
    }
  }))
  unique(expanded[nzchar(expanded)])
}

period_matches_master <- function(student_period, master_period) {
  student_period <- clean_text(student_period)
  master_period <- clean_text(master_period)
  if (length(student_period) == 0 || length(master_period) == 0 ||
      is.na(student_period[[1]]) || is.na(master_period[[1]])) return(FALSE)
  any(expand_period_label(student_period[[1]]) %in% expand_period_label(master_period[[1]]))
}

expand_schedule_rows <- function(data) {
  if (nrow(data) == 0) {
    data$.period_expanded <- character()
    return(data)
  }

  bind_rows(lapply(seq_len(nrow(data)), function(index) {
    periods <- expand_period_label(data$.period[[index]])
    if (length(periods) == 0) periods <- NA_character_
    row <- data[rep(index, length(periods)), , drop = FALSE]
    row$.period_expanded <- periods
    row
  }))
}

safe_ratio <- function(numerator, denominator) {
  ifelse(denominator > 0, numerator / denominator, NA_real_)
}

safe_sheet <- function(workbook, name, data) {
  addWorksheet(workbook, name)
  if (nrow(data) == 0) {
    writeData(workbook, name, data.frame(Message = "No records found."))
  } else {
    writeDataTable(workbook, name, data, tableStyle = "TableStyleMedium2")
    freezePane(workbook, name, firstRow = TRUE)
    setColWidths(workbook, name, cols = seq_len(ncol(data)), widths = "auto", ignoreMergedCells = TRUE)
    percentage_columns <- which(grepl("Percentage|Ratio", names(data), ignore.case = TRUE))
    if (length(percentage_columns) > 0) {
      percentage_style <- createStyle(numFmt = "0.0%")
      addStyle(
        workbook, name, percentage_style,
        rows = 2:(nrow(data) + 1), cols = percentage_columns,
        gridExpand = TRUE, stack = TRUE
      )
    }
  }
}

value_card <- function(title, value_output, accent) {
  div(
    class = "audit-card",
    style = paste0("border-top: 5px solid ", accent, ";"),
    div(class = "audit-card-title", title),
    div(class = "audit-card-value", textOutput(value_output, inline = TRUE))
  )
}

mapping_section <- function(title, ...) {
  tags$details(
    class = "mapping-section",
    tags$summary(title),
    div(class = "mapping-content", ...)
  )
}

# =============================================================================
# USER INTERFACE
# =============================================================================

ui <- fluidPage(
  tags$head(
    tags$style(HTML("\n      body { background: #f5f7fb; color: #172033; }\n      .well { background: #ffffff; border: 1px solid #dfe5ef; border-radius: 10px; }\n      .audit-card { background: white; border-radius: 10px; padding: 16px; margin: 6px 0 16px;\n                    min-height: 112px; box-shadow: 0 2px 8px rgba(20, 35, 60, .08); }\n      .audit-card-title { color: #5b667a; font-size: 13px; font-weight: 600; text-transform: uppercase; }\n      .audit-card-value { color: #172033; font-size: 32px; font-weight: 700; margin-top: 8px; }\n      .help-note { color: #5b667a; font-size: 12px; margin-top: -6px; margin-bottom: 12px; }\n      .btn-primary { background-color: #2855a6; border-color: #2855a6; }\n      .mapping-section { border-top: 1px solid #e4e8f0; padding: 10px 0; }\n      .mapping-section summary { cursor: pointer; font-weight: 650; color: #2855a6; }\n      .mapping-content { padding-top: 12px; }\n      h2 { font-weight: 700; }\n    "))
  ),
  titlePanel("Student Schedule Audit Dashboard"),
  sidebarLayout(
    sidebarPanel(
      fileInput("student_file", "1. Student schedules CSV", accept = ".csv"),
      fileInput("master_file", "2. Master schedule CSV", accept = ".csv"),
      fileInput("dl_file", "3. DL core level CSV", accept = ".csv"),
      div(class = "help-note", "All three files are required. Student IDs are imported as text."),

      mapping_section(
        "Student schedule columns",
        selectInput("student_id_col", "Student ID", choices = NULL),
        selectInput("course_code_col", "Course code", choices = NULL),
        selectInput("student_period_col", "Schedule/period", choices = NULL),
        selectInput("student_teacher_col", "Teacher", choices = NULL),
        selectInput("student_description_col", "Course description", choices = NULL),
        selectInput("sped_status_col", "SPED/IEP status", choices = NULL),
        selectInput("first_name_col", "First name (optional)", choices = c("None" = "")),
        selectInput("last_name_col", "Last name (optional)", choices = c("None" = ""))
      ),

      mapping_section(
        "Master schedule columns",
        selectInput("master_class_col", "Class/course code", choices = NULL),
        selectInput("master_period_col", "Schedule/period", choices = NULL),
        selectInput("master_teacher_col", "Teacher", choices = NULL),
        selectInput("master_description_col", "Course description", choices = NULL)
      ),

      mapping_section(
        "DL core level columns",
        selectInput("dl_id_col", "Student ID", choices = NULL),
        selectInput("dl_english_col", "English placement", choices = NULL),
        selectInput("dl_math_col", "Math placement", choices = NULL),
        selectInput("dl_science_col", "Science placement", choices = NULL),
        selectInput("dl_social_col", "Social science placement", choices = NULL)
      ),

      mapping_section(
        "Audit and DL rules",
        textInput(
          "expected_periods", "Expected periods",
          value = "01(Reg),02(Reg),03(Reg),04(Reg),05(Reg),06(Reg),07(Reg),08(Reg)"
        ),
        textInput("active_iep_value", "Value indicating an active IEP", value = "Active"),
        textInput("service_codes", "DL service placement codes", value = "SC,CTT"),
        textAreaInput(
          "subject_overrides", "Optional course-subject overrides",
          placeholder = "DEMO-001=Science\nDEMO-002=Math",
          rows = 4
        ),
        div(
          class = "help-note",
          "Use Course Code=English, Math, Science, Social Science, or Non-Core. Automatic classifications appear in the Subject Classification tab."
        )
      ),

      actionButton("run_audit", "Run dashboard analysis", class = "btn-primary", width = "100%"),
      tags$br(), tags$br(),
      downloadButton("download_report", "Download Excel analysis", width = "100%")
    ),
    mainPanel(
      uiOutput("audit_message"),
      fluidRow(
        column(3, value_card("Students", "student_count", "#2855a6")),
        column(3, value_card("Teacher-Periods", "teacher_period_count", "#0f766e")),
        column(3, value_card("Active IEP Students", "active_iep_count", "#c2415d")),
        column(3, value_card("DL Service Assignments", "dl_service_count", "#7c3aed"))
      ),
      tabsetPanel(
        id = "audit_tabs",
        tabPanel("Student Summary", DTOutput("summary_table")),
        tabPanel("Teacher-Period Enrollment", DTOutput("teacher_period_table")),
        tabPanel("DL Ratio Analysis", DTOutput("dl_ratio_table")),
        tabPanel("Subject Classification", DTOutput("subject_table")),
        tabPanel("Master Schedule Validation", DTOutput("master_validation_table")),
        tabPanel("Duplicate Courses", DTOutput("duplicate_table")),
        tabPanel("Missing Periods", DTOutput("missing_table")),
        tabPanel("More Than 8", DTOutput("over_table")),
        tabPanel("No Scheduled DL Students", DTOutput("no_schedule_table")),
        tabPanel("Period Overlaps", DTOutput("overlap_table")),
        tabPanel("DL Data Quality", DTOutput("dl_quality_table")),
        tabPanel(
          "Student Schedule",
          selectizeInput("student_filter", "Student", choices = NULL, selected = "All"),
          DTOutput("student_schedule_table")
        )
      )
    )
  )
)

# =============================================================================
# ANALYSIS ENGINE
# =============================================================================

run_schedule_analysis <- function(student_raw, master_raw, dl_raw, settings) {
  expected <- parse_list(settings$expected_periods)
  service_codes <- normalize_key(parse_list(settings$service_codes))
  active_value <- tolower(trimws(settings$active_iep_value))

  master <- master_raw |>
    transmute(
      .class_code = normalize_key(.data[[settings$master_class_col]]),
      .master_period = clean_text(.data[[settings$master_period_col]]),
      .master_teacher = clean_text(.data[[settings$master_teacher_col]]),
      .master_description = clean_text(.data[[settings$master_description_col]])
    ) |>
    filter(!is.na(.class_code)) |>
    group_by(.class_code) |>
    summarise(
      .master_period = first_nonblank(.master_period),
      .master_teacher = first_nonblank(.master_teacher),
      .master_description = first_nonblank(.master_description),
      .groups = "drop"
    )

  dl_duplicate_ids <- dl_raw |>
    transmute(.student_id = normalize_key(.data[[settings$dl_id_col]])) |>
    filter(!is.na(.student_id)) |>
    count(.student_id, name = "DL File Row Count") |>
    filter(`DL File Row Count` > 1)

  dl <- dl_raw |>
    transmute(
      .student_id = normalize_key(.data[[settings$dl_id_col]]),
      .english = normalize_key(.data[[settings$dl_english_col]]),
      .math = normalize_key(.data[[settings$dl_math_col]]),
      .science = normalize_key(.data[[settings$dl_science_col]]),
      .social = normalize_key(.data[[settings$dl_social_col]])
    ) |>
    filter(!is.na(.student_id)) |>
    group_by(.student_id) |>
    summarise(
      .english = first_nonblank(.english),
      .math = first_nonblank(.math),
      .science = first_nonblank(.science),
      .social = first_nonblank(.social),
      .groups = "drop"
    )

  working <- student_raw
  working$.source_row <- seq_len(nrow(working)) + 1L
  working$.student_id <- normalize_key(working[[settings$student_id_col]])
  working$.course_code <- normalize_key(working[[settings$course_code_col]])
  working$.period <- clean_text(working[[settings$student_period_col]])
  working$.teacher <- clean_text(working[[settings$student_teacher_col]])
  working$.student_description <- clean_text(working[[settings$student_description_col]])
  working$.sped_status <- clean_text(working[[settings$sped_status_col]])

  first_name <- if (has_choice(settings$first_name_col)) clean_text(working[[settings$first_name_col]]) else NA_character_
  last_name <- if (has_choice(settings$last_name_col)) clean_text(working[[settings$last_name_col]]) else NA_character_
  working$.student_name <- trimws(paste(ifelse(is.na(first_name), "", first_name), ifelse(is.na(last_name), "", last_name)))
  working$.student_name[working$.student_name == ""] <- NA_character_

  overrides <- parse_subject_overrides(settings$subject_overrides)
  working <- working |>
    left_join(master, by = c(".course_code" = ".class_code")) |>
    left_join(dl, by = ".student_id") |>
    left_join(overrides, by = ".course_code") |>
    mutate(
      .course_description = coalesce(.student_description, .master_description),
      .automatic_subject = classify_subject(.course_description),
      .core_subject = coalesce(.subject_override, .automatic_subject),
      .subject_source = case_when(
        !is.na(.subject_override) ~ "Override",
        !is.na(.automatic_subject) ~ "Automatic",
        TRUE ~ "Unclassified"
      ),
      .analysis_teacher = coalesce(.master_teacher, .teacher, "Unassigned"),
      .active_iep = !is.na(.sped_status) & tolower(.sped_status) == active_value,
      .master_match = !is.na(.master_period),
      .master_period_match = mapply(period_matches_master, .period, .master_period),
      .master_teacher_match = case_when(
        is.na(.teacher) | is.na(.master_teacher) ~ NA,
        TRUE ~ tolower(.teacher) == tolower(.master_teacher)
      ),
      .placement = case_when(
        .core_subject == "English" ~ .english,
        .core_subject == "Math" ~ .math,
        .core_subject == "Science" ~ .science,
        .core_subject == "Social Science" ~ .social,
        TRUE ~ NA_character_
      ),
      .receives_dl_service = .active_iep & !is.na(.placement) & .placement %in% service_codes
    )

  expanded <- expand_schedule_rows(working)
  detail_rows <- expanded |>
    filter(!is.na(.student_id), !is.na(.course_code), !is.na(.period_expanded)) |>
    distinct(.student_id, .course_code, .period_expanded, .keep_all = TRUE)

  scheduled_students <- working |>
    filter(!is.na(.student_id)) |>
    group_by(.student_id) |>
    summarise(
      .student_name = first_nonblank(.student_name),
      .sped_status = first_nonblank(.sped_status),
      .groups = "drop"
    )

  student_master <- full_join(
    scheduled_students,
    dl |> transmute(.student_id, .in_dl_file = TRUE),
    by = ".student_id"
  ) |>
    mutate(.in_dl_file = coalesce(.in_dl_file, FALSE))

  summary_rows <- lapply(student_master$.student_id, function(student_id) {
    periods <- unique(detail_rows$.period_expanded[detail_rows$.student_id == student_id])
    periods <- periods[!is.na(periods)]
    ordered_periods <- c(expected[expected %in% periods], sort(periods[!periods %in% expected]))
    missing_periods <- expected[!expected %in% periods]
    unexpected_periods <- periods[!periods %in% expected]
    period_count <- length(periods)
    student_name <- first_nonblank(student_master$.student_name[student_master$.student_id == student_id])
    sped_status <- first_nonblank(student_master$.sped_status[student_master$.student_id == student_id])
    in_dl <- isTRUE(student_master$.in_dl_file[student_master$.student_id == student_id][[1]])

    status <- if (period_count == 0) {
      "No scheduled periods"
    } else if (period_count < length(expected)) {
      "Fewer than expected"
    } else if (period_count > length(expected)) {
      "More than expected"
    } else if (length(missing_periods) > 0) {
      "Missing expected period(s)"
    } else {
      "Complete"
    }

    data.frame(
      `Student ID` = student_id,
      `Student Name` = ifelse(is.na(student_name), "", student_name),
      `SPED Status` = ifelse(is.na(sped_status), "", sped_status),
      `In DL Core File` = ifelse(in_dl, "Yes", "No"),
      `Scheduled Period Count` = period_count,
      `Scheduled Periods` = paste(ordered_periods, collapse = ", "),
      `Missing Periods` = paste(missing_periods, collapse = ", "),
      `Unexpected Periods` = paste(unexpected_periods, collapse = ", "),
      Status = status,
      check.names = FALSE,
      stringsAsFactors = FALSE
    )
  })

  summary <- if (length(summary_rows) == 0) {
    data.frame(
      `Student ID` = character(), `Student Name` = character(), `SPED Status` = character(),
      `In DL Core File` = character(), `Scheduled Period Count` = integer(),
      `Scheduled Periods` = character(), `Missing Periods` = character(),
      `Unexpected Periods` = character(), Status = character(), check.names = FALSE
    )
  } else {
    bind_rows(summary_rows) |> arrange(`Student ID`)
  }

  duplicate_keys <- working |>
    filter(!is.na(.student_id), !is.na(.course_code)) |>
    count(.student_id, .course_code, name = "Duplicate Count") |>
    filter(`Duplicate Count` > 1)

  duplicates <- working |>
    inner_join(duplicate_keys, by = c(".student_id", ".course_code")) |>
    transmute(
      `Student ID` = .student_id,
      `Student Name` = coalesce(.student_name, ""),
      `Course Code` = .course_code,
      Period = .period,
      Teacher = .analysis_teacher,
      `Duplicate Count` = `Duplicate Count`,
      `Source Row` = .source_row
    ) |>
    arrange(`Student ID`, `Course Code`, Period)

  overlaps <- detail_rows |>
    group_by(.student_id, .period_expanded) |>
    summarise(
      .student_name = first_nonblank(.student_name),
      `Distinct Course Count` = n_distinct(.course_code),
      `Course Codes` = paste_nonblank(.course_code),
      Teachers = paste_nonblank(.analysis_teacher),
      `Source Rows` = paste(sort(unique(.source_row)), collapse = ", "),
      .groups = "drop"
    ) |>
    filter(`Distinct Course Count` > 1) |>
    transmute(
      `Student ID` = .student_id,
      `Student Name` = coalesce(.student_name, ""),
      Period = .period_expanded,
      `Distinct Course Count`, `Course Codes`, Teachers, `Source Rows`
    ) |>
    arrange(`Student ID`, Period)

 
  teacher_period <- detail_rows |>
    group_by(.analysis_teacher, .period_expanded) |>
    summarise(
      `Course Codes` = paste_nonblank(.course_code),
      `Course Descriptions` = paste_nonblank(.course_description),
      `Core Subjects` = paste_nonblank(.core_subject),
      `Total Enrollment` = n(),
      `Active IEP Count` = sum(.active_iep %in% TRUE),
      `Active IEP Percentage` = safe_ratio(`Active IEP Count`, `Total Enrollment`),
      `SC Count` = sum(.active_iep %in% TRUE & .placement == "SC", na.rm = TRUE),
      `No Service Count` = sum(.active_iep %in% TRUE & !(.placement %in% c("SC", "CTT"))      ),
      `CTT Count` = sum(.active_iep %in% TRUE & .placement == "CTT", na.rm = TRUE),
      `DL Service Count` = sum(.receives_dl_service %in% TRUE),
      `DL Service Percentage` = safe_ratio(`DL Service Count`, `Total Enrollment`),
      
      `Active IEP Students and Services` = {
        active_rows <- .active_iep %in% TRUE
        
        if (!any(active_rows)) {
          ""
        } else {
          student_names <- coalesce(
            .student_name[active_rows],
            ""
          )
          
          student_ids <- .student_id[active_rows]
          
          student_labels <- ifelse(
            nzchar(student_names),
            paste0(
              student_names,
              " (",
              student_ids,
              ")"
            ),
            student_ids
          )
          
          service_labels <- case_when(
            .placement[active_rows] == "SC" ~ "SC",
            .placement[active_rows] == "CTT" ~ "CTT",
            TRUE ~ "None"
          )
          
          student_entries <- paste0(
            student_labels,
            " — ",
            .course_code[active_rows],
            " — ",
            service_labels
          )
          
          paste(
            sort(unique(student_entries)),
            collapse = "; "
          )
        }
      },
      
      .groups = "drop"
    ) |>
    transmute(
      Teacher = .analysis_teacher,
      Period = .period_expanded,
      `Course Codes`,
      `Course Descriptions`,
      `Core Subjects` = ifelse(
        nzchar(`Core Subjects`),
        `Core Subjects`,
        "Unclassified"
      ),
      `Total Enrollment`,
      `Active IEP Count`,
      `Active IEP Percentage`,
      `SC Count`,
      `CTT Count`,
      `No Service Count`,
      `DL Service Count`,
      `DL Service Percentage`,
      `Active IEP Students and Services`
    ) |>
    arrange(Teacher, Period)

  dl_ratio <- detail_rows |>
    filter(.core_subject %in% c("English", "Math", "Science", "Social Science")) |>
    group_by(.analysis_teacher, .period_expanded, .core_subject) |>
    summarise(
      `Course Codes` = paste_nonblank(.course_code),
      `Course Descriptions` = paste_nonblank(.course_description),
      `Total Enrollment` = n(),
      `Active IEP Count` = sum(.active_iep %in% TRUE),
      `Active IEP Percentage` = safe_ratio(`Active IEP Count`, `Total Enrollment`),
      `SC Count` = sum(.active_iep %in% TRUE & .placement == "SC", na.rm = TRUE),
      `CTT Count` = sum(.active_iep %in% TRUE & .placement == "CTT", na.rm = TRUE),
      `No Placement Count` = sum(.active_iep %in% TRUE & is.na(.placement)),
      `DL Service Count` = sum(.receives_dl_service %in% TRUE),
      `DL Service Ratio` = safe_ratio(`DL Service Count`, `Total Enrollment`),
      .groups = "drop"
    ) |>
    transmute(
      Teacher = .analysis_teacher,
      Period = .period_expanded,
      Subject = .core_subject,
      `Course Codes`, `Course Descriptions`, `Total Enrollment`,
      `Active IEP Count`, `Active IEP Percentage`, `SC Count`, `CTT Count`,
      `No Placement Count`, `DL Service Count`, `DL Service Ratio`
    ) |>
    arrange(Teacher, Period, Subject)

  subject_classification <- working |>
    filter(!is.na(.course_code)) |>
    distinct(.course_code, .course_description, .core_subject, .subject_source) |>
    transmute(
      `Course Code` = .course_code,
      `Course Description` = coalesce(.course_description, ""),
      `Core Subject` = coalesce(.core_subject, "Unclassified"),
      `Classification Source` = .subject_source
    ) |>
    arrange(`Core Subject`, `Course Code`)

  master_validation <- working |>
    filter(!is.na(.student_id), !is.na(.course_code)) |>
    transmute(
      `Student ID` = .student_id,
      `Student Name` = coalesce(.student_name, ""),
      `Course Code` = .course_code,
      `Student Period` = .period,
      `Master Period` = coalesce(.master_period, ""),
      `Course Found in Master` = ifelse(.master_match, "Yes", "No"),
      `Period Matches Master` = ifelse(.master_match & .master_period_match, "Yes", "No"),
      `Student Teacher` = coalesce(.teacher, ""),
      `Master Teacher` = coalesce(.master_teacher, ""),
      `Teacher Matches Master` = case_when(
        is.na(.master_teacher_match) ~ "Not compared",
        .master_teacher_match ~ "Yes",
        TRUE ~ "No"
      ),
      `Source Row` = .source_row
    ) |>
    filter(
      `Course Found in Master` == "No" |
        `Period Matches Master` == "No" |
        `Teacher Matches Master` == "No"
    ) |>
    arrange(`Student ID`, `Course Code`)

  invalid_dl <- bind_rows(
    dl |> filter(!is.na(.english), !.english %in% service_codes) |>
      transmute(`Student ID` = .student_id, Subject = "English", Placement = .english,
                Issue = "Unrecognized nonblank placement; not counted"),
    dl |> filter(!is.na(.math), !.math %in% service_codes) |>
      transmute(`Student ID` = .student_id, Subject = "Math", Placement = .math,
                Issue = "Unrecognized nonblank placement; not counted"),
    dl |> filter(!is.na(.science), !.science %in% service_codes) |>
      transmute(`Student ID` = .student_id, Subject = "Science", Placement = .science,
                Issue = "Unrecognized nonblank placement; not counted"),
    dl |> filter(!is.na(.social), !.social %in% service_codes) |>
      transmute(`Student ID` = .student_id, Subject = "Social Science", Placement = .social,
                Issue = "Unrecognized nonblank placement; not counted")
  )

  active_missing_dl <- working |>
    filter(.active_iep, is.na(.english), is.na(.math), is.na(.science), is.na(.social)) |>
    distinct(.student_id) |>
    transmute(
      `Student ID` = .student_id, Subject = "", Placement = "",
      Issue = "Active IEP student has no matched DL core level record or placements"
    )

  duplicate_dl_quality <- dl_duplicate_ids |>
    transmute(
      `Student ID` = .student_id, Subject = "", Placement = "",
      Issue = paste0("Duplicate student ID in DL file (", `DL File Row Count`, " rows)")
    )

  dl_quality <- bind_rows(invalid_dl, active_missing_dl, duplicate_dl_quality) |>
    arrange(`Student ID`, Subject)

  schedule_detail <- working |>
    transmute(
      `Source Row` = .source_row,
      `Student ID` = .student_id,
      `Student Name` = coalesce(.student_name, ""),
      `SPED Status` = coalesce(.sped_status, ""),
      `Course Code` = .course_code,
      `Course Description` = coalesce(.course_description, ""),
      Teacher = .analysis_teacher,
      Period = .period,
      `Core Subject` = coalesce(.core_subject, "Unclassified"),
      `DL Placement for Subject` = coalesce(.placement, ""),
      `Receives DL Service in Subject` = ifelse(.receives_dl_service, "Yes", "No")
    )

  list(
    expected = expected,
    student_raw = student_raw,
    master_raw = master_raw,
    dl_raw = dl_raw,
    summary = summary,
    teacher_period = teacher_period,
    dl_ratio = dl_ratio,
    subject_classification = subject_classification,
    master_validation = master_validation,
    duplicates = duplicates,
    missing = summary |> filter(nzchar(`Missing Periods`)),
    over = summary |> filter(`Scheduled Period Count` > length(expected)),
    no_schedule = summary |> filter(`Scheduled Period Count` == 0, `In DL Core File` == "Yes"),
    overlaps = overlaps,
    dl_quality = dl_quality,
    schedule_detail = schedule_detail
  )
}

# =============================================================================
# SERVER
# =============================================================================

server <- function(input, output, session) {
  read_upload <- function(file_input, label) {
    req(file_input)
    tryCatch(
      read_csv_as_text(file_input$datapath),
      error = function(error) {
        showNotification(
          paste("Could not read the", label, "CSV:", error$message),
          type = "error", duration = NULL
        )
        NULL
      }
    )
  }

  student_data <- reactive(read_upload(input$student_file, "student schedule"))
  master_data <- reactive(read_upload(input$master_file, "master schedule"))
  dl_data <- reactive(read_upload(input$dl_file, "DL core level"))

  observeEvent(student_data(), {
    columns <- names(student_data())
    req(length(columns) > 0)
    updateSelectInput(session, "student_id_col", choices = columns,
                      selected = guess_column(columns, c("studentid", "studentnumber", "localid", "id")))
    updateSelectInput(session, "course_code_col", choices = columns,
                      selected = guess_column(columns, c("course", "coursecode", "class")))
    updateSelectInput(session, "student_period_col", choices = columns,
                      selected = guess_column(columns, c("schedule", "period", "block")))
    updateSelectInput(session, "student_teacher_col", choices = columns,
                      selected = guess_column(columns, c("teacher", "instructor")))
    updateSelectInput(session, "student_description_col", choices = columns,
                      selected = guess_column(columns, c("description", "coursename", "title")))
    updateSelectInput(session, "sped_status_col", choices = columns,
                      selected = guess_column(columns, c("specialedstus", "spedstatus", "iepstatus", "specialeducation")))
    updateSelectInput(session, "first_name_col", choices = c("None" = "", columns),
                      selected = if (any(grepl("first", tolower(columns)))) guess_column(columns, c("firstname", "first")) else "")
    updateSelectInput(session, "last_name_col", choices = c("None" = "", columns),
                      selected = if (any(grepl("last", tolower(columns)))) guess_column(columns, c("lastname", "last")) else "")
  })

  observeEvent(master_data(), {
    columns <- names(master_data())
    req(length(columns) > 0)
    updateSelectInput(session, "master_class_col", choices = columns,
                      selected = guess_column(columns, c("class", "coursecode", "course")))
    updateSelectInput(session, "master_period_col", choices = columns,
                      selected = guess_column(columns, c("schedule", "period", "block")))
    updateSelectInput(session, "master_teacher_col", choices = columns,
                      selected = guess_column(columns, c("teacher", "name", "instructor")))
    updateSelectInput(session, "master_description_col", choices = columns,
                      selected = guess_column(columns, c("description", "classname", "title")))
  })

  observeEvent(dl_data(), {
    columns <- names(dl_data())
    req(length(columns) > 0)
    updateSelectInput(session, "dl_id_col", choices = columns,
                      selected = guess_column(columns, c("studentid", "studentnumber", "localid", "id")))
    updateSelectInput(session, "dl_english_col", choices = columns,
                      selected = guess_column(columns, c("english", "ela")))
    updateSelectInput(session, "dl_math_col", choices = columns,
                      selected = guess_column(columns, c("math", "mathematics")))
    updateSelectInput(session, "dl_science_col", choices = columns,
                      selected = guess_column(columns, c("science", "sci")))
    updateSelectInput(session, "dl_social_col", choices = columns,
                      selected = guess_column(columns, c("socialscience", "socialstudies", "social")))
  })

  audit <- eventReactive(input$run_audit, {
    student_raw <- student_data()
    master_raw <- master_data()
    dl_raw <- dl_data()

    required_student <- c(
      input$student_id_col, input$course_code_col, input$student_period_col,
      input$student_teacher_col, input$student_description_col, input$sped_status_col
    )
    required_master <- c(
      input$master_class_col, input$master_period_col,
      input$master_teacher_col, input$master_description_col
    )
    required_dl <- c(
      input$dl_id_col, input$dl_english_col, input$dl_math_col,
      input$dl_science_col, input$dl_social_col
    )

    req(student_raw, master_raw, dl_raw, required_student, required_master, required_dl)
    validate(need(length(unique(required_student)) == length(required_student),
                  "Student schedule mappings must use different columns."))
    validate(need(length(unique(required_master)) == length(required_master),
                  "Master schedule mappings must use different columns."))
    validate(need(length(unique(required_dl)) == length(required_dl),
                  "DL core level mappings must use different columns."))
    validate(need(length(parse_list(input$expected_periods)) > 0,
                  "Enter at least one expected period."))
    validate(need(length(parse_list(input$service_codes)) > 0,
                  "Enter at least one DL service placement code."))
    validate(need(nzchar(trimws(input$active_iep_value)),
                  "Enter the student-schedule value that identifies an active IEP."))

    settings <- list(
      student_id_col = input$student_id_col,
      course_code_col = input$course_code_col,
      student_period_col = input$student_period_col,
      student_teacher_col = input$student_teacher_col,
      student_description_col = input$student_description_col,
      sped_status_col = input$sped_status_col,
      first_name_col = input$first_name_col,
      last_name_col = input$last_name_col,
      master_class_col = input$master_class_col,
      master_period_col = input$master_period_col,
      master_teacher_col = input$master_teacher_col,
      master_description_col = input$master_description_col,
      dl_id_col = input$dl_id_col,
      dl_english_col = input$dl_english_col,
      dl_math_col = input$dl_math_col,
      dl_science_col = input$dl_science_col,
      dl_social_col = input$dl_social_col,
      expected_periods = input$expected_periods,
      active_iep_value = input$active_iep_value,
      service_codes = input$service_codes,
      subject_overrides = input$subject_overrides
    )

    run_schedule_analysis(student_raw, master_raw, dl_raw, settings)
  }, ignoreInit = TRUE)

  observeEvent(audit(), {
    labels <- ifelse(
      nzchar(audit()$summary$`Student Name`),
      paste0(audit()$summary$`Student ID`, " — ", audit()$summary$`Student Name`),
      audit()$summary$`Student ID`
    )
    choices <- setNames(c("All", audit()$summary$`Student ID`), c("All", labels))
    updateSelectizeInput(session, "student_filter", choices = choices, selected = "All", server = TRUE)
  })

  output$audit_message <- renderUI({
    result <- audit()
    unclassified <- sum(result$subject_classification$`Core Subject` == "Unclassified")
    if (unclassified > 0) {
      div(
        class = "alert alert-warning",
        paste(
          "Analysis complete.", unclassified,
          "course(s) could not be assigned to a core subject automatically. Review Subject Classification and add overrides before relying on DL ratios for those courses."
        )
      )
    } else {
      div(class = "alert alert-success", "Analysis complete. All scheduled courses were classified or marked by an override.")
    }
  })

  output$student_count <- renderText({ format(nrow(audit()$summary), big.mark = ",") })
  output$teacher_period_count <- renderText({ format(nrow(audit()$teacher_period), big.mark = ",") })
  output$active_iep_count <- renderText({
    format(sum(tolower(audit()$summary$`SPED Status`) == tolower(input$active_iep_value), na.rm = TRUE), big.mark = ",")
  })
  output$dl_service_count <- renderText({
    format(sum(audit()$dl_ratio$`DL Service Count`, na.rm = TRUE), big.mark = ",")
  })

  dt <- function(data) {
    table <- datatable(
      data,
      rownames = FALSE,
      filter = "top",
      extensions = "Buttons",
      options = list(
        pageLength = 15,
        scrollX = TRUE,
        autoWidth = TRUE,
        dom = "Bfrtip",
        buttons = c("copy", "csv")
      )
    )
    percentage_columns <- names(data)[grepl("Percentage|Ratio", names(data), ignore.case = TRUE)]
    if (length(percentage_columns) > 0) {
      table <- formatPercentage(table, columns = percentage_columns, digits = 1)
    }
    table
  }

  output$summary_table <- renderDT(dt(audit()$summary), server = FALSE)
  output$teacher_period_table <- renderDT(dt(audit()$teacher_period), server = FALSE)
  output$dl_ratio_table <- renderDT(dt(audit()$dl_ratio), server = FALSE)
  output$subject_table <- renderDT(dt(audit()$subject_classification), server = FALSE)
  output$master_validation_table <- renderDT(dt(audit()$master_validation), server = FALSE)
  output$duplicate_table <- renderDT(dt(audit()$duplicates), server = FALSE)
  output$missing_table <- renderDT(dt(audit()$missing), server = FALSE)
  output$over_table <- renderDT(dt(audit()$over), server = FALSE)
  output$no_schedule_table <- renderDT(dt(audit()$no_schedule), server = FALSE)
  output$overlap_table <- renderDT(dt(audit()$overlaps), server = FALSE)
  output$dl_quality_table <- renderDT(dt(audit()$dl_quality), server = FALSE)

  output$student_schedule_table <- renderDT({
    result <- audit()
    selected <- input$student_filter
    data <- if (is.null(selected) || selected == "All") {
      result$schedule_detail
    } else {
      result$schedule_detail |> filter(`Student ID` == selected)
    }
    dt(data)
  }, server = FALSE)

  output$download_report <- downloadHandler(
    filename = function() paste0("student_schedule_dl_analysis_", Sys.Date(), ".xlsx"),
    content = function(file) {
      result <- audit()
      workbook <- createWorkbook()

      overview <- data.frame(
        Measure = c(
          "Analysis date", "Expected periods", "Students analyzed",
          "Teacher-period combinations", "Active IEP students", "DL service assignments counted",
          "Students missing expected periods", "Students with more periods",
          "DL students with no schedule", "Period overlaps",
          "Master schedule exceptions", "Unclassified courses", "DL data-quality issues"
        ),
        Value = c(
          as.character(Sys.Date()), paste(result$expected, collapse = ", "), nrow(result$summary),
          nrow(result$teacher_period), 
          sum(tolower(result$summary$`SPED Status`) == tolower(input$active_iep_value), na.rm = TRUE),
          sum(result$dl_ratio$`DL Service Count`, na.rm = TRUE),
          nrow(result$missing), nrow(result$over), nrow(result$no_schedule),
          nrow(result$overlaps), nrow(result$master_validation),
          sum(result$subject_classification$`Core Subject` == "Unclassified"), nrow(result$dl_quality)
        ),
        stringsAsFactors = FALSE
      )

      safe_sheet(workbook, "Overview", overview)
      safe_sheet(workbook, "Student Summary", result$summary)
      safe_sheet(workbook, "Teacher Period Enrollment", result$teacher_period)
      safe_sheet(workbook, "DL Subject Ratio", result$dl_ratio)
      safe_sheet(workbook, "Subject Classification", result$subject_classification)
      safe_sheet(workbook, "Master Validation", result$master_validation)
      safe_sheet(workbook, "Duplicate Courses", result$duplicates)
      safe_sheet(workbook, "Missing Periods", result$missing)
      safe_sheet(workbook, "More Than Expected", result$over)
      safe_sheet(workbook, "No Scheduled DL Students", result$no_schedule)
      safe_sheet(workbook, "Period Overlaps", result$overlaps)
      safe_sheet(workbook, "DL Data Quality", result$dl_quality)
      safe_sheet(workbook, "Schedule Detail", result$schedule_detail)
      safe_sheet(workbook, "Original Student Schedule", result$student_raw)
      safe_sheet(workbook, "Original Master Schedule", result$master_raw)
      safe_sheet(workbook, "Original DL Core Level", result$dl_raw)

      saveWorkbook(workbook, file, overwrite = TRUE)
    }
  )
}

# Launch the dashboard
shinyApp(ui = ui, server = server)
