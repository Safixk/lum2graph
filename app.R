library(shiny)
library(readxl)
library(dplyr)
library(tidyr)
library(stringr)
library(ggplot2)
library(scales)
library(DT)
library(glue)

locate_plate_matrix <- function(raw_df, plate_rows = 8, plate_cols = 12) {
  first_col <- str_trim(as.character(raw_df[[1]]))
  row_letters <- str_sub(first_col, 1, 1)
  candidates <- which(row_letters %in% LETTERS)

  for (start_row in candidates) {
    end_row <- start_row + plate_rows - 1
    if (end_row > nrow(raw_df)) {
      next
    }

    expected_rows <- LETTERS[seq_len(plate_rows)]
    candidate_rows <- row_letters[start_row:end_row]

    if (!identical(candidate_rows, expected_rows)) {
      next
    }

    header_row <- start_row - 1
    col_slice <- seq.int(2, min(ncol(raw_df), plate_cols + 1))
    if (length(col_slice) < plate_cols) {
      next
    }
    header_guess <- if (header_row >= 1) raw_df[header_row, col_slice, drop = TRUE] else rep(NA_character_, plate_cols)
    header_values <- as.character(unlist(header_guess))

    col_labels <- if (all(!is.na(header_values) & header_values != "")) {
      header_values
    } else {
      as.character(seq_len(plate_cols))
    }

    plate_slice <- raw_df[start_row:end_row, col_slice]
    return(list(rows = expected_rows, cols = col_labels, data = plate_slice))
  }

  NULL
}

clean_plate_data <- function(block) {
  if (is.null(block)) {
    return(NULL)
  }

  colnames(block$data) <- block$cols
  plate_tbl <- as_tibble(block$data)
  plate_tbl$Row <- factor(block$rows, levels = rev(block$rows))
  plate_tbl <- relocate(plate_tbl, Row)

  plate_tbl |> pivot_longer(
    cols = -Row,
    names_to = "Column",
    values_to = "Reading"
  ) |> mutate(
    Column = factor(Column, levels = unique(block$cols)),
    Well = paste0(as.character(Row), Column),
    Reading = suppressWarnings(as.numeric(Reading)),
    Label = Well
  )
}

ui <- fluidPage(
  titlePanel("BioTek Synergy H1 Plate Explorer"),
  sidebarLayout(
    sidebarPanel(
      fileInput("plate_file", "Upload BioTek .xlsx file", accept = c(".xlsx")),
      uiOutput("sheet_picker"),
      numericInput("rows", "Plate rows (e.g., 8 for 96-well, 16 for 384-well)", value = 8, min = 1, max = 32),
      numericInput("cols", "Plate columns (e.g., 12 for 96-well, 24 for 384-well)", value = 12, min = 1, max = 48),
      checkboxInput("show_labels", "Show custom labels on plot", value = TRUE),
      selectInput(
        "palette",
        "Color palette",
        choices = c("viridis", "magma", "plasma", "inferno", "cividis"),
        selected = "viridis"
      ),
      textInput("plot_title", "Plot title", value = "Plate heatmap"),
      downloadButton("download_plot", "Download plot (PNG)"),
      width = 3
    ),
    mainPanel(
      tabsetPanel(
        tabPanel(
          "Data",
          h4("Detected plate data"),
          verbatimTextOutput("detect_message"),
          DTOutput("plate_table")
        ),
        tabPanel(
          "Plot",
          plotOutput("plate_plot", height = 600),
          verbatimTextOutput("summary_stats")
        )
      )
    )
  )
)

server <- function(input, output, session) {
  available_sheets <- reactive({
    req(input$plate_file)
    excel_sheets(input$plate_file$datapath)
  })

  output$sheet_picker <- renderUI({
    req(input$plate_file)
    selectInput("sheet", "Sheet", choices = available_sheets())
  })

  raw_matrix <- reactive({
    req(input$plate_file, input$sheet)
    read_xlsx(input$plate_file$datapath, sheet = input$sheet, col_names = FALSE, .name_repair = "minimal")
  })

  plate_block <- reactive({
    locate_plate_matrix(raw_matrix(), plate_rows = input$rows, plate_cols = input$cols)
  })

  plate_data <- reactive({
    clean_plate_data(plate_block())
  })

  plate_store <- reactiveVal(NULL)

  observeEvent(plate_data(), {
    plate_store(plate_data())
  })

  output$detect_message <- renderText({
    req(input$plate_file, input$sheet)
    if (is.null(plate_block())) {
      "No plate-shaped data block detected. Adjust the row/column settings or check the file structure."
    } else {
      glue::glue(
        "Detected a {input$rows} x {input$cols} plate starting at the first row labeled '{plate_block()$rows[1]}'.\n",
        "Column labels: {paste(plate_block()$cols, collapse = ', ')}"
      )
    }
  })

  output$plate_table <- renderDT({
    req(plate_store())
    datatable(
      plate_store() |> arrange(desc(Row), Column),
      editable = list(target = "cell", disable = list(columns = c(0, 1, 2, 3))),
      options = list(pageLength = 12, dom = "tip")
    )
  })

  proxy <- dataTableProxy("plate_table")

  observeEvent(input$plate_table_cell_edit, {
    info <- input$plate_table_cell_edit
    if (info$col == 4) {
      new_data <- plate_store()
      new_data[info$row, info$col + 1] <- info$value
      plate_store(new_data)
      proxy |> replaceData(new_data, resetPaging = FALSE, rownames = FALSE)
    }
  })

  plot_data <- reactive({
    req(plate_store())
    plate_store()
  })

  output$plate_plot <- renderPlot({
    req(plot_data())

    fill_scale <- switch(
      input$palette,
      viridis = scale_fill_viridis_c(option = "viridis", na.value = "grey90"),
      magma = scale_fill_viridis_c(option = "magma", na.value = "grey90"),
      plasma = scale_fill_viridis_c(option = "plasma", na.value = "grey90"),
      inferno = scale_fill_viridis_c(option = "inferno", na.value = "grey90"),
      cividis = scale_fill_viridis_c(option = "cividis", na.value = "grey90"),
      scale_fill_viridis_c(na.value = "grey90")
    )

    ggplot(plot_data(), aes(x = Column, y = Row, fill = Reading)) +
      geom_tile(color = "white", size = 0.4) +
      fill_scale +
      labs(title = input$plot_title, x = "Column", y = "Row", fill = "Reading") +
      coord_equal() +
      scale_y_discrete(drop = FALSE) +
      scale_x_discrete(drop = FALSE) +
      theme_minimal(base_size = 14) +
      theme(panel.grid = element_blank()) +
      {if (isTRUE(input$show_labels)) geom_text(aes(label = Label), size = 3, color = "black")}
  })

  output$summary_stats <- renderText({
    req(plot_data())
    readings <- plot_data()$Reading

    if (all(is.na(readings))) {
      return("All readings are NA; adjust detection settings or verify the file contents.")
    }

    stats <- plot_data() |> summarise(
      wells = n(),
      min_value = min(Reading, na.rm = TRUE),
      max_value = max(Reading, na.rm = TRUE),
      mean_value = mean(Reading, na.rm = TRUE),
      median_value = median(Reading, na.rm = TRUE)
    )

    glue::glue(
      "Wells: {stats$wells}\n",
      "Min: {round(stats$min_value, 3)}\n",
      "Max: {round(stats$max_value, 3)}\n",
      "Mean: {round(stats$mean_value, 3)}\n",
      "Median: {round(stats$median_value, 3)}"
    )
  })

  output$download_plot <- downloadHandler(
    filename = function() {
      paste0("plate-plot-", Sys.Date(), ".png")
    },
    content = function(file) {
      req(plot_data())
      g <- ggplot(plot_data(), aes(x = Column, y = Row, fill = Reading)) +
        geom_tile(color = "white", size = 0.4) +
        switch(
          input$palette,
          viridis = scale_fill_viridis_c(option = "viridis", na.value = "grey90"),
          magma = scale_fill_viridis_c(option = "magma", na.value = "grey90"),
          plasma = scale_fill_viridis_c(option = "plasma", na.value = "grey90"),
          inferno = scale_fill_viridis_c(option = "inferno", na.value = "grey90"),
          cividis = scale_fill_viridis_c(option = "cividis", na.value = "grey90"),
          scale_fill_viridis_c(na.value = "grey90")
        ) +
        labs(title = input$plot_title, x = "Column", y = "Row", fill = "Reading") +
        coord_equal() +
        scale_y_discrete(drop = FALSE) +
        scale_x_discrete(drop = FALSE) +
        theme_minimal(base_size = 14) +
        theme(panel.grid = element_blank()) +
        {if (isTRUE(input$show_labels)) geom_text(aes(label = Label), size = 3, color = "black")}

      ggsave(file, plot = g, width = 8, height = 6, dpi = 300)
    }
  )
}

shinyApp(ui, server)
