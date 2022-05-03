## Rob Glaubius
## Avenir Health
## 2022-04-15

library(shiny)
library(shinyjs)

ui = fluidPage(
    useShinyjs(),
    tags$head(HTML("<title>Custom extract</title>")),
    titlePanel(h1("Custom extraction of CSAVR and CSAVR-adjacent data")),
    sidebarLayout(
        sidebarPanel(
            h3("1. Upload PJNZs"),
            fileInput(inputId="pjnz_list",
                      label="Choose PJNZs to upload",
                      accept=".pjnz",
                      multiple=TRUE,
                      buttonLabel="Browse...",
                      placeholder="No file selected"),
            h3("2. Download Excel summary"),
            disabled(downloadButton(outputId="download_xlsx", label="Download"))
        ),
        mainPanel(
            tableOutput("pjnz_name")
        )
    )
)

server = function(input, output, session) {
    options(shiny.maxRequestSize=75*2^20) # 75MB file size limit
    
    pjnz_meta = reactive({
        req(input$pjnz_list)
        return(input$pjnz_list)
    })
    
    output$pjnz_name = renderTable({
        rval = data.frame(PJNZ=pjnz_meta()$name)
        colnames(rval) = c("PJNZs uploaded")
        return(rval)
    })
    
    summary_xlsx = reactive({
      withProgress(message = "Extracting PJNZ data", value=0, {
        pjnz_list = pjnz_meta()$datapath
        pjnz_data = list()
        for (k in 1:length(pjnz_list)) {
          pjnz_data[[k]] = extract_pjnz_data(pjnz_list[[k]])
          incProgress(1 / length(pjnz_list), detail=sprintf("File %d of %d", k, length(pjnz_list)))
        }
      })
      withProgress(message = "Building summary workbook", value=0, {
        pjnz_summ = process_pjnzs(pjnz_data)
      })
      return(pjnz_summ)
    })
    
    observeEvent(input$pjnz_list, {
      req(summary_xlsx())
      enable("download_xlsx")  
    })
    
    output$download_xlsx = downloadHandler(
        filename = "csavr-custom-extract.xlsx",
        content = function(xlsx_name) {saveWorkbook(summary_xlsx(), file=xlsx_name, overwrite=TRUE)}
    )
}

shinyApp(ui=ui, server=server)

