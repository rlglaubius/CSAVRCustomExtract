## Rob Glaubius
## Avenir Health
## 2022-04-15

library(shiny)

ui = fluidPage(
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
            downloadButton(outputId="download_xlsx", label="Download")
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
        pjnz_list = pjnz_meta()$datapath
        return(process_pjnzs(pjnz_list))
    })
    
    output$download_xlsx = downloadHandler(
        filename = "csavr-custom-extract.xlsx",
        content = function(xlsx_name) {saveWorkbook(summary_xlsx(), file=xlsx_name, overwrite=TRUE)}
    )
}

shinyApp(ui=ui, server=server)

