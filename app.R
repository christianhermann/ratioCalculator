if (!require(shiny)) install.packages('shiny')
library(shiny)
if (!require(openxlsx)) install.packages('openxlsx')
library(openxlsx)
if (!require(tidyverse)) install.packages('tidyverse')
library(tidyverse)
if (!require(svDialogs)) install.packages('svDialogs')
library(svDialogs)
if (!require(ggthemes)) install.packages('ggthemes')
library(ggthemes)
if (!require(ggpubr)) install.packages('ggpubr')
library(ggpubr)

get_plot_data <- function(summary_list, series) {
    names(summary_list) <- series
    data <- do.call("rbind", map(summary_list, function(x) data.frame(x[["calc_Ratio"]])))
    data <- rownames_to_column(data, "Series")
    colnames(data)[2] <- "Ratio"
    data$Series <-  word(data$Series,1,-3,".")
    data$Series <-
        factor(data$Series, levels = unique(data$Series))
    return(data)
}



#' Theme for Boxplots
#'
#' @param base_size
#' @param base_family
#'
#' @return
#' @export
#'
#' @examples
theme_chris_boxplot <-
    function(base_size = 11, axis_size = 10,base_family = "Arial")
    {
        theme_foundation(base_size = base_size, base_family = base_family) %+replace%
            
            theme(
                axis.text.x = element_text(color = "black", size = axis_size, angle = 45),
                axis.text.y = element_text(color = "black", hjust = 1, size = axis_size,margin = margin(r = 5)),
                axis.text = element_text(color = "black", size = axis_size),
                axis.ticks = element_blank(),
                axis.title.x = element_blank(),
                panel.background = element_blank(),
                plot.background = element_blank(),
                plot.title = element_blank(),
                legend.position = "none",
                panel.border = element_blank(),
                axis.line = element_line(colour = "black", linetype = "solid"),
                axis.line.x = element_blank(),
                axis.ticks.y = element_line(),
                plot.margin = unit(c(0.1,0.2,0.2,1),"cm"),
                # panel.grid.major = element_line(colour = "grey",size = 0.25),
                # panel.grid.minor = element_line(colour = "grey", size = 0.25),
                # panel.grid.minor.y = element_blank(),
                panel.grid.minor = element_blank(),
                panel.grid.major = element_blank(),
                # aspect.ratio = 0.618,
                legend.title = element_blank())
    }



summary_list <- list(data.frame())
sheet_names <- c()
active_plot <- c()

# Define UI for application that draws a histogram
ui <- fluidPage(

    # App title ----
    titlePanel("Ratio Calculator"),

    sidebarLayout(sidebarPanel(
        tabsetPanel(
            tabPanel("Calc",
        withMathJax(),
        helpText("Ratio="),
        helpText("$$\\frac{Component 1 - Before}{Component 2 - Before}$$"),
        selectizeInput(
            "Component_1",
            "Component 1",
            choices = list(),
            multiple = F
        ),
        selectizeInput(
            "Component_2",
            "Component 2",
            choices = list(),
            multiple = F
        ),
        selectizeInput(
            "Offset",
            "Before",
            choices = list(),
            multiple = F
        ),
        actionButton("addRatio", "Add Ratio Column")
        ),
        tabPanel("Plot",
                     selectizeInput(
                         "Series",
                         "Series",
                         choices = list(),
                         multiple = TRUE
                     ),
                 fluidRow(
                     checkboxGroupInput(
                         "Boxplot_Extras",
                         "Extras",
                         choiceNames = list(
                             "Jitter",
                             "Observations",
                             "P-Values",
                             "Significances",
                             "Custom Y-Axis"
                         ),
                         choiceValues = list(
                             "Jitter",
                             "Observations",
                             "P-Values",
                             "Significances",
                             "Custom_Y-Axis"
                         ),
                         selected = list(
                             "Jitter",
                             "Observations",
                             "P-Values",
                             "Significances"
                         )
                     )),
                 selectizeInput(
                     "ref",
                     "Reference",
                     choices = list(),
                     multiple = F
                 ),
                 sliderInput(
                     "X_Axis_Angle_CD",
                     "Angle X-Axis Label",
                     0,
                     360,
                     value = 70,
                     step = 1
                 ),
                 fluidRow(
                 column(
                     width = 6,
                     numericInput(
                         "yMax",
                         "y-Max",
                         value = 10,
                         min = -100000,
                         max = 100000,
                         step = 1
                     )
                 ),
                 column(
                     width = 6,
                     numericInput(
                         "yMin",
                         "y-Min",
                         value = -10,
                         min = -100000,
                         max = 100000,
                         step = 1
                     )
                 ))
                 )))



                        ,mainPanel(actionButton("Import_Data", "Import Data"),

                                   actionButton("Save_xls", "Save .xls Table"),
                                   hr(),
                                   
                                   tabsetPanel(
                                       tabPanel("Summaries",
                                   selectizeInput(
                                       "choose_series",
                                       label = "Choose Series",
                                       choices = list("No data loaded" = 1, "No data loaded" = 2),
                                       selected = 1),
                    DT::dataTableOutput('Summary_data')
                ),
                tabPanel("Plot",
                         column(
                             width = 2,
                             numericInput(
                                 "Save_width",
                                 "Width",
                                 value = 100,
                                 min = 0,
                                 max = 10000,
                                 step = 0.1
                             )
                         ),
                         column(
                             width = 2,
                             numericInput(
                                 "Save_height",
                                 "Heigth",
                                 value = 100,
                                 min = 0,
                                 max = 10000,
                                 step = 0.1
                             )
                         ),
                         fluidRow(
                         mainPanel(imageOutput("Boxplot_Ratio"))
                         )))))
)


# Define server logic required to draw a histogram
server <- function(input, output,session) {


    output$Summary_data <-
        DT::renderDataTable(summary_list[[input$choose_series]],
                            options = list(
                                escape = F,
                                scrollX = T,
                                scrollY = 500
                            ))



    observeEvent(input$Import_Data, {

   #     dir_summary <- dlg_open(title = "Select summary excel file", filters = matrix(c("Excel files", "*.xls;*.xlsx"),1,2))$res
        dir_summary <- dlg_open(title = "Select summary excel file")$res
        
        try({
        sheet_names <- getSheetNames(dir_summary)

        excel_data <-
            map(sheet_names, function(sheet)
                excel_data <- read.xlsx(dir_summary, sheet = sheet))

        excel_data <-
            map(excel_data, function(x) {
                colnames(x) <- make.unique(colnames(x))
                x
            })

        excel_data <- map(excel_data, tibble::as.tibble)

        sheet_names <<-
            map_chr(sheet_names, function(sheet) {
                if (str_ends(sheet, " "))
                    return(stri_replace_last_fixed(sheet, " ", ""))
                return(sheet)
            })
        names(excel_data) <- sheet_names
        summary_list <<- excel_data
        })
        updateSelectizeInput(session, "choose_series", choices = sheet_names, selected = sheet_names[1])
        updateSelectizeInput(session,
                             "Component_1",
                             choices =  colnames(summary_list[[sheet_names[1]]])[!str_detect(colnames(summary_list[[sheet_names[1]]]), "Outlie.")])
        updateSelectizeInput(session,
                             "Component_2",
                             choices =  colnames(summary_list[[sheet_names[1]]])[!str_detect(colnames(summary_list[[sheet_names[1]]]), "Outlie.")])
        updateSelectizeInput(session,
                             "Offset",
                             choices =  colnames(summary_list[[sheet_names[1]]])[!str_detect(colnames(summary_list[[sheet_names[1]]]), "Outlie.")])
        updateSelectizeInput(session,
                             "Series",
                             choices =  sheet_names)
        updateSelectizeInput(session,
                             "ref",
                             choices =  sheet_names)

    }
)
    observeEvent(input$addRatio,{
        calculated_Ratio <- map(summary_list, function(x) {calc_Ratio <- round((x[input$Component_1] - x[input$Offset])/(x[input$Component_2] - x[input$Offset]),2); names(calc_Ratio) <- "calc_Ratio";return(calc_Ratio)})
        summary_list <<- map2(summary_list, calculated_Ratio, cbind)
        summary_list <<- map(summary_list, function(summary) {summary$Median <- median(summary$calc_Ratio); return(summary)})
        output$Summary_data <-
            DT::renderDataTable(summary_list[[input$choose_series]],
                                options = list(
                                    escape = F,
                                    scrollX = T,
                                    scrollY = 500
                                ))
        })

    observeEvent(input$Save_xls, {


        directory <- dlgSave(title = "Save Summary as Excel file")$res
        try({
        wb = createWorkbook()
        map2(summary_list, sheet_names, function(summary, sheet, directory) {
            sheet <- addWorksheet(wb, sheet)
            writeData(wb, sheet, summary)
            saveWorkbook(wb, file.path(paste0(directory, ".xlsx")), overwrite = TRUE)
        }, directory)
        })

    })
    
    observe({
        if (length(input$Series) == 0) return()
        ph <- input$Series
        compareVector <-
            c(
                "Jitter",
                "Observations",
                "P-Values",
                "Significances",
                "Custom_Y-Axis"
            )
        extrasVec <- compareVector %in% input$Boxplot_Extras
        
        if("calc_Ratio" %in% colnames(first(summary_list))) {

        data <- get_plot_data(flatten(map(input$Series, function(x, y)
            y[x], summary_list)), input$Series)
        Ratio_plot <- ggplot(data, aes( x = Series, y = Ratio, col = Series)) + geom_boxplot() + theme_chris_boxplot()  + theme(axis.text.x = element_text(angle = input$X_Axis_Angle_CD))
        if(extrasVec[1] == T) Ratio_plot <- Ratio_plot + geom_jitter()
        if(extrasVec[5] == T) Ratio_plot <- Ratio_plot + ylim(c(input$yMin, input$yMax))
        if(!is.null(ggplot_build(Ratio_plot)$layout$panel_scales_y[[1]]$limits)) {y <- max(ggplot_build(Ratio_plot)$layout$panel_scales_y[[1]]$limits*0.7)}
        else{y <- max(ggplot_build(Ratio_plot)$layout$panel_scales_y[[1]]$range$range)}
        if(extrasVec[2] == T) {
            n_fun <- function(x) {
                return(data.frame(y = y * 1.4,
                                  label = paste0("n = ",length(x))))
            }
            
            Ratio_plot <- Ratio_plot + stat_summary(
                colour = "black",
                fun.data = n_fun,
                geom = "text",
            )
        }
        ref = input$ref;
        if(extrasVec[3] == T) Ratio_plot <- Ratio_plot + stat_compare_means(method = "wilcox.test", paired = F, ref.group = ref, label = "p.format", label.y = y * 1.25)
        if(extrasVec[4] == T) Ratio_plot <- Ratio_plot + stat_compare_means(method = "wilcox.test", paired = F, ref.group = ref, label = "p.signif", label.y = y * 1.1)

        active_plot <- Ratio_plot
        active_plot <<- active_plot
        output$Boxplot_Ratio <- renderImage({
            # A temp file to save the output. It will be deleted after renderImage
            # sends it, because deleteFile=TRUE.
            outfile <- tempfile(fileext = '.png')

            # Generate a png
            ggsave(
                outfile,
                active_plot ,
                device = "png",
                width = input$Save_width,
                height = input$Save_height,
                dpi = 300,
                units = "mm",
                limitsize = F
            )

            # Return a list
            list(src = outfile,
                 alt = "This is alternate text")
        }, deleteFile = TRUE)}
    })
}

# Run the application
shinyApp(ui = ui, server = server)
