library(shiny)
library(shinydashboard)
library(shinyjs)
library(reticulate)
library(shinyWidgets)

# 指定 Python 路径
#use_python("/opt/anaconda3/bin/python3")  # 本地Python路径
Sys.setenv(RETICULATE_CONDA = "/root/miniconda3/bin/conda") # 服务器conda环境

# ==== UI ====
ui <- dashboardPage(
  dashboardHeader(title = "📄 PDF 转 Word 平台"),
  dashboardSidebar(
    sidebarMenu(
      menuItem("文件转换", tabName = "convert", icon = icon("exchange-alt")),
      br(),
      fileInput("pdf_file", "上传 PDF 文件", accept = ".pdf"),
      radioButtons("method", "选择转换方式：",
                   choices = list("使用 pdf2docx（格式保留好）" = "pdf2docx",
                                  "使用 PyMuPDF（兼容性强）" = "pymupdf")),
      actionButton("convert_btn", "开始转换", icon = icon("play")),
      br(), br(),
      uiOutput("download_ui")
    )
  ),
  dashboardBody(
    useShinyjs(),
    tabItems(
      tabItem(tabName = "convert",
              fluidRow(
                box(
                  width = 12, status = "primary", solidHeader = TRUE,
                  title = "转换进度",
                  progressBar(
                    id = "progress", value = 0, total = 100,
                    title = "等待开始...", display_pct = TRUE
                  )
                )
              ),
              fluidRow(
                box(
                  width = 12, status = "info", solidHeader = TRUE,
                  title = "提示信息",
                  verbatimTextOutput("status")
                )
              )
      )
    )
  )
)

# ==== Server ====
server <- function(input, output, session) {
  
  # 创建响应式值来存储转换状态
  rv <- reactiveValues(
    conversion_done = FALSE,
    word_path = NULL
  )
  
  observe({
    if (is.null(input$pdf_file)) {
      output$status <- renderText("请上传一个 PDF 文件。")
    } else {
      output$status <- renderText(paste("已上传文件：", input$pdf_file$name))
    }
  })
  
  output$download_ui <- renderUI(NULL)
  
  # 在Python环境中注册进度更新函数
  observe({
    # 定义进度更新函数
    update_progress_python <- function(progress) {
      # 在shiny的上下文中更新进度条
      shinyjs::runjs(sprintf(
        "Shiny.setInputValue('python_progress', %f);",
        progress
      ))
    }
    
    # 将函数注册到Python环境
    py$update_progress_r <- update_progress_python
  })
  
  # 监听Python进度更新
  observeEvent(input$python_progress, {
    updateProgressBar(
      session = session, 
      id = "progress", 
      value = input$python_progress, 
      title = paste0("处理进度: ", round(input$python_progress, 2), "%")
    )
  })
  
  observeEvent(input$convert_btn, {
    req(input$pdf_file)
    
    # 重置状态
    rv$conversion_done <- FALSE
    rv$word_path <- NULL
    output$download_ui <- renderUI(NULL)
    
    pdf_path <- input$pdf_file$datapath
    word_path <- tempfile(fileext = ".docx")
    
    # 清空进度条
    updateProgressBar(session = session, id = "progress", value = 0, title = "开始转换中...")
    
    tryCatch({
      if (input$method == "pdf2docx") {
        # 检查并安装 pdf2docx
        if (!py_module_available("pdf2docx")) {
          py_install("pdf2docx", pip = TRUE)
        }
        
        # Python 脚本（pdf2docx 方式）
        py_run_string("
import sys
from pdf2docx import Converter

def convert_pdf_to_word_pdf2docx(pdf_file, output_file):
    try:
        # 初始化转换器
        cv = Converter(pdf_file)
        
        # 转换所有页面
        cv.convert(output_file, start=0, end=None)
        
        # 关闭转换器
        cv.close()
        return True
    except Exception as e:
        print(f'转换错误: {str(e)}')
        return False
        ")
        
        # 执行转换
        success <- py$convert_pdf_to_word_pdf2docx(pdf_path, word_path)
        
        if (success) {
          rv$word_path <- word_path
          rv$conversion_done <- TRUE
          updateProgressBar(session, id = 'progress', value = 100, title = '转换完成 ✅')
          output$status <- renderText('转换完成，可以下载 Word 文件。')
        } else {
          stop("pdf2docx 转换失败")
        }
        
      } else if (input$method == "pymupdf") {
        # 检查并安装所需的Python包
        if (!py_module_available("fitz")) {
          py_install("PyMuPDF", pip = TRUE)
        }
        if (!py_module_available("docx")) {
          py_install("python-docx", pip = TRUE)
        }
        
        # Python 脚本（PyMuPDF 方式）- 带进度更新
        py_run_string("
import fitz
from docx import Document
import math

def convert_pdf_to_word_pymupdf(pdf_file, output_file, progress_callback):
    try:
        # 打开PDF文件
        pdf_document = fitz.open(pdf_file)
        total_pages = len(pdf_document)
        
        # 创建Word文档
        doc = Document()
        
        # 添加标题
        doc.add_heading('PDF转换结果', 0)
        
        # 逐页处理
        for page_num in range(total_pages):
            page = pdf_document.load_page(page_num)
            
            # 提取文本
            text = page.get_text()
            
            if text.strip():
                # 添加页面标题
                doc.add_heading(f'第 {page_num + 1} 页', level=1)
                
                # 添加文本内容
                paragraphs = text.split('\\n')
                for paragraph in paragraphs:
                    if paragraph.strip():
                        doc.add_paragraph(paragraph)
                
                # 添加分页符（除了最后一页）
                if page_num < total_pages - 1:
                    doc.add_page_break()
            
            # 更新进度
            progress = ((page_num + 1) / total_pages) * 100
            if progress_callback:
                progress_callback(progress)
        
        # 保存文档
        doc.save(output_file)
        pdf_document.close()
        return True
        
    except Exception as e:
        print(f'转换错误: {str(e)}')
        return False
        ")
        
        # 执行转换，传递进度回调函数
        success <- py$convert_pdf_to_word_pymupdf(pdf_path, word_path, py$update_progress_r)
        
        if (success) {
          rv$word_path <- word_path
          rv$conversion_done <- TRUE
          updateProgressBar(session, id = 'progress', value = 100, title = '转换完成 ✅')
          output$status <- renderText('转换完成，可以下载 Word 文件。')
        } else {
          stop("PyMuPDF 转换失败")
        }
      }
      
    }, error = function(e) {
      updateProgressBar(session, id = "progress", value = 0, title = "转换失败 ❌")
      output$status <- renderText(paste("转换出错：", e$message))
    })
    
    # 如果转换成功，显示下载按钮
    if (rv$conversion_done) {
      output$download_ui <- renderUI({
        downloadButton("download_word", "下载 Word 文件", class = "btn-success")
      })
    }
  })
  
  # 下载处理
  output$download_word <- downloadHandler(
    filename = function() {
      paste0(tools::file_path_sans_ext(input$pdf_file$name), ".docx")
    },
    content = function(file) {
      req(rv$word_path)
      file.copy(rv$word_path, file)
    }
  )
}

# ==== 启动应用 ====
shinyApp(ui = ui, server = server)
