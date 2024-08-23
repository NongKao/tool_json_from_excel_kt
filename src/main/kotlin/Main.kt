package org.example

import jnafilechooser.api.JnaFileChooser
import kotlinx.serialization.Contextual
import kotlinx.serialization.Serializable
import org.apache.poi.ss.usermodel.WorkbookFactory
import java.awt.*
import java.io.File
import java.io.FileInputStream
import javax.swing.*

class App {

    private var lastExportedFilePath: String = ""
    private lateinit var messageLabel: JLabel
    private lateinit var openFileButton: JButton
    private lateinit var jsonTextArea: JTextArea

    private fun createAndShowGUI() {
        val screenSize: Dimension = Toolkit.getDefaultToolkit().screenSize
        val windowWidth = screenSize.width / 2
        val windowHeight = screenSize.height / 2

        val frame = JFrame("Convert Excel to JSON")
        frame.defaultCloseOperation = JFrame.EXIT_ON_CLOSE
        frame.setSize(windowWidth, windowHeight)
        frame.setLocationRelativeTo(null)

        val mainPanel = JPanel(GridBagLayout())
        val gbc = GridBagConstraints()
        gbc.insets = Insets(5, 5, 5, 5) // Khoảng cách giữa các thành phần
        gbc.gridx = 0

        val asciiArt = """
        <b> <font color='red'>
                             _                          _______      __
                             | |                        / ____\ \    / /
                             | |     ___  _ __   __ _  | |     \ \  / / 
        Create by            | |    / _ \| '_ \ / _` | | |      \ \/ /                       
                             | |___| (_) | | | | (_| | | |____   \  /   
                             |______\___/|_| |_|\__, |  \_____|   \/    
                                                 __/ |                  
                                                |___/                   
                </font></b>
        """.trimIndent()

        val asciiLabel = JLabel("<html><pre>$asciiArt</pre></html>")
        asciiLabel.horizontalAlignment = SwingConstants.CENTER
        asciiLabel.font = Font(Font.MONOSPACED, Font.PLAIN, 12)

        gbc.gridy = 0
        mainPanel.add(asciiLabel, gbc)

        // Tạo nút "Chọn File"
        val chooseFileButton = JButton("Select File")
        chooseFileButton.alignmentX = Component.CENTER_ALIGNMENT
        chooseFileButton.addActionListener {
            messageLabel.isVisible = false
            openFileButton.isVisible = false
            jsonTextArea.text = "" // Xóa nội dung trước đó
            val filePath = chooseFile()
            doConvertToJson(filePath)
        }

        gbc.gridy = 1
        gbc.fill = GridBagConstraints.NONE
        gbc.anchor = GridBagConstraints.CENTER
        mainPanel.add(chooseFileButton, gbc)

        messageLabel = JLabel("")
        messageLabel.font = Font(Font.SANS_SERIF, Font.BOLD, 14)
        messageLabel.horizontalAlignment = SwingConstants.CENTER

        gbc.gridy = 2
        mainPanel.add(messageLabel, gbc)

        openFileButton = JButton("Open File")
        openFileButton.isVisible = false
        openFileButton.alignmentX = Component.CENTER_ALIGNMENT
        openFileButton.addActionListener {
            val jsonFile = File(lastExportedFilePath)
            if (jsonFile.exists()) {
                openFileInExplorer(jsonFile)
            } else {
                println("File không tồn tại.")
            }
        }

        gbc.gridy = 3
        gbc.weighty = 0.0
        gbc.anchor = GridBagConstraints.CENTER
        mainPanel.add(openFileButton, gbc)

        jsonTextArea = JTextArea(10, 40)
        jsonTextArea.lineWrap = true
        jsonTextArea.wrapStyleWord = true
        jsonTextArea.isEditable = false

        val scrollPane = JScrollPane(jsonTextArea)
        scrollPane.verticalScrollBarPolicy = JScrollPane.VERTICAL_SCROLLBAR_AS_NEEDED
        scrollPane.horizontalScrollBarPolicy = JScrollPane.HORIZONTAL_SCROLLBAR_AS_NEEDED

        gbc.gridy = 4
        gbc.fill = GridBagConstraints.BOTH
        gbc.weighty = 1.0
        mainPanel.add(scrollPane, gbc)

        frame.add(mainPanel)
        frame.isVisible = true
    }

    private fun doConvertToJson(filePath: String?) {
        if (filePath != null) {
            val outputFilePath = filePath.substring(0, filePath.lastIndexOf('.')) + ".json"
            try {
                val excelData = readExcelFile(filePath)
                val jsonContent = writeJsonToFile(excelData, outputFilePath)
                jsonTextArea.text = jsonContent
                println("Đã chuyển đổi $filePath thành $outputFilePath")
                messageLabel.text = "Convert to Json success"
                messageLabel.foreground = Color.GREEN
                messageLabel.isVisible = true
                openFileButton.isVisible = true
            } catch (e: Exception) {
                e.printStackTrace()
                messageLabel.text = "Please try again!"
                messageLabel.foreground = Color.RED
                messageLabel.isVisible = true
                openFileButton.isVisible = false
            }
        } else {
            messageLabel.text = "Please try again!"
            messageLabel.foreground = Color.RED
            messageLabel.isVisible = true
            openFileButton.isVisible = false
            println("Không có file nào được chọn.")
        }
    }


    fun run() {
        createAndShowGUI()
    }

    private fun chooseFile(): String? {
        val fileChooser = JnaFileChooser()
        fileChooser.setTitle("Chọn file Excel")
        fileChooser.addFilter("Excel files", "xls", "xlsx", "csv")

        return if (fileChooser.showOpenDialog(null)) {
            fileChooser.selectedFile.absolutePath
        } else {
            null
        }
    }

    @Serializable
    data class ExcelRow(val data: Map<String, @Contextual Any>)

    private fun readExcelFile(filePath: String): List<ExcelRow> {
        val inputStream = FileInputStream(File(filePath))
        val workbook = WorkbookFactory.create(inputStream)
        val sheet = workbook.getSheetAt(0)
        val headerRow = sheet.getRow(0)
        val rows = mutableListOf<ExcelRow>()

        for (rowIndex in 1 until sheet.physicalNumberOfRows) {
            val row = sheet.getRow(rowIndex)
            val rowData = mutableMapOf<String, Any>()
            for (cellIndex in 0 until headerRow.physicalNumberOfCells) {
                val headerCell = headerRow.getCell(cellIndex)
                val dataCell = row.getCell(cellIndex)

                val cellValue: Any = when {
                    dataCell.cellType == org.apache.poi.ss.usermodel.CellType.NUMERIC && dataCell.numericCellValue % 1 == 0.0 -> dataCell.numericCellValue.toInt()

                    //phép toán
                    dataCell.cellType == org.apache.poi.ss.usermodel.CellType.FORMULA -> {
                        val evaluator = workbook.creationHelper.createFormulaEvaluator()
                        val evaluatedCell = evaluator.evaluate(dataCell)
                        if (evaluatedCell.cellType == org.apache.poi.ss.usermodel.CellType.NUMERIC && evaluatedCell.numberValue % 1 == 0.0) {
                            evaluatedCell.numberValue.toInt()
                        } else {
                            evaluatedCell.stringValue
                        }
                    }

                    // Chuyển đổi chuỗi "TRUE"/"FALSE" thành Boolean
                    dataCell.toString().uppercase() == "TRUE" -> true
                    dataCell.toString().uppercase() == "FALSE" -> false

                    // Chuyển đổi giá trị số nguyên từ chuỗi
                    dataCell.toString().matches(Regex("^[+-]?\\d+$")) -> dataCell.toString().toInt()

                    // Các trường hợp còn lại để dưới dạng chuỗi
                    else -> dataCell.toString()
                }

                rowData[headerCell.toString()] = cellValue

//                rowData[headerCell.toString()] = value
            }
            rows.add(ExcelRow(data = rowData))
        }

        workbook.close()
        return rows
    }

    private fun writeJsonToFile(data: List<ExcelRow>, outputFilePath: String): String {
        val json = buildString {
            append(
                "{\n" +
                        "    \"code\": 200,\n" +
                        "    \"data\": [\n"
            )
            data.forEachIndexed { index, row ->
                append("  {\n")
                row.data.forEach { (key, value) ->
                    append("    \"$key\": ")
                    when (value) {
                        is Boolean -> append("$value")
                        is Int -> append("$value")
                        is Long -> append("$value")
                        else -> append("\"$value\"")
                    }
                    append(",\n")
                }
                deleteCharAt(lastIndex - 1)
                append("  }")
                if (index < data.size - 1) append(",")
                append("\n")
            }
            append("] }")
        }
        File(outputFilePath).writeText(json)

        return json
    }

    private fun openFileInExplorer(file: File) {
        try {
            if (Desktop.isDesktopSupported()) {
                val desktop = Desktop.getDesktop()
                if (desktop.isSupported(Desktop.Action.BROWSE_FILE_DIR)) {
                    desktop.browseFileDirectory(file)
                } else {
                    if (System.getProperty("os.name").toLowerCase().contains("win")) {
                        Runtime.getRuntime().exec("explorer.exe /select,${file.absolutePath}")
                    } else if (System.getProperty("os.name").toLowerCase().contains("mac")) {
                        Runtime.getRuntime().exec(arrayOf("open", "-R", file.absolutePath))
                    } else {
                        Runtime.getRuntime().exec(arrayOf("xdg-open", file.parent))
                    }
                }
            } else {
                println("Desktop không hỗ trợ mở File Explorer.")
            }
        } catch (e: Exception) {
            e.printStackTrace()
        }
    }


}


fun main() {
    val app = App()
    SwingUtilities.invokeLater {
        app.run()
    }
}


class GradientTextPanel : JPanel() {
    val asciiArt = """
                 
                              _                          _______      __
                             | |                        / ____\ \    / /
                             | |     ___  _ __   __ _  | |     \ \  / / 
        Create by            | |    / _ \| '_ \ / _` | | |      \ \/ /                       
                             | |___| (_) | | | | (_| | | |____   \  /   
                             |______\___/|_| |_|\__, |  \_____|   \/    
                                                 __/ |                  
                                                |___/                   
        """.trimIndent()


    override fun paintComponent(g: Graphics) {
        super.paintComponent(g)
        val g2d = g as Graphics2D

        // Thiết lập gradient chạy từ trái sang phải
        val gradientPaint = GradientPaint(0f, 0f, Color.RED, width.toFloat(), 0f, Color.BLUE, true)
        g2d.paint = gradientPaint

        // Thiết lập font
        val font = Font(Font.MONOSPACED, Font.BOLD, 24)
        g2d.font = font

        // Lấy kích thước văn bản
        val fontMetrics = g2d.fontMetrics
        val x = (width - fontMetrics.stringWidth(asciiArt.split("\n").first())) / 2
        val y = fontMetrics.ascent

        // Vẽ văn bản bôi đậm với gradient
        asciiArt.split("\n").forEachIndexed { index, line ->
            g2d.drawString(line, x, y + index * fontMetrics.height)
        }
    }
}