package ksp;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class pdf {

    public static void main(String[] args) throws IOException {
        // Path to the Excel file containing the PDF paths
        String inputExcelPath ="D://steno//Input_steno_pdf.xlsx";//"D://steno//pdf.xlsx"
        // Path to output the extracted Excel data
        String outputExcelPath = "D://steno//Extracted.xlsx";//"D://steno//Extracted.xlsx"

        // Step 1: Read the Excel file containing PDF paths
        FileInputStream fis = new FileInputStream(new File(inputExcelPath));
        Workbook workbook = WorkbookFactory.create(fis);
        Sheet sheet = workbook.getSheetAt(0); // Assume the PDF paths are in the first sheet

        // Step 2: Create a new workbook to store extracted data
        Workbook outputWorkbook = new XSSFWorkbook();
        Sheet outputSheet = outputWorkbook.createSheet("Extracted Data");
        Sheet fullTextSheet = outputWorkbook.createSheet("Full Extracted Text");

        // Step 3: Create header row for output sheet
        Row headerRow = outputSheet.createRow(0);
        headerRow.createCell(0).setCellValue("Unit");

        // Step 4: Iterate over each row of the input Excel file to process PDF files
        int rowNum = 1; // Start writing data from row 1 in output sheet
        for (Row row : sheet) {
            // Assuming PDF file path is in the first column of the input Excel
            Cell cell = row.getCell(0);
            if (cell != null && cell.getCellType() == CellType.STRING) {
                String pdfFilePath = cell.getStringCellValue();
                System.out.println("Processing PDF: " + pdfFilePath);

                // Step 5: Extract text from the PDF
                String extractedText = extractTextFromPDF(pdfFilePath);
                if (extractedText == null || extractedText.isEmpty()) {
                    System.out.println("No text extracted from the PDF: " + pdfFilePath);
                    continue; // Skip to next PDF if no text is found
                }

           

                // Step 8: Write the entire extracted text into the full text sheet
                String[] lines = extractedText.split("\n");
                int lineNum = 0;
                for (String line : lines) {
                    Row textRow = fullTextSheet.createRow(lineNum++);
                    textRow.createCell(0).setCellValue(line.trim()); // Write each line of the text in a new row
                }
            }
        }

        // Step 9: Write the output to an Excel file
        try (FileOutputStream fileOut = new FileOutputStream(outputExcelPath)) {
            outputWorkbook.write(fileOut); // Write the workbook to the file
        }

        // Close the workbooks
        workbook.close();
        outputWorkbook.close();

        System.out.println("Data has been successfully written to " + outputExcelPath);
    }

    // Function to extract text from PDF
    private static String extractTextFromPDF(String pdfFilePath) {
        File pdfFile = new File(pdfFilePath);
        if (!pdfFile.exists()) {
            System.out.println("File not found: " + pdfFilePath);
            return null;
        }

        try (PDDocument document = PDDocument.load(pdfFile)) {
            PDFTextStripper pdfStripper = new PDFTextStripper();
            return pdfStripper.getText(document); // Extract text from PDF
        } catch (IOException e) {
            System.out.println("Error extracting text from PDF: " + e.getMessage());
            return null;
        }
    }

    // Function to extract data using regex
    private static String extractData(String text, String regex) {
        Pattern pattern = Pattern.compile(regex);
        Matcher matcher = pattern.matcher(text);
        if (matcher.find()) {
            return matcher.group(2).trim();  // Extract the number (group 2)
        }
        return "Not found";  // Return a default value like "Not found" if the match doesn't occur
    }
}

