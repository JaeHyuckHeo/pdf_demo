package com.mutec.demo;

import com.itextpdf.text.BaseColor;
import com.itextpdf.text.Document;
import com.itextpdf.text.DocumentException;
import com.itextpdf.text.Phrase;
import com.itextpdf.text.pdf.PdfPCell;
import com.itextpdf.text.pdf.PdfPTable;
import com.itextpdf.text.pdf.PdfWriter;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.core.io.ClassPathResource;

import java.io.*;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

public class ExcelToPDFConverter {

    public void getDoc(String excelFileName) throws IOException, DocumentException {
        String filePath = new ClassPathResource("files").getURI().getPath();
        FileInputStream fileInputStream = new FileInputStream(new File(filePath + "/" + excelFileName));
        Workbook workbook = new XSSFWorkbook(fileInputStream);

        Sheet sheet = workbook.getSheetAt(0);
        List<String> headerList = setHeader(sheet);

        for (int i = 1; i < sheet.getPhysicalNumberOfRows(); i++) {
            Document document = new Document();
            String fileName = "C:\\Temp\\pdf_" + i + ".pdf";
            PdfWriter.getInstance(document, new FileOutputStream(fileName));

            document.open();
            PdfPTable table = new PdfPTable(sheet.getRow(0).getPhysicalNumberOfCells());
            addPDFData(true, headerList, table);
            List<String> rowList = getRowData(i, sheet);
            if (rowList.isEmpty()) {
                document.add(table);
                document.close();
                Files.delete(Paths.get(fileName));
                continue;
            }
            addPDFData(false, rowList, table);
            document.add(table);
            document.close();
        }

    }

    public static List<String> setHeader(Sheet sheet) {
        return getRow(0, sheet);
    }

    public static List<String> getRowData(int index, Sheet sheet) {
        return getRow(index, sheet);
    }

    public static List<String> getRow(int index, Sheet sheet) {
        List<String> list = new ArrayList<>();

        for (Cell cell : sheet.getRow(index)) {
            switch (cell.getCellTypeEnum()) {
                case STRING:
                    list.add(cell.getStringCellValue());
                    break;
                case NUMERIC:
                    list.add(String.valueOf(cell.getNumericCellValue()));
                    break;
                case BOOLEAN:
                    list.add(String.valueOf(cell.getBooleanCellValue()));
                    break;
                case FORMULA:
                    list.add(cell.getCellFormula().toString());
                    break;
            }
        }

        return list;
    }

    private static void addPDFData(boolean isHeader, List<String> list, PdfPTable table) {
        list.stream()
                .forEach(column -> {
                    PdfPCell header = new PdfPCell();
                    if (isHeader) {
                        header.setBackgroundColor(BaseColor.LIGHT_GRAY);
                        header.setBorderWidth(2);
                    }
                    header.setPhrase(new Phrase(column));
                    table.addCell(header);
                });
    }



//            rep.setContentType("application/zip"); // application/octet-stream
//            rep.setHeader("Content-Disposition", "inline; filename=\"all.zip\"");
//            try (
//    ZipOutputStream zos = new ZipOutputStream(rep.getOutputStream())) {
//        for (int i = 0; i < 3; i++) {
//            ZipEntry ze = new ZipEntry("document-" + i + ".pdf");
//            zos.putNextEntry(ze);
//
//            // It would be nice to write the PDF immediately to zos.
//            // However then you must take care to not close the PDF (and zos),
//            // but just flush (= write all buffered).
//            //PdfWriter pw = PdfWriter.getInstance(document[i], zos);
//            //...
//            //pw.flush(); // Not closing pw/zos
//
//            // Or write the PDF to memory:
//            ByteArrayOutputStream baos = new ...
//            PdfWriter pw = PdfWriter.getInstance(document[i], baos);
//         ...
//            pw.close();
//            byte[] bytes = baos.toByteArray();
//            zos.write(baos, 0, baos.length);
//
//            zos.closeEntry();
//        }
//    }

}
