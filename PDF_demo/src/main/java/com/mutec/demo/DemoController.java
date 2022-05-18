package com.mutec.demo;

import com.itextpdf.text.*;
import com.itextpdf.text.pdf.BaseFont;
import com.itextpdf.text.pdf.PdfPCell;
import com.itextpdf.text.pdf.PdfPTable;
import com.itextpdf.text.pdf.PdfWriter;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.core.io.ClassPathResource;
import org.springframework.util.FileCopyUtils;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.servlet.ModelAndView;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import javax.servlet.http.HttpSession;
import java.io.*;
import java.net.URLEncoder;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.LinkedList;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

@RestController
public class DemoController {

    @GetMapping("/pdf")
    public ModelAndView getPDF(ModelAndView mav,
                               HttpServletRequest req) throws IOException {
        try {
            String filePath = new ClassPathResource("files").getURI().getPath();
            File dir = new File(filePath);
            File files[] = dir.listFiles();

            mav.addObject("fileList", files);
        } catch (IOException e) {

        }

        mav.setViewName("list");
        return mav;
    }

    @GetMapping("/fileLsit/{list}")
    public void postPDF(HttpServletRequest req,
                          HttpServletResponse res,
                          @PathVariable("list") String list) {
        try {
            List<byte[]> bytesList = new ArrayList<>();
            for (String file : list.split(",")) {
                byte[] bytes = makePDF(file, res);
                bytesList.add(bytes);
            }
            filDown(res, bytesList); //파일다운로드
        } catch (IOException e) {
        }
    }
    public byte[] makePDF(String fileNm, HttpServletResponse rep) throws IOException {
        ByteArrayOutputStream baos = new ByteArrayOutputStream();

        try {
            String filePath = new ClassPathResource("files").getURI().getPath();
            File file = new File(filePath + "/" + fileNm);
            FileInputStream input_document = new FileInputStream(file);
            XSSFWorkbook my_xls_workbook = new XSSFWorkbook(new FileInputStream(file));
            XSSFSheet my_worksheet = my_xls_workbook.getSheetAt(0);

            Document iText_xls_2_pdf = new Document();
            PdfWriter.getInstance(iText_xls_2_pdf, baos);
            iText_xls_2_pdf.open();

            int numberOfCells = my_worksheet.getRow(0).getPhysicalNumberOfCells();
            PdfPTable my_table = new PdfPTable(numberOfCells);

            BaseFont bf = BaseFont.createFont("font/malgun.ttf", BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
            Font font = new Font(bf, 12, Font.BOLD|Font.UNDERLINE, BaseColor.BLACK);

            for (int rowIdx = 0; rowIdx <= my_worksheet.getLastRowNum(); rowIdx++) {
                PdfPCell table_cell;
                XSSFRow row = my_worksheet.getRow(rowIdx);
                if (row == null) {
                    PdfPCell blankRow = new PdfPCell(new PdfPTable(numberOfCells));
                    my_table.addCell(blankRow);
                    continue;
                }
                for (int cellIdx = 0; cellIdx < numberOfCells; cellIdx++) {
                    Cell cell = row.getCell(cellIdx, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
                    if(cell == null) {
                        table_cell= new PdfPCell(new Paragraph(" ", font));
                    } else {
                        switch (cell.getCellTypeEnum()) {
                            case NUMERIC:
                                table_cell=new PdfPCell(new Paragraph(String.valueOf(cell.getNumericCellValue()), font));
                                break;
                            default:
                                table_cell=new PdfPCell(new Paragraph(cell.getStringCellValue(), font));
                                break;
                        }
                    }
                    my_table.addCell(table_cell);
                }
            }

            iText_xls_2_pdf.add(my_table);

            iText_xls_2_pdf.close();
            input_document.close();

            baos.flush();
            baos.close();

            return baos.toByteArray();
        } catch (IOException | DocumentException e) {
            return null;
        }
    }

    public void filDown(HttpServletResponse response,
                        List<byte[]> bytesList) throws IOException {

        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        ZipOutputStream zout = new ZipOutputStream(baos);

        int idx = 0;
        for (byte[] bytes : bytesList) {
            ZipEntry zip = new ZipEntry(Integer.toString(idx++) + ".pdf");
            zout.putNextEntry(zip);
            zout.write(bytes);
            zout.closeEntry();
        }

        zout.close();
        response.setHeader("Content-Disposition", "attachment; filename=\"DATA.ZIP\"");
        response.setContentType("application/zip");
        response.getOutputStream().write(baos.toByteArray());
        response.flushBuffer();

        baos.close();
    }
}
