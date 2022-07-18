package poc.robin.excelExport.util;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTAutoFilter;


import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTFilterColumn;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTFilters;
import poc.robin.excelExport.entity.Employee;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.util.Date;
import java.util.List;

public class ExcelDocExportUtil {

    private XSSFWorkbook workbook;
    private XSSFSheet sheet;
    private List<Employee> employeeList;
    int firstRow, lastRow, firstColumn,lastColumn;

    public ExcelDocExportUtil(List<Employee> employeeList) {
        this.employeeList = employeeList;
        workbook = new XSSFWorkbook();
        firstRow = lastRow = firstColumn = lastColumn = 0;
    }

    private void writeHeaderLine() {
        sheet = workbook.createSheet("Employees");

        Row row = sheet.createRow(0);

        CellStyle style = workbook.createCellStyle();
        XSSFFont font = workbook.createFont();
        font.setBold(true);
        font.setFontHeight(18);
        style.setFillBackgroundColor(IndexedColors.BLACK.getIndex());
        style.setFont(font);

        createCell(row, 0, "Employee ID", style);
        createCell(row, 1, "First Name", style);
        createCell(row, 2, "Last Name", style);
        createCell(row, 3, "Email Address", style);
        createCell(row, 4, "Phone Number", style);
        createCell(row, 5, "Job Title", style);
        createCell(row, 6, "Hire Date", style);
        createCell(row, 7, "Manager ID", style);

    }

    private void createCell(Row row, int columnCount, Object value, CellStyle style) {
        sheet.autoSizeColumn(columnCount);
        Cell cell = row.createCell(columnCount);

        /*
        System.out.print("cell.getRow()=" + cell.getRowIndex()
                + " cell.getColumnIndex()=" + cell.getColumnIndex());
        if(null!= value)
            System.out.print(" value.getClass()" + value.getClass()
                + " value.toString()" + value.toString());
        System.out.println();
        */

        if (value instanceof Integer) {
            ((Cell) cell).setCellValue((Integer) value);
        } else if (value instanceof Long) {
            ((Cell) cell).setCellValue((Long) value);
        }else if (value instanceof Date) {
            ((Cell) cell).setCellValue((Date) value);
        }
        else {
            ((Cell) cell).setCellValue((String) value);
        }
        ((Cell) cell).setCellStyle(style);
    }

    private void writeDataLines() {
        int rowCount = 1;

        CellStyle style = workbook.createCellStyle();
        XSSFFont font = workbook.createFont();
        font.setFontHeight(14);
        style.setFont(font);

        int columnCount = 0;
        for (Employee employee : employeeList) {
            Row row = sheet.createRow(rowCount++);
            columnCount = 0;

            createCell(row, columnCount++, employee.getEmployeeId(), style);
            createCell(row, columnCount++, employee.getFirstName(), style);
            createCell(row, columnCount++, employee.getLast_name(), style);
            createCell(row, columnCount++, employee.getEmail(), style);
            createCell(row, columnCount++, employee.getPhone(), style);
            createCell(row, columnCount++, employee.getJobTitle(), style);
            createCell(row, columnCount++, employee.getHireDate().toString().substring(0,10),style);
            createCell(row, columnCount++, employee.getManagerId(), style);

            lastColumn = columnCount; //TODO : Repeating - Find a better way to do this
        }

        lastRow = rowCount;
        firstColumn = 0;
        //sheet.setAutoFilter(new CellRangeAddress(firstRow, lastRow, firstColumn, --lastColumn));
        sheet.setAutoFilter(CellRangeAddress.valueOf("A1:H1"));
    }

    private void writeDataLines(Boolean customFlag) {
        int rowCount = 1;

        CellStyle style = workbook.createCellStyle();
        XSSFFont font = workbook.createFont();
        font.setFontHeight(14);
        style.setFont(font);

        int columnCount = 0;
        for (Employee employee : employeeList) {
            Row row = sheet.createRow(rowCount++);
            columnCount = 0;

            createCell(row, columnCount++, employee.getEmployeeId(), style);
            createCell(row, columnCount++, employee.getFirstName(), style);
            createCell(row, columnCount++, employee.getLast_name(), style);
            createCell(row, columnCount++, employee.getEmail(), style);
            createCell(row, columnCount++, employee.getPhone(), style);
            createCell(row, columnCount++, employee.getJobTitle(), style);
            createCell(row, columnCount++, employee.getHireDate().toString().substring(0,10),style);
            createCell(row, columnCount++, employee.getManagerId(), style);

            lastColumn = columnCount; //TODO : Repeating - Find a better way to do this
        }

        lastRow = rowCount;
        firstColumn = 0;

        System.out.println(sheet.getSheetName());
        CTAutoFilter ctAutoFilter = null;

        //Show information for only Manager ID = 1
        if(null != sheet.getCTWorksheet() && null == sheet.getCTWorksheet().getAutoFilter())
            ctAutoFilter = sheet.getCTWorksheet().addNewAutoFilter();
        else ctAutoFilter = sheet.getCTWorksheet().getAutoFilter();

        CTFilterColumn ctFilterColumn = ctAutoFilter.addNewFilterColumn();
        ctFilterColumn.setColId(7);
        CTFilters ctFilters = ctFilterColumn.addNewFilters();
        ctFilters.addNewFilter().setVal("1");
        for (Row row : sheet)
        {
            System.out.println(row.getRowNum());
            if(row.getRowNum() != 0)
            {
                for (Cell cell : row)
                {
                    if (7 == cell.getColumnIndex() && cell.getNumericCellValue() != 1)
                    {
/*                        System.out.println("Inside");
                        System.out.println("Column=" + cell.getColumnIndex()
                                + " Value="+cell.getNumericCellValue());
*/
                        XSSFRow roww = (XSSFRow) cell.getRow();
                        if (0 != roww.getRowNum())
                            roww.getCTRow().setHidden(true);
                    }
                }
            }

        }


    }


    public void export(HttpServletResponse response, Boolean customFlag) throws IOException {
        writeHeaderLine();

        if(!customFlag)
            writeDataLines();
        else
            writeDataLines(customFlag);

        ServletOutputStream outputStream = response.getOutputStream();
        workbook.write(outputStream);
        workbook.close();

        outputStream.close();
    }
}