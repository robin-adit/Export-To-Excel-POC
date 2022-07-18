package poc.robin.excelExport.controller;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;
import poc.robin.excelExport.entity.Employee;
import poc.robin.excelExport.service.EmployeeService;
import poc.robin.excelExport.util.ExcelDocExportUtil;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;

@RestController
public class EmployeeFetchController {

    @Autowired
    private EmployeeService employeeService;

    //To Generate a Json File of All Employees
    @GetMapping(value = "/getEmployees",produces = "application/json")
    public List<Employee> FetchEmployees()
    {
            List<Employee> employeeList = employeeService.fetchAllEmployees();
            employeeList.forEach(employee -> System.out.println(employee.toString()));
            return employeeList;
    }

    //Generate Excel file of ALL EMPLOYEES WITH FILTERS and NO SELECTION
    @GetMapping("/exportEmployees")
    public void exportToExcel(HttpServletResponse response) throws IOException {
        response.setContentType("application/octet-stream");
        DateFormat dateFormatter = new SimpleDateFormat("yyyy-MM-dd_HH:mm:ss");
        String currentDateTime = dateFormatter.format(new Date());

        String headerKey = "Content-Disposition";
        String headerValue = "attachment; filename=Employees_" + currentDateTime + ".xlsx";
        response.setHeader(headerKey, headerValue);

        List<Employee> employeeList = employeeService.fetchAllEmployees();

        ExcelDocExportUtil excelDocExporter = new ExcelDocExportUtil(employeeList);

        excelDocExporter.export(response,false);
    }

    //Generate Excel file of ALL EMPLOYEES WITH FILTERS and a SELECTION
    //Employees with Manager_id=1
    //TODO : Customize as per requirement
    @GetMapping("/exportEmployeesCustom")
    public void exportToExcelCustom(HttpServletResponse response) throws IOException
    {
        response.setContentType("application/octet-stream");
        DateFormat dateFormatter = new SimpleDateFormat("yyyy-MM-dd_HH:mm:ss");
        String currentDateTime = dateFormatter.format(new Date());

        String headerKey = "Content-Disposition";
        String headerValue = "attachment; filename=Employees_" + currentDateTime + ".xlsx";
        response.setHeader(headerKey, headerValue);

        List<Employee> employeeList = employeeService.fetchAllEmployees();

        ExcelDocExportUtil excelDocExporter = new ExcelDocExportUtil(employeeList);

        excelDocExporter.export(response,true);
    }

}