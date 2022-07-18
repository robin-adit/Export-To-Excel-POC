package poc.robin.excelExport.service;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import poc.robin.excelExport.entity.Employee;
import poc.robin.excelExport.repository.EmployeeRepository;

import java.util.List;

@Service
public class EmployeeService {

    @Autowired
    private EmployeeRepository employeeRepository;

    public List<Employee> fetchAllEmployees()
    {
        return employeeRepository.findAll();
    }

}
