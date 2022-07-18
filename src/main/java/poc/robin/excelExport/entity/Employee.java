package poc.robin.excelExport.entity;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import lombok.ToString;

import javax.persistence.Column;
import javax.persistence.Entity;
import javax.persistence.Id;
import javax.persistence.Table;
import java.util.Date;

@Entity
@Table(name = "employees")
@NoArgsConstructor
@AllArgsConstructor
@Data
@ToString //Added just for testing. TODO : Remove
public class Employee {
    @Id
    @Column(name="employee_id")
    private Long employeeId;
    @Column(name="first_name")
    private String firstName;
    @Column(name="last_name")
    private String last_name;
    @Column(name="email")
    private String email;
    @Column(name="phone")
    private String phone;
    @Column(name="hire_date")
    private Date hireDate;
    @Column(name="manager_id")
    private Long managerId;
    @Column(name="job_title")
    private String jobTitle;
}