package com.app.service;

import com.app.dao.EmployeeRepository;
import com.app.pojos.Employee;
import jakarta.servlet.ServletOutputStream;
import jakarta.servlet.http.HttpServletResponse;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.apache.poi.hssf.usermodel.*;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.io.InputStream;
import java.net.http.HttpResponse;
import java.util.Iterator;
import java.util.List;

@Service
public class EmployeeService implements IEmployeeService {
    @Autowired
    private EmployeeRepository empRepo;

    public void generateExcel(HttpServletResponse response) throws IOException {

        List<Employee> employees=empRepo.findAll();
        HSSFWorkbook workbook=new HSSFWorkbook();
        HSSFSheet sheet =workbook.createSheet("Employee Info");

        HSSFCellStyle headerStyle = workbook.createCellStyle();
        headerStyle.setFillForegroundColor(IndexedColors.BLUE.getIndex());
       headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        HSSFFont font = workbook.createFont();
        font.setBold(true);
        font.setColor(IndexedColors.WHITE.getIndex());
        headerStyle.setFont(font);

        HSSFRow row =sheet.createRow(0);
        row.createCell(0).setCellValue("Emp Id");
        row.createCell(1).setCellValue("First Name");
        row.createCell(2).setCellValue("Last Name");
        row.createCell(3).setCellValue("Email");
        row.createCell(4).setCellValue("Department");
        row.createCell(5).setCellValue("Salary");
        for (int i = 0; i < 6; i++) {
            row.getCell(i).setCellStyle(headerStyle);
        }

        int dataIndex=1;

        for (Employee employee:employees){
            HSSFRow dataRow=sheet.createRow(dataIndex);
            dataRow.createCell(0).setCellValue(employee.getId());
            dataRow.createCell(1).setCellValue(employee.getFirstName());
            dataRow.createCell(2).setCellValue(employee.getLastName());
            dataRow.createCell(3).setCellValue(employee.getEmail());
            dataRow.createCell(4).setCellValue(employee.getDepartment());
            dataRow.createCell(5).setCellValue(employee.getSalary());
            dataIndex++;
        }
        ServletOutputStream ops = response.getOutputStream();
        workbook.write(ops);
        workbook.close();
        ops.close();

    }

    @Override
    public void generateCsv(HttpServletResponse response) throws IOException {
        List<Employee> employees=empRepo.findAll();
        response.getWriter().write("Emp Id,First Name,Last Name,Email,Department,Salary\n");
        for(Employee emp:employees){
            response.getWriter().write(emp.getId()+","
                            +emp.getFirstName()+","
                            +emp.getLastName()+","
                            +emp.getEmail()+","
                            +emp.getDepartment()+","
                            +emp.getSalary()+"\n"
            );
        }

    }

    @Override
    public void uploadExcelFile(MultipartFile file) throws IOException {
        InputStream inputStream  =file.getInputStream();
        Workbook workbook = new HSSFWorkbook(inputStream);
        Sheet sheet =workbook.getSheetAt(0);

        Iterator<Row> rowIterator = sheet.iterator();

        if(rowIterator.hasNext()){
            rowIterator.next();
        }

        while (rowIterator.hasNext()){
            Row row=rowIterator.next();
            Long id= (long) row.getCell(0).getNumericCellValue();
            String firstName=row.getCell(1).getStringCellValue();
            String lastName=row.getCell(2).getStringCellValue();
            String email=row.getCell(3).getStringCellValue();
            String department=row.getCell(4).getStringCellValue();
            Double salary=row.getCell(5).getNumericCellValue();
            Employee emp=new Employee();
            emp.setId(id);
            emp.setFirstName(firstName);
            emp.setLastName(lastName);
            emp.setEmail(email);
            emp.setDepartment(department);
            emp.setSalary(salary);
            empRepo.save(emp);
        }

        workbook.close();
    }
}
