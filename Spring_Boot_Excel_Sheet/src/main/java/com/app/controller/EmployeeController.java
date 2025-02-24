package com.app.controller;

import com.app.dao.EmployeeRepository;
import com.app.service.EmployeeService;
import jakarta.servlet.http.HttpServletResponse;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;
import java.io.IOException;

@RequestMapping("/employee")
@RestController
public class EmployeeController {

    @Autowired
    EmployeeService empService;

    @Autowired
    EmployeeRepository empRepo;

    @GetMapping("/file")
    public void generateExcelSheet(@RequestParam(value = "format" ,defaultValue ="excel") String format, HttpServletResponse response) throws IOException {
      if("excel".equalsIgnoreCase(format)){
          response.setContentType("application/vnd.ms-excel");
          response.setHeader("Content-Disposition","attachment;filename=employee.xls");
          empService.generateExcel(response);
      } else if ("csv".equalsIgnoreCase(format)) {
          response.setContentType("text/csv");
          response.setHeader("Content-Disposition", "attachment;filename=employee.csv");
          empService.generateCsv(response);
      }else {
          response.setStatus(HttpServletResponse.SC_BAD_REQUEST);
          response.getWriter().write("Invalid format. Please specify 'excel' or 'csv'.");
      }

    }

    @PostMapping("/upload/excel")
    public ResponseEntity<String> uploadExcelFile(@RequestParam("file")MultipartFile file){
        try {
            empService.uploadExcelFile(file);
            return ResponseEntity.ok("file upload successfully!!!");

        }catch (IOException ex){
            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).body("file upload failed!!!");

        }

    }


}
