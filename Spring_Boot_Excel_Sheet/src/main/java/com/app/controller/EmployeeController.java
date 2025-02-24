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

    @GetMapping("/excel")
    public void generateExcelSheet(HttpServletResponse response) throws IOException {
      response.setContentType("application/octet-stream");
      response.setHeader("Content-Disposition","attachment;filename=employee.xls");
      empService.generateExcel(response);
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
