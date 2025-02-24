package com.app.service;

import jakarta.servlet.http.HttpServletResponse;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;

public interface IEmployeeService {

    public void generateExcel(HttpServletResponse response) throws IOException;
    public void uploadExcelFile(MultipartFile file)throws IOException;


}
