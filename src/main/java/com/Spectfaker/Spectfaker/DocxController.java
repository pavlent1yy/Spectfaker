package com.Spectfaker.Spectfaker;

import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.multipart.MultipartFile;

import jakarta.servlet.http.HttpServletResponse;
import java.io.IOException;

@Controller
public class DocxController {

    private final DocxService docxService;

    public DocxController(DocxService docxService) {
        this.docxService = docxService;
    }

    @GetMapping("/")
    public String index() {
        return "upload";
    }

    @PostMapping("/upload")
    public void handleFileUpload(@RequestParam("file") MultipartFile file, HttpServletResponse response) throws IOException {
        docxService.processDocxFile(file, response);
    }



}
