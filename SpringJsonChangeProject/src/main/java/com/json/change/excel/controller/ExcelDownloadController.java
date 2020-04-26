package com.json.change.excel.controller;

import java.util.Locale;

import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;

import com.json.change.excel.service.ExcelService;

@Controller
public class ExcelDownloadController {

	@Autowired
	private ExcelService service;

	@RequestMapping(value = "/downloadExcelFile", method = RequestMethod.POST)
    public String downloadExcelFile(Model model, String jsonData) {
		SXSSFWorkbook workbook = service.excelFileDownloadProcess(service.makeJsonList(jsonData));
		model.addAttribute("locale", Locale.KOREA);
		model.addAttribute("workbook", workbook);
		model.addAttribute("workbookName", "형태소분석이라는데 뭔지는 모르겠다.");
		
        return "excelDownloadView";
    }
}
