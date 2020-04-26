package com.json.change.excel.service;

import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.json.simple.JSONArray;
import org.json.simple.JSONObject;
import org.json.simple.parser.JSONParser;
import org.json.simple.parser.ParseException;
import org.springframework.stereotype.Service;

import com.json.change.excel.dto.JsonChageDTO;

@Service
public class ExcelService {

	/**
	 * 엑셀 파일로 만들 리스트 생성
	 * @param jsonData
	 * @return 엑셀파일리스트
	 */
	public List<JsonChageDTO> makeJsonList(String jsonData) {
		List<JsonChageDTO> list = new ArrayList<JsonChageDTO>();
		JSONParser jsonParser = new JSONParser();
		JSONObject jsonObject;
		JsonChageDTO dto = new JsonChageDTO();
		
		try {
			jsonObject = (JSONObject) jsonParser.parse(jsonData);
			JSONArray rootJsonArray = (JSONArray) jsonObject.get("sentence");
			JSONArray morpJsonArray;
			JSONObject rootInArray;
			JSONObject morpInArray;
			
			String text = "";
			String morpJsonArrayInLemma = "";
			String morpJsonArrayInType = "";
			String morpJsonArrayInTypeStr = "";
			
			for(int i=0; i<rootJsonArray.size(); i++){
				rootInArray = (JSONObject) rootJsonArray.get(i);
				morpJsonArray = (JSONArray) rootInArray.get("morp");
				dto = new JsonChageDTO();
				
				text = (String) rootInArray.get("text");
				
				for(int j=0; j<morpJsonArray.size(); j++) {
					morpInArray =  (JSONObject) morpJsonArray.get(j);
					
					if(j == morpJsonArray.size() - 1) {
						morpJsonArrayInLemma += (String) morpInArray.get("lemma");
						morpJsonArrayInType += (String) morpInArray.get("type");
						morpJsonArrayInTypeStr += (String) morpInArray.get("typeStr");
					}else {
						morpJsonArrayInLemma += (String) morpInArray.get("lemma") +  " + ";
						morpJsonArrayInType += (String) morpInArray.get("type") + " + ";
						morpJsonArrayInTypeStr += (String) morpInArray.get("typeStr") + " + ";
					}
				}
				
				dto.setText(text);
				dto.setLemma(morpJsonArrayInLemma);
				dto.setType(morpJsonArrayInType);
				dto.setTypeStr(morpJsonArrayInTypeStr);
				list.add(dto);
			}
			
		} catch (ParseException e) {
			e.printStackTrace();
		}
		return list;
	}

	/**
	 * 리스트를 간단한 엑셀 워크북 객체로 생성
	 * 
	 * @param list
	 * @return 생성된 워크북
	 */
	public SXSSFWorkbook makeSimpleFruitExcelWorkbook(List<JsonChageDTO> list) {
		SXSSFWorkbook workbook = new SXSSFWorkbook();
		// 시트 생성
		SXSSFSheet sheet = workbook.createSheet("종광씨가 부탁을해서 만듬 근데 엑셀명은 모르겠어서 그냥 적음");
		JsonChageDTO dto = new JsonChageDTO();

		// 시트 열 너비 설정
		sheet.setColumnWidth(0, 1500);
		sheet.setColumnWidth(0, 3000);
		sheet.setColumnWidth(0, 3000);
		sheet.setColumnWidth(0, 1500);

		// 헤더 행 생
		Row headerRow = sheet.createRow(0);
		// 해당 행의 첫번째 열 셀 생성
		Cell headerCell = headerRow.createCell(0);
		headerCell.setCellValue("원인문장");
		// 해당 행의 두번째 열 셀 생성
		headerCell = headerRow.createCell(1);
		headerCell.setCellValue("lemma");
		// 해당 행의 세번째 열 셀 생성
		headerCell = headerRow.createCell(2);
		headerCell.setCellValue("type");
		// 해당 행의 네번째 열 셀 생성
		headerCell = headerRow.createCell(3);
		headerCell.setCellValue("typeStr");

		// 과일표 내용 행 및 셀 생성
		Row bodyRow = null;
		Cell bodyCell = null;
		for (int i = 0; i < list.size(); i++) {
			dto = list.get(i);

			// 행 생성
			bodyRow = sheet.createRow(i + 1);
			// 데이터 - text
			bodyCell = bodyRow.createCell(0);
			bodyCell.setCellValue(dto.getText());
			// 데이터 - lemma
			bodyCell = bodyRow.createCell(1);
			bodyCell.setCellValue(dto.getLemma());
			// 데이터 - type
			bodyCell = bodyRow.createCell(2);
			bodyCell.setCellValue(dto.getType());
			// 데이터 - typeStr
			bodyCell = bodyRow.createCell(3);
			bodyCell.setCellValue(dto.getTypeStr());
		}
		return workbook;
	}

	/**
	 * 생성한 엑셀 워크북을 컨트롤레에서 받게해줄 메소드
	 * 
	 * @param list
	 * @return
	 */
	public SXSSFWorkbook excelFileDownloadProcess(List<JsonChageDTO> list) {
		return this.makeSimpleFruitExcelWorkbook(list);
	}

}
