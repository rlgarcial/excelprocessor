/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package rlgarcial.excelprocessor.writer;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Map;
import java.util.Set;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author rlgarcial
 */
public class XSSFWorkBookWriter {
    
    private XSSFWorkbook workbook;
    
    private List<Map<String, Object[]>> workbookSheetDataList;
    
    public XSSFWorkbook writeData(XSSFWorkbook workbook,  Map<String, Object[]>... sheetData) {
        this.initializeVars(workbook, sheetData);
        
        
        return null;
    }
    
    public void writeSheetData(Map<String, Object[]> sheetData) {
        Set<String> keyset = sheetData.keySet();
        int rownum = 0;
        for(String key : keyset) {
            
        }
    }
    
    public void writeRowData() {
        
    }
    
    private void initializeVars(XSSFWorkbook workbook,  Map<String, Object[]>... sheetData) {
        this.workbook = workbook;
        this.workbookSheetDataList = new ArrayList<>();
        this.workbookSheetDataList.addAll(Arrays.asList(sheetData));
    }

    public XSSFWorkbook getWorkbook() {
        return workbook;
    }
    
}
