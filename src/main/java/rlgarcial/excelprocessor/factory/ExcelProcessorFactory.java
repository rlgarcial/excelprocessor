/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package rlgarcial.excelprocessor.factory;

import java.io.IOException;
import java.io.InputStream;
import java.util.Iterator;
import org.apache.commons.io.FilenameUtils;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author rlgarcial
 */
public class ExcelProcessorFactory {

    /**
     * Crea una nueva instancia de <code>XSSFWorkbook</code>
     *
     * @param input
     * @return <code>XSSFWorkbook</code>
     * @throws IOException
     */
    public XSSFWorkbook createXSSFWorkbook(InputStream input) throws IOException {
        return new XSSFWorkbook(input);
    }

    /**
     * Crea una nueva instancia de <code>HSSFWorkbook</code>
     *
     * @param input
     * @return <code>HSSFWorkbook</code>
     * @throws IOException
     */
    public HSSFWorkbook createHSSFWorkbook(InputStream input) throws IOException {
        return new HSSFWorkbook(input);
    }

    /**
     * Crea un <code>Iterator</code> para una hoja Ecxel de un libro con extensión .xls
     *
     * @param input
     * @return <code>Iterator</code>
     * @throws IOException
     */
    private Iterator createHSSFRowIterator(InputStream input) throws IOException {
        HSSFWorkbook workbook = this.createHSSFWorkbook(input);
        HSSFSheet sheet = workbook.getSheetAt(0);
        Iterator rowIterator = sheet.rowIterator();

        return rowIterator;
    }

    /**
     * Crea un <code>Iterator</code> para una hoja Ecxel de un libro con extensión .xlsx
     *
     * @param input
     * @return <code>Iterator</code>
     * @throws IOException
     */
    private Iterator createXSSFRowIterator(InputStream input) throws IOException {
        XSSFWorkbook workbook = this.createXSSFWorkbook(input);
        XSSFSheet sheet = workbook.getSheetAt(0);
        Iterator rowIterator = sheet.rowIterator();

        return rowIterator;
    }

    /**
     * Crea un <code>Iterator</code> para la hoja Excel
     * @param input
     * @param filename
     * @return
     * @throws IOException 
     */
    public Iterator createRowIterator(InputStream input, String filename) throws IOException {
        String extension = FilenameUtils.getExtension(filename);
        switch (extension) {
            case "xls":
                return createHSSFRowIterator(input);
            case "xlsx":
                return createXSSFRowIterator(input);
        }

        return null;
    }

    private int getCountOfHSSFSheetColumns(InputStream input) throws IOException {
        HSSFWorkbook workbook = this.createHSSFWorkbook(input);
        HSSFSheet sheet = workbook.getSheetAt(0);
        HSSFRow header = (HSSFRow) sheet.getRow(0);
        
        return header.getLastCellNum();
    }

    private int getCountOfXSSFSheetColumns(InputStream input) throws IOException {
        XSSFWorkbook workbook = this.createXSSFWorkbook(input);
        XSSFSheet sheet = workbook.getSheetAt(0);
        XSSFRow header = (XSSFRow) sheet.getRow(0);

        return header.getLastCellNum();
    }

    /**
     * Devuelve la cantidad de columnas del cabecero de la hoja
     * 
     * @param input
     * @param filename
     * @return
     * @throws IOException 
     */
    public int getCountOfHeaderColumns(InputStream input, String filename) throws IOException {
        String extension = FilenameUtils.getExtension(filename);
        switch (extension) {
            case "xls":
                return getCountOfHSSFSheetColumns(input);
            case "xlsx":
                return getCountOfXSSFSheetColumns(input);
        }
        
        return -1;
    }
    
    public boolean areSheetColumnsCountCorrect(InputStream input, String filename, int correctSheetColumns) throws IOException {
        return this.getCountOfHeaderColumns(input, filename) == correctSheetColumns;
    }

}
