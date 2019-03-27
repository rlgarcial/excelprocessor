/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package rlgarcial.excelprocessor.validator;


import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;

/**
 *
 * @author rlgarcial
 */
public class RowValidator {
    
    public int isBlankRow(HSSFRow row, int numeCell) {
        int cont = 0;
        for (int x = 0; x < numeCell; x++) {
            if (isBlank(row, x)) {
                cont++;
            }
        }
        return cont;
    }

    /**
     * Evalua el contenido de la fila para verificar si está en blanco.
     * 
     * @param row Row
     * @param i index
     * @return <code>true</code> si el valor de la fila está en blanco. <code>false</code> de otro modo.
     */
    public boolean isBlank(HSSFRow row, int i) {
        return ((row.getCell(i) == null) || (row.getCell(i).getCellType() == HSSFCell.CELL_TYPE_BLANK) || (row.getCell(i).toString().trim().equals("")));
    }
    
    /**
     * Obtiene el número de celdas vacías en la fila
     * @param row Fila
     * @param numeCell Número de columnas totales en la fila
     * @return <code>int</code> Número de c vacías en la fila
     */
    public int getNumOfBlankCellsInRow(XSSFRow row, int numeCell) {
        int cont = 0;
        for (int x = 0; x < numeCell; x++) {
            if (isBlank(row, x)) {
                cont++;
            }
        }
        
        return cont;
    }
    
    public boolean isBlankRow(XSSFRow row, int numeCell) {
        int cont = 0;
        for (int x = 0; x < numeCell; x++) {
            if (isBlank(row, x)) {
                cont++;
            }
        }
        
        return cont > 0;
    }

    /**
     * Evalua el contenido de la fila para verificar si está en blanco.
     * 
     * @param row Row
     * @param i index
     * @return <code>true</code> si el valor de la fila está en blanco. <code>false</code> de otro modo.
     */
    public boolean isBlank(XSSFRow row, int i) {
        return ((row.getCell(i) == null) || (row.getCell(i).getCellType() == XSSFCell.CELL_TYPE_BLANK) || (row.getCell(i).toString().trim().equals("")));
    }
    
    /*
    AGREGAR LISTA DE POSIBLES ERRORES
    */
}
