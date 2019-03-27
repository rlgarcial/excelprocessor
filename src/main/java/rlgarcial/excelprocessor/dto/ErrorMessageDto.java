/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package rlgarcial.excelprocessor.dto;

import java.io.Serializable;
import java.util.Objects;

/**
 *
 * @author rlgarcial
 */
public class ErrorMessageDto implements Serializable {
    
    private String sheetName;
    private int cellNumber;
    private int rowNumber;
    private String column;
    private String message;

    public ErrorMessageDto(int cellNumber, int rowNumber, String column, String message) {
        this.cellNumber = cellNumber;
        this.rowNumber = rowNumber;
        this.column = column;
        this.message = message;
    }

    
    
    public ErrorMessageDto(String sheetName, int cellNumber, int rowNumber, String column, String message) {
        this(cellNumber, rowNumber, column, message);
        this.sheetName = sheetName;
    }

    @Override
    public int hashCode() {
        int hash = 5;
        hash = 67 * hash + Objects.hashCode(this.sheetName);
        hash = 67 * hash + this.cellNumber;
        hash = 67 * hash + this.rowNumber;
        hash = 67 * hash + Objects.hashCode(this.column);
        return hash;
    }

    @Override
    public boolean equals(Object obj) {
        if (this == obj) {
            return true;
        }
        if (obj == null) {
            return false;
        }
        if (getClass() != obj.getClass()) {
            return false;
        }
        final ErrorMessageDto other = (ErrorMessageDto) obj;
        if (this.cellNumber != other.cellNumber) {
            return false;
        }
        if (this.rowNumber != other.rowNumber) {
            return false;
        }
        if (!Objects.equals(this.sheetName, other.sheetName)) {
            return false;
        }
        if (!Objects.equals(this.column, other.column)) {
            return false;
        }
        return true;
    }
    
    public String getSheetName() {
        return sheetName;
    }

    public void setSheetName(String sheetName) {
        this.sheetName = sheetName;
    }
    
    public int getCellNumber() {
        return cellNumber;
    }

    public void setCellNumber(int cellNumber) {
        this.cellNumber = cellNumber;
    }

    public int getRowNumber() {
        return rowNumber;
    }

    public void setRowNumber(int rowNumber) {
        this.rowNumber = rowNumber;
    }

    public String getColumn() {
        return column;
    }

    public void setColumn(String Column) {
        this.column = Column;
    }

    public String getMessage() {
        return message;
    }

    public void setMessage(String message) {
        this.message = message;
    }

    @Override
    public String toString() {
        return "ErrorMessageDto{" + "sheetName=" + sheetName + ", cellNumber=" + cellNumber + ", rowNumber=" + rowNumber + ", column=" + column + ", message=" + message + '}';
    }
    
}
