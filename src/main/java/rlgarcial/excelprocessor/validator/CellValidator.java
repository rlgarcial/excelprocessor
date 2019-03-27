/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package rlgarcial.excelprocessor.validator;

/**
 *
 * @author rlgarcial
 */
import java.io.Serializable;
import java.text.DecimalFormat;
import java.util.ArrayList;
import java.sql.Date;
import java.util.List;
import javax.annotation.PostConstruct;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFCell;
import rlgarcial.excelprocessor.dto.ErrorMessageDto;


/**
 *
 * @author rlgarcial
 */
public class CellValidator implements Serializable {

    private List<ErrorMessageDto> HSSFCellErrorMessages;
    private List<ErrorMessageDto> XSSFCellErrorMessages;

    @PostConstruct
    public void init() {
        HSSFCellErrorMessages = new ArrayList<>();
        XSSFCellErrorMessages = new ArrayList<>();
    }

    /**
     * Limpia los mensajes de error tipo HSSFCellErrorMessages
     */
    public void clearHSSFCellErrorMessages() {
        HSSFCellErrorMessages.clear();
    }

    /**
     * Limpia los mensajes de error tipo XSSFCellErrorMessages
     */
    public void clearXSSFCellErrorMessages() {
        XSSFCellErrorMessages.clear();
    }

    /**
     * Limpia todos los mensajes de error
     */
    public void clearAllMessages() {
        XSSFCellErrorMessages.clear();
        HSSFCellErrorMessages.clear();
    }

    /**
     * Convierte el valor de la celda en número decimal
     *
     * @param cell Celda
     * @param isMandatory si el valor es true, se validará que el valor de la
     * celda exista forzosamente. En caso contrario, se agrega un error a la
     * lista de errores.
     * @return <code>Double</code> si el valor concuerda con los criterios ||
     * <code>null</code> de otro modo
     */
    public Double toDecimalValue(HSSFCell cell, boolean isMandatory) {
        String valueContainer;
        DecimalFormat decimalFormat = formatNumericCellValue(cell);
        // es valor mandatory
        if (isMandatory) {
            // es mandatory y no es nulo
            if (isNotBlankNumericCell(cell) && null != decimalFormat) {
                double cellDoubleValue = Double.parseDouble(cell.toString().trim());
                valueContainer = decimalFormat.format(cellDoubleValue);

                return Double.parseDouble(valueContainer);
            } // si el campo no es numérico
            else if (null != cell) {
                addErrorToNotNumericCellValue(cell);
            } // si el campo es mandatory y es nulo
            else {
                addErrorToMandatoryCellValue(cell);
            }
        } // es int no mandatory y tiene valor
        else if (isNotBlankNumericCell(cell) && null != decimalFormat) {
            double cellDoubleValue = Double.parseDouble(cell.toString().trim());
            valueContainer = decimalFormat.format(cellDoubleValue);

            return Double.parseDouble(valueContainer);
        } else {
            return null;
        }

        return null;
    }

    /**
     * Convierte el valor de la celda en número entero
     *
     * @param cell Celda
     * @param isMandatory si el valor es true, se validará que el valor de la
     * celda exista forzosamente. En caso contrario, se agrega un error a la
     * lista de errores.
     * @return <code>Double</code> si el valor concuerda con los criterios ||
     * <code>null</code> de otro modo.
     */
    public Long toIntegerValue(HSSFCell cell, boolean isMandatory) {
        String valueContainer;
        DecimalFormat decimalFormat = new DecimalFormat("#################");

        // es int mandatory
        if (isMandatory) {
            // es int mandatory y no es nulo
            if (isNotBlankNumericCell(cell)) {
                double cellDoubleValue = Double.parseDouble(cell.toString().trim());
                valueContainer = decimalFormat.format(cellDoubleValue);

                return Long.parseLong(valueContainer);
            } // si el campo no es numérico
            else if (null != cell) {
                addErrorToNotNumericCellValue(cell);
            } // siel campo es mandatory y es nulo
            else {
                addErrorToMandatoryCellValue(cell);
            }
            // es int no mandatory y tiene valor
        } else if (isNotBlankNumericCell(cell)) {
            double cellDoubleValue = Double.parseDouble(cell.toString().trim());
            valueContainer = decimalFormat.format(cellDoubleValue);

            return Long.parseLong(valueContainer);
        }
        return null;
    }

    public Date toDateValue(XSSFCell cell, String regex, String format) {
        if (cell.getRawValue().matches(regex)) {
            System.out.println("RAW DATE BUILT: " + getStringDateFormated(cell));
            return Date.valueOf(getStringDateFormated(cell));
        } else {
            addErrorToNotMatchDateRegexCellValue(cell, format);
        }
        return null;
    }

    private String getStringDateFormated(XSSFCell cell) {
        String rawDate = cell.getRawValue();
        StringBuilder dateBuilder = new StringBuilder();
        String year = rawDate.substring(0, 4);
        String month = rawDate.substring(4, 6);
        String day = rawDate.substring(6, 8);
        dateBuilder.append(year);
        dateBuilder.append("-");
        dateBuilder.append(month);
        dateBuilder.append("-");
        dateBuilder.append(day);

        return dateBuilder.toString();
    }
    
    private String getStringDateFormated(HSSFCell cell) {
        String rawDate = cell.getStringCellValue();
        StringBuilder dateBuilder = new StringBuilder();
        String year = rawDate.substring(0, 4);
        String month = rawDate.substring(4, 6);
        String day = rawDate.substring(6, 8);
        dateBuilder.append(day);
        dateBuilder.append("/");
        dateBuilder.append(month);
        dateBuilder.append("/");
        dateBuilder.append(year);

        return dateBuilder.toString();
    }

    public Date toDateValue(HSSFCell cell, String regex, String format) {
        if (cell.toString().matches(regex)) {
                return Date.valueOf(getStringDateFormated(cell));
            } else {
                addErrorToNotMatchDateRegexCellValue(cell, regex);
            }
        return null;
    }

    /**
     * Evalua que la celda no esté vacía y que su contenido sea numérico
     *
     * @param cell Celda
     * @return <code>true</code> si la celda no está vacía y su valor es
     * numérico. <code>false</code> de otro modo.
     */
    public boolean isNotBlankNumericCell(HSSFCell cell) {
        return !isBlank(cell) && cell.getCellType() == HSSFCell.CELL_TYPE_NUMERIC;
    }

    /**
     * Crea una instancia <code>DecimalFormat</code> para evaluar el contenido
     * de una celda con valores numéricos.
     *
     * @param cell Celda
     * @return <code>DecimalFormat</code> si el contenido de la celda no está
     * vacío y es de tipo numérico. <code>null</code> de otro modo.
     */
    public DecimalFormat formatNumericCellValue(HSSFCell cell) {
        if (isNotBlankNumericCell(cell)) {

            if (cell.toString().contains(".")) {
                return new DecimalFormat("##########.##");
            } else {
                return new DecimalFormat("##########");
            }
        }

        return null;
    }

    /**
     * Devuelve el valor de la celda en formato texto.
     *
     * @param cell Celda.
     * @return <code>String</code> Valor de la celda.
     */
    public String toTextValue(HSSFCell cell) {
        return cell.toString();
    }

    /**
     * Devuelve el valor de la celda en formato texto evaluando la longitud y si
     * éste es obligatorio.
     *
     * @param cell Celda
     * @param lenght Longitud máxima del valor de la celda
     * @param isMandatory si el valor es true, se validará que el valor de la
     * celda exista forzosamente. En caso contrario, se agrega un error a la
     * lista de errores.
     * @return <code>String</code> Valor de la celda.
     */
    public String toTextValue(HSSFCell cell, int lenght, boolean isMandatory) {
        // el campo es obligado
        if (isMandatory) {
            // es campo obligado y no es nulo
            if (!isBlank(cell) && cell.toString().length() <= lenght) {
                return cell.toString();
            } else {
                addErrorToMandatoryCellValue(cell);
            }
            // es campo no obligado pero tiene valor
        } else if (!isBlank(cell)) {
            return cell.toString();
        }
        return null;
    }

    /**
     * Devuelve el valor de la celda en formato texto evaluando la longitud, la
     * expresión regular y si éste es obligatorio.
     *
     * @param cell Celda
     * @param lenght Longitud máxima del valor de la celda
     * @param regex Expresión regular que debe cumplir el valor de la celda
     * @param isMandatory si el valor es true, se validará que el valor de la
     * celda exista forzosamente. En caso contrario, se agrega un error a la
     * lista de errores.
     *
     * @return <code>String</code> Valor de la celda
     * @throws IllegalStateException
     */
    public String toTextValue(HSSFCell cell, int lenght, String regex, boolean isMandatory) {
        // el campo es obligado
        if (isMandatory) {
            // es campo obligado y no es nulo
            if (!isBlank(cell) && cell.toString().length() <= lenght) {
                return cell.toString();
            } else {
                addErrorToMandatoryCellValue(cell);
            }
            // es campo no obligado pero tiene valor
        } else if (!isBlank(cell)) {
            if (cell.toString().matches(regex)) {
                return cell.toString();
            } else {
                addErrorToRegexNotMatchCellValue(cell);
            }
        }
        return null;
    }

    /**
     * Agrega un error si el valor es obligatorio
     *
     * @param cell Celda
     */
    private void addErrorToMandatoryCellValue(HSSFCell cell) {
        Object[] cellRelatedData = cellRelatedData(cell);
        ErrorMessageDto errorMessage = new ErrorMessageDto(
                cell.getSheet().getSheetName(),
                (int) cellRelatedData[0],
                (int) cellRelatedData[1],
                (String) cellRelatedData[2],
                "El campo no puede estar vacío"
        );

        if (-1 == HSSFCellErrorMessages.indexOf(errorMessage)) {
            HSSFCellErrorMessages.add(errorMessage);
        }
    }

    /**
     * Agrega un error si el valor no concuerda con una expresión regular
     *
     * @param cell
     */
    private void addErrorToRegexNotMatchCellValue(HSSFCell cell) {
        Object[] cellRelatedData = cellRelatedData(cell);
        ErrorMessageDto errorMessage = new ErrorMessageDto(
                cell.getSheet().getSheetName(),
                (int) cellRelatedData[0],
                (int) cellRelatedData[1],
                (String) cellRelatedData[2],
                "El campo no concuerda con el formato"
        );

        if (-1 == HSSFCellErrorMessages.indexOf(errorMessage)) {
            HSSFCellErrorMessages.add(errorMessage);
        }

    }

    /**
     * Agrega un error si la longitud del valor de la celda es mayor a la
     * longitud máxima indicada.
     *
     * @param cell Celda
     * @param lenght Longitud máxima
     */
    private void addErrorToOutOfLenghtCellValue(HSSFCell cell, int lenght) {
        Object[] cellRelatedData = cellRelatedData(cell);
        ErrorMessageDto errorMessage = new ErrorMessageDto(
                cell.getSheet().getSheetName(),
                (int) cellRelatedData[0],
                (int) cellRelatedData[1],
                (String) cellRelatedData[2],
                "El campo tiene una longitud superior a la permitida (" + lenght + ")"
        );

        if (-1 == HSSFCellErrorMessages.indexOf(errorMessage)) {
            HSSFCellErrorMessages.add(errorMessage);
        }

    }

    /**
     * Agrega un error si el valor de la celda no es de tipo numérico.
     *
     * @param cell Celda
     */
    private void addErrorToNotNumericCellValue(HSSFCell cell) {
        Object[] cellRelatedData = cellRelatedData(cell);
        ErrorMessageDto errorMessage = new ErrorMessageDto(
                cell.getSheet().getSheetName(),
                (int) cellRelatedData[0],
                (int) cellRelatedData[1],
                (String) cellRelatedData[2],
                "El campo no es de tipo numérico"
        );

        if (-1 == HSSFCellErrorMessages.indexOf(errorMessage)) {
            HSSFCellErrorMessages.add(errorMessage);
        }
    }

    private void addErrorToNotMatchDateRegexCellValue(XSSFCell cell, String regex) {
        Object[] cellRelatedData = cellRelatedData(cell);
        ErrorMessageDto errorMessage = new ErrorMessageDto(
                cell.getSheet().getSheetName(),
                (int) cellRelatedData[0],
                (int) cellRelatedData[1],
                (String) cellRelatedData[2],
                "La fecha no cuenta con el formato correcto (" + regex + ")"
        );

        if (-1 == XSSFCellErrorMessages.indexOf(errorMessage)) {
            XSSFCellErrorMessages.add(errorMessage);
        }
    }

    private void addErrorToNotMatchDateRegexCellValue(HSSFCell cell, String regex) {
        Object[] cellRelatedData = cellRelatedData(cell);
        ErrorMessageDto errorMessage = new ErrorMessageDto(
                cell.getSheet().getSheetName(),
                (int) cellRelatedData[0],
                (int) cellRelatedData[1],
                (String) cellRelatedData[2],
                "La fecha no cuenta con el formato correcto (" + regex + ")"
        );

        if (-1 == HSSFCellErrorMessages.indexOf(errorMessage)) {
            HSSFCellErrorMessages.add(errorMessage);
        }
    }

    /**
     * Evalua el contenido de la celda para verificar si está en blanco.
     *
     * @param cell Celda
     * @return <code>true</code> si el valor de la celda está en blanco.
     * <code>false</code> de otro modo.
     */
    public boolean isBlank(HSSFCell cell) {
        return ((cell == null) || (cell.getCellType() == HSSFCell.CELL_TYPE_BLANK) || (cell.toString().trim().equals("")));
    }

    /**
     * Convierte el valor de la celda en número decimal
     *
     * @param cell Celda
     * @param isMandatory si el valor es true, se validará que el valor de la
     * celda exista forzosamente. En caso contrario, se agrega un error a la
     * lista de errores.
     * @return <code>Double</code> si el valor concuerda con los criterios ||
     * <code>null</code> de otro modo
     */
    public Double toDecimalValue(XSSFCell cell, boolean isMandatory) {
        String valueContainer;
        DecimalFormat decimalFormat = formatNumericCellValue(cell);
        // es valor mandatory
        if (isMandatory) {
            // es mandatory y no es nulo
            if (isNotBlankNumericCell(cell) && null != decimalFormat) {
                double cellDoubleValue = Double.parseDouble(cell.toString().trim());
                valueContainer = decimalFormat.format(cellDoubleValue);

                return Double.parseDouble(valueContainer);
            } // si el campo no es numérico
            else if (null != cell) {
                addErrorToNotNumericCellValue(cell);
            } // si el campo es mandatory y es nulo
            else {
                addErrorToMandatoryCellValue(cell);
            }
        } // es int no mandatory y tiene valor
        else if (isNotBlankNumericCell(cell) && null != decimalFormat) {
            double cellDoubleValue = Double.parseDouble(cell.toString().trim());
            valueContainer = decimalFormat.format(cellDoubleValue);

            return Double.parseDouble(valueContainer);
        } else {
            return null;
        }

        return null;
    }

    /**
     * Convierte el valor de la celda en número entero
     *
     * @param cell Celda
     * @param isMandatory si el valor es true, se validará que el valor de la
     * celda exista forzosamente. En caso contrario, se agrega un error a la
     * lista de errores.
     * @return <code>Double</code> si el valor concuerda con los criterios ||
     * <code>null</code> de otro modo.
     */
    public Long toIntegerValue(XSSFCell cell, boolean isMandatory) {
        String valueContainer;
        DecimalFormat decimalFormat = new DecimalFormat("#################");

        // es int mandatory
        if (isMandatory) {
            // es int mandatory y no es nulo
            if (!isBlank(cell)) {
                try {
                    double cellDoubleValue = Double.parseDouble(cell.toString().trim());
                    valueContainer = decimalFormat.format(cellDoubleValue);

                    return Long.parseLong(valueContainer);
                } catch (Exception e) {
                    addErrorToNotNumericCellValue(cell);
                    return null;
                }
            } // si el campo no es numérico
            else if (null != cell) {
                addErrorToNotNumericCellValue(cell);
            } // si el campo es mandatory y es nulo
            else {
                addErrorToMandatoryCellValue(cell);
            }
            // es int no mandatory y tiene valor
        } else if (!isBlank(cell)) {
            try {
                double cellDoubleValue = Double.parseDouble(cell.toString().trim());
                valueContainer = decimalFormat.format(cellDoubleValue);

                return Long.parseLong(valueContainer);
            } catch (Exception e) {
                addErrorToNotNumericCellValue(cell);
                return null;
            }
        }
        return null;
    }

    /**
     * Evalua que la celda no esté vacía y que su contenido sea numérico
     *
     * @param cell Celda
     * @return <code>true</code> si la celda no está vacía y su valor es
     * numérico. <code>false</code> de otro modo.
     */
    public boolean isNotBlankNumericCell(XSSFCell cell) {
        return !isBlank(cell) && cell.getCellType() == HSSFCell.CELL_TYPE_NUMERIC;
    }

    /**
     * Crea una instancia <code>DecimalFormat</code> para evaluar el contenido
     * de una celda con valores numéricos.
     *
     * @param cell Celda
     * @return <code>DecimalFormat</code> si el contenido de la celda no está
     * vacío y es de tipo numérico. <code>null</code> de otro modo.
     */
    public DecimalFormat formatNumericCellValue(XSSFCell cell) {
        if (isNotBlankNumericCell(cell)) {

            if (cell.toString().contains(".")) {
                return new DecimalFormat("##########.##");
            } else {
                return new DecimalFormat("##########");
            }
        }

        return null;
    }

    /**
     * Devuelve el valor de la celda en formato texto.
     *
     * @param cell Celda.
     * @return <code>String</code> Valor de la celda.
     */
    public String toTextValue(XSSFCell cell) {
        return cell.toString();
    }

    /**
     * Devuelve el valor de la celda en formato texto evaluando la longitud y si
     * éste es obligatorio.
     *
     * @param cell Celda
     * @param lenght Longitud máxima del valor de la celda
     * @param isMandatory si el valor es true, se validará que el valor de la
     * celda exista forzosamente. En caso contrario, se agrega un error a la
     * lista de errores.
     * @return <code>String</code> Valor de la celda.
     */
    public String toTextValue(XSSFCell cell, int lenght, boolean isMandatory) {
        // el campo es obligado
        if (isMandatory) {
            // es campo obligado y no es nulo
            if (!isBlank(cell) && cell.toString().length() <= lenght) {
                return cell.toString();
            } else {
                addErrorToMandatoryCellValue(cell);
            }
            // es campo no obligado pero tiene valor
        } else if (!isBlank(cell)) {
            return cell.toString();
        }
        return null;
    }

    /**
     * Devuelve el valor de la celda en formato texto evaluando la longitud, la
     * expresión regular y si éste es obligatorio.
     *
     * @param cell Celda
     * @param lenght Longitud máxima del valor de la celda
     * @param regex Expresión regular que debe cumplir el valor de la celda
     * @param isMandatory si el valor es true, se validará que el valor de la
     * celda exista forzosamente. En caso contrario, se agrega un error a la
     * lista de errores.
     *
     * @return <code>String</code> Valor de la celda || <code>null</code> si el
     * campo no cumple con las validaciones
     */
    public String toTextValue(XSSFCell cell, int lenght, String regex, boolean isMandatory) {
        // el campo es obligado
        if (isMandatory) {
            // el campo obligado y no es nulo
            if (!isBlank(cell)) {
                // el campo es obligado, no es nulo y tiene longitud aceptada
                if (cell.toString().length() <= lenght) {
                    // el campo es obligado, no es nulo, tiene longitud aceptada y cumple con la expresión regular
                    if (cell.toString().matches(regex)) {
                        return cell.toString();
                    } else {
                        addErrorToRegexNotMatchCellValue(cell);
                    }
                } else {
                    addErrorToOutOfLenghtCellValue(cell, lenght);
                }
            } else {
                addErrorToMandatoryCellValue(cell);
            }
            // es campo no obligado pero tiene valor
        } else if (!isBlank(cell)) {
            // el campo no es obligado, no es nulo y tiene longitud aceptada
            if (cell.toString().length() <= lenght) {
                // el campo no es obligado, no es nulo, tiene longitud aceptada y cumple con la expresión regular
                if (cell.toString().matches(regex)) {
                    return cell.toString();
                } else {
                    addErrorToRegexNotMatchCellValue(cell);
                }
            } else {
                addErrorToOutOfLenghtCellValue(cell, lenght);
            }
        }

        return null;
    }

    /**
     * Agrega un error si el valor es obligatorio
     *
     * @param cell Celda
     */
    private void addErrorToMandatoryCellValue(XSSFCell cell) {
        Object[] cellRelatedData = cellRelatedData(cell);
        ErrorMessageDto errorMessage = new ErrorMessageDto(
                cell.getSheet().getSheetName(),
                (int) cellRelatedData[0],
                (int) cellRelatedData[1],
                (String) cellRelatedData[2],
                "El campo no pueda estar vacío"
        );

        if (-1 == XSSFCellErrorMessages.indexOf(errorMessage)) {
            XSSFCellErrorMessages.add(errorMessage);
        }

    }

    /**
     * Agrega un error si el valor no concuerda con una expresión regular
     *
     * @param cell
     */
    private void addErrorToRegexNotMatchCellValue(XSSFCell cell) {
        Object[] cellRelatedData = cellRelatedData(cell);
        ErrorMessageDto errorMessage = new ErrorMessageDto(
                cell.getSheet().getSheetName(),
                (int) cellRelatedData[0],
                (int) cellRelatedData[1],
                (String) cellRelatedData[2],
                "El campo no concuerda con el formato"
        );

        if (-1 == XSSFCellErrorMessages.indexOf(errorMessage)) {
            XSSFCellErrorMessages.add(errorMessage);
        }

    }

    /**
     * Agrega un error si la longitud del valor de la celda es mayor a la
     * longitud máxima indicada.
     *
     * @param cell Celda
     * @param lenght Longitud máxima
     */
    private void addErrorToOutOfLenghtCellValue(XSSFCell cell, int lenght) {
        Object[] cellRelatedData = cellRelatedData(cell);
        ErrorMessageDto errorMessage = new ErrorMessageDto(
                cell.getSheet().getSheetName(),
                (int) cellRelatedData[0],
                (int) cellRelatedData[1],
                (String) cellRelatedData[2],
                "El campo tiene una longitud superior a la permitida (" + lenght + ")"
        );

        if (-1 == XSSFCellErrorMessages.indexOf(errorMessage)) {
            XSSFCellErrorMessages.add(errorMessage);
        }

    }

    /**
     * Agrega un error si el valor de la celda no es de tipo numérico.
     *
     * @param cell Celda
     */
    private void addErrorToNotNumericCellValue(XSSFCell cell) {
        Object[] cellRelatedData = cellRelatedData(cell);
        ErrorMessageDto errorMessage = new ErrorMessageDto(
                cell.getSheet().getSheetName(),
                (int) cellRelatedData[0],
                (int) cellRelatedData[1],
                (String) cellRelatedData[2],
                "El campo no es de tipo numérico"
        );

        if (-1 == XSSFCellErrorMessages.indexOf(errorMessage)) {
            XSSFCellErrorMessages.add(errorMessage);
        }
    }

    /**
     * Evalua el contenido de la celda para verificar si está en blanco.
     *
     * @param cell Celda
     * @return <code>true</code> si el valor de la celda está en blanco.
     * <code>false</code> de otro modo.
     */
    public boolean isBlank(XSSFCell cell) {
        return ((cell == null) || (cell.getCellType() == HSSFCell.CELL_TYPE_BLANK) || (cell.toString().trim().equals("")));
    }

    /**
     * Devuelve un arreglo de <code>Object</code> con la información de la
     * celda.
     *
     * <p>
     * cellRelatedData[0] = número de celda (<code>int</code>)</p>
     * <p>
     * cellRelatedData[1] = número de fila (<code>int</code>)</p>
     * <p>
     * cellRelatedData[2] = Cabecera de la columna (<code>String</code>)</p>
     *
     * @param cell Celda
     * @return <code>Object[]</code>
     */
    public Object[] cellRelatedData(HSSFCell cell) {
        Object[] cellRelatedData = new Object[3];

        int cellNumber = (int) cell.getCellNum() + 1;
        int rowNumber = cell.getRow().getRowNum() + 1;
        String column = cell.getRow().getSheet().getRow(0).getCell(cell.getCellNum() + 1).toString();
        cellRelatedData[0] = cellNumber;
        cellRelatedData[1] = rowNumber;
        cellRelatedData[2] = column;

        return cellRelatedData;
    }

    /**
     * Devuelve un arreglo de <code>Object</code> con la información de la
     * celda.
     * <p>
     * cellRelatedData[0] = número de celda (<code>int</code>)</p>
     * <p>
     * cellRelatedData[1] = número de fila (<code>int</code>)</p>
     * <p>
     * cellRelatedData[2] = Cabecera de la columna (<code>String</code>)</p>
     *
     * @param cell Celda
     * @return <code>Object[]</code>
     */
    public Object[] cellRelatedData(XSSFCell cell) {
        CellReference.convertNumToColString(cell.getColumnIndex());
        Object[] cellRelatedData = new Object[3];
        int cellNumber = (int) cell.getColumnIndex() + 1;
        int rowNumber = cell.getRow().getRowNum() + 1;
        String column = CellReference.convertNumToColString(cell.getColumnIndex());
        cellRelatedData[0] = cellNumber;
        cellRelatedData[1] = rowNumber;
        cellRelatedData[2] = column;

        return cellRelatedData;
    }

    /**
     * Retorna la lista de errores de tipo <code>HSSFCell</code>
     *
     * @return <code>List</code> Lista de errores
     */
    public List<ErrorMessageDto> getHSSFCellErrorMessages() {
        return HSSFCellErrorMessages;
    }

    /**
     * Retorna la lista de errores.
     *
     * @return <code>List</code> Lista de errores de tipo <code>XSSFCel</code>
     */
    public List<ErrorMessageDto> getXSSFCellErrorMessages() {
        return XSSFCellErrorMessages;
    }
}
