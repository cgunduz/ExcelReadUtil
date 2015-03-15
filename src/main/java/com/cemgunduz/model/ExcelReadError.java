package com.cemgunduz.model;

/**
 * Created by cgunduz on 4/7/14.
 */
public class ExcelReadError {

    private ExcelReadErrorType excelReadErrorType;
    private Integer rowNumber;
    private Integer coloumnNumber;

    public ExcelReadError(){}

    public ExcelReadError(ExcelReadErrorType excelReadErrorType)
    {
        this.excelReadErrorType = excelReadErrorType;
    }

    public ExcelReadError(ExcelReadErrorType excelReadErrorType, Integer rowNumber)
    {
        this.excelReadErrorType = excelReadErrorType;
        this.rowNumber = rowNumber;
    }

    public ExcelReadError(ExcelReadErrorType excelReadErrorType, Integer rowNumber, Integer coloumnNumber)
    {
        this.excelReadErrorType = excelReadErrorType;
        this.rowNumber = rowNumber;
        this.coloumnNumber = coloumnNumber;
    }

    public ExcelReadErrorType getExcelReadErrorMessage() {
        return excelReadErrorType;
    }

    public void setExcelReadErrorMessage(ExcelReadErrorType excelReadErrorType) {
        this.excelReadErrorType = excelReadErrorType;
    }

    public Integer getRowNumber() {
        return rowNumber;
    }

    public void setRowNumber(Integer rowNumber) {
        this.rowNumber = rowNumber;
    }

    public Integer getColoumnNumber() {
        return coloumnNumber;
    }

    public void setColoumnNumber(Integer coloumnNumber) {
        this.coloumnNumber = coloumnNumber;
    }
}
