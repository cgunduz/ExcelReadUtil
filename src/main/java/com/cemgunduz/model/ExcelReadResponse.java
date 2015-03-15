package com.cemgunduz.model;

import java.util.ArrayList;
import java.util.List;

/**
 * Created by cgunduz on 4/8/14.
 */
public class ExcelReadResponse<T> {

    public ExcelReadResponse()
    {
        excelReadErrorList = new ArrayList<ExcelReadError>();
        emptySheets = new ArrayList<Integer>();
        successful = true;
    }

    private List<ExcelSheet> excelSheetList;
    private List<T> entityList;

    private boolean successful;
    private List<ExcelReadError> excelReadErrorList;

    private List<Integer> emptySheets;

    public List<Integer> getEmptySheets() {
        return emptySheets;
    }

    public boolean isSuccessful() {
        return successful;
    }

    public List<ExcelSheet> getExcelSheetList() {
        return excelSheetList;
    }

    public void setExcelSheetList(List<ExcelSheet> excelSheetList) {
        this.excelSheetList = excelSheetList;
    }

    public List<T> getEntityList() {
        return entityList;
    }

    public void setEntityList(List<T> entityList) {
        this.entityList = entityList;
    }

    public void setSuccessful(boolean successful) {
        this.successful = successful;
    }

    public List<ExcelReadError> getExcelReadErrorList() {
        return excelReadErrorList;
    }

    public void setExcelReadErrorList(List<ExcelReadError> excelReadErrorList) {
        this.excelReadErrorList = excelReadErrorList;
    }
}
