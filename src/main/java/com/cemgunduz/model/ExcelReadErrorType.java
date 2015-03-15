package com.cemgunduz.model;

/**
 * Created by cgunduz on 4/7/14.
 */
public enum ExcelReadErrorType {

    TOO_MANY_INPUTS("More inputs than expected at the row"),
    NOT_ENOUGH_INPUTS("Less inputs than expected at the row"),
    ILLEGAL_TYPE("Type mismatch at row"),
    IO_EXCEPTION("File can not be read"),
    CLASS_INSTANTIATION_EXCEPTION("No public default constructor"),
    INVALID_FORMAT_EXCEPTION("File corrupted or in an unsupported format"),
    EMPTY_SHEET("Desired sheet is empty"),
    UNEXPECTED("Unexpected exception thrown cause unknown");

    private String message;

    private ExcelReadErrorType(String message)
    {
        this.message = message;
    }

    public String getMessage()
    {
        return message;
    }
}
