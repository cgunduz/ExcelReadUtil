package com.cemgunduz.model;

import org.apache.poi.ss.usermodel.Cell;

import java.util.ArrayList;
import java.util.List;

/**
 * Created by cgunduz on 3/12/14.
 */
public class ExcelSheet {

    private List<List<Cell>> cellMatrice;

    public ExcelSheet()
    {
        cellMatrice = new ArrayList<List<Cell>>();
    }

    /**
     * Direct referance to cell matrice is provided but advised aganist
     *
     * @return
     */
    public List<List<Cell>> getCellMatrice()
    {
        return cellMatrice;
    }

    /**
     * Getting a cell from an excel sheet by refering its two dimensional space
     *
     * @param rowNumber
     * @param cellNumber
     * @return
     */
    public Cell getCell(int rowNumber, int cellNumber)
    {
        return cellMatrice.get(rowNumber).get(cellNumber);
    }

    /**
     * Getting the string value of a cell from an excel sheet by refering its two dimensional space
     *
     * @param rowNumber
     * @param cellNumber
     * @return
     */
    public String getCellValue(int rowNumber, int cellNumber)
    {
        return getCell(rowNumber,cellNumber).toString();
    }

    public int getTotalRows()
    {
        return cellMatrice.size();
    }

    public int getTotalCellsInARow(int cellNo)
    {
        return cellMatrice.get(cellNo).size();
    }
}
