/*********************************************************************
*
*      Copyright (C) 2002 Andrew Khan
*
* This library is free software; you can redistribute it and/or
* modify it under the terms of the GNU Lesser General Public
* License as published by the Free Software Foundation; either
* version 2.1 of the License, or (at your option) any later version.
*
* This library is distributed in the hope that it will be useful,
* but WITHOUT ANY WARRANTY; without even the implied warranty of
* MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
* Lesser General Public License for more details.
*
* You should have received a copy of the GNU Lesser General Public
* License along with this library; if not, write to the Free Software
* Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA 02111-1307 USA
***************************************************************************/

package jxl;

import java.util.regex.Pattern;
import jxl.format.CellFormat;

/**
 * Represents a sheet within a workbook.  Provides a handle to the individual
 * cells, or lines of cells (grouped by Row or Column)
 */
public interface Sheet
{
  /**
   * Returns the cell specified at this row and at this column.
   * If a column/row combination forms part of a merged group of cells
   * then (unless it is the first cell of the group) a blank cell
   * will be returned
   *
   * @param column the column number
   * @param row the row number
   * @return the cell at the specified co-ordinates
   */
  public Cell getCell(int column, int row);

  /**
   * Returns the cell for the specified location eg. "A4".  Note that this
   * method is identical to calling getCell(CellReferenceHelper.getColumn(loc),
   * CellReferenceHelper.getRow(loc)) and its implicit performance
   * overhead for string parsing.  As such,this method should therefore
   * be used sparingly
   *
   * @param loc the cell reference
   * @return the cell at the specified co-ordinates
   */
  public Cell getCell(String loc);

  /**
   * Returns the number of rows in this sheet
   *
   * @return the number of rows in this sheet
   */
  public int getRows();

  /**
   * Returns the number of columns in this sheet
   *
   * @return the number of columns in this sheet
   */
  public int getColumns();

  /**
   * Gets all the cells on the specified row
   *
   * @param row the rows whose cells are to be returned
   * @return the cells on the given row
   */
  public Cell[] getRow(int row);

  /**
   * Gets all the cells on the specified column
   *
   * @param col the column whose cells are to be returned
   * @return the cells on the specified column
   */
  public Cell[] getColumn(int col);

  /**
   * Gets the name of this sheet
   *
   * @return the name of the sheet
   */
  public String getName();

  /**
   * Determines whether the sheet is hidden
   *
   * @return whether or not the sheet is hidden
   * @deprecated in favour of the getSettings() method
   */
  public boolean isHidden();

  /**
   * Determines whether the sheet is protected
   *
   * @return whether or not the sheet is protected
   * @deprecated in favour of the getSettings() method
   */
  public boolean isProtected();

  /**
   * Gets the cell whose contents match the string passed in.
   * If no match is found, then null is returned.  The search is performed
   * on a row by row basis, so the lower the row number, the more
   * efficiently the algorithm will perform
   *
   * @param  contents the string to match
   * @return the Cell whose contents match the paramter, null if not found
   */
  public Cell findCell(String contents);

  /**
   * Gets the cell whose contents match the string passed in.
   * If no match is found, then null is returned.  The search is performed
   * on a row by row basis, so the lower the row number, the more
   * efficiently the algorithm will perform
   * 
   * @param contents the string to match
   * @param firstCol the first column within the range
   * @param firstRow the first row of the range
   * @param lastCol the last column within the range
   * @param lastRow the last row within the range
   * @param reverse indicates whether to perform a reverse search or not
   * @return the Cell whose contents match the parameter, null if not found
   */
  public Cell findCell(String contents, 
                       int firstCol, 
                       int firstRow, 
                       int lastCol, 
                       int lastRow, 
                       boolean reverse);

  /**
   * Gets the cell whose contents match the regular expressionstring passed in.
   * If no match is found, then null is returned.  The search is performed
   * on a row by row basis, so the lower the row number, the more
   * efficiently the algorithm will perform
   * 
   * @param pattern the regular expression string to match
   * @param firstCol the first column within the range
   * @param firstRow the first row of the rang
   * @param lastCol the last column within the range
   * @param lastRow the last row within the range
   * @param reverse indicates whether to perform a reverse search or not
   * @return the Cell whose contents match the parameter, null if not found
   */
  public Cell findCell(Pattern pattern, 
                       int firstCol, 
                       int firstRow,
                       int lastCol,  
                       int lastRow, 
                       boolean reverse);

  /**
   * Gets the cell whose contents match the string passed in.
   * If no match is found, then null is returned.  The search is performed
   * on a row by row basis, so the lower the row number, the more
   * efficiently the algorithm will perform.  This method differs
   * from the findCell method in that only cells with labels are
   * queried - all numerical cells are ignored.  This should therefore
   * improve performance.
   *
   * @param  contents the string to match
   * @return the Cell whose contents match the paramter, null if not found
   */
  public LabelCell findLabelCell(String contents);

  /**
   * Gets the hyperlinks on this sheet
   *
   * @return an array of hyperlinks
   */
  public Hyperlink[] getHyperlinks();

  /**
   * Gets the cells which have been merged on this sheet
   *
   * @return an array of range objects
   */
  public Range[] getMergedCells();

  /**
   * Gets the settings used on a particular sheet
   *
   * @return the sheet settings
   */
  public SheetSettings getSettings();

  /**
   * Gets the column format for the specified column
   *
   * @param col the column number
   * @return the column format, or NULL if the column has no specific format
   * @deprecated Use getColumnView and the CellView bean instead
   */
  public CellFormat getColumnFormat(int col);

  /**
   * Gets the column width for the specified column
   *
   * @param col the column number
   * @return the column width, or the default width if the column has no
   *         specified format
   * @deprecated Use getColumnView instead
   */
  public int getColumnWidth(int col);

  /**
   * Gets the column width for the specified column
   *
   * @param col the column number
   * @return the column format, or the default format if no override is
             specified
   */
  public CellView getColumnView(int col);

  /**
   * Gets the row height for the specified column
   *
   * @param row the row number
   * @return the row height, or the default height if the column has no
   *         specified format
   * @deprecated use getRowView instead
   */
  public int getRowHeight(int row);

  /**
   * Gets the row height for the specified column
   *
   * @param row the row number
   * @return the row format, which may be the default format if no format
   *         is specified
   */
  public CellView getRowView(int row);

  /**
   * Accessor for the number of images on the sheet
   *
   * @return the number of images on this sheet
   */
  public int getNumberOfImages();

  /**
   * Accessor for the image
   *
   * @param i the 0 based image number
   * @return  the image at the specified position
   */
  public Image getDrawing(int i);

  /**
   * Accessor for the page breaks on this sheet
   *
   * @return the page breaks on this sheet
   */
  public int[] getRowPageBreaks();

  /**
   * Accessor for the page breaks on this sheet
   *
   * @return the page breaks on this sheet
   */
  public int[] getColumnPageBreaks();

}







