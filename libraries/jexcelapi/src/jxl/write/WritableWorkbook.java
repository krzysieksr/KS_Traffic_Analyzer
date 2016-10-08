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

package jxl.write;

import java.io.IOException;

import jxl.Range;
import jxl.Sheet;
import jxl.Workbook;
import jxl.format.Colour;
import jxl.format.UnderlineStyle;

/**
 * A writable workbook
 */
public abstract class WritableWorkbook
{
  // Globally available stuff

  /**
   * The default font for Cell formats
   */
  public static final WritableFont  ARIAL_10_PT =
    new WritableFont(WritableFont.ARIAL);

  /**
   * The font used for hyperlinks
   */
  public static final WritableFont HYPERLINK_FONT =
    new WritableFont(WritableFont.ARIAL,
                     WritableFont.DEFAULT_POINT_SIZE,
                     WritableFont.NO_BOLD,
                     false,
                     UnderlineStyle.SINGLE,
                     Colour.BLUE);

  /**
   * The default style for cells
   */
  public static final WritableCellFormat NORMAL_STYLE =
    new WritableCellFormat(ARIAL_10_PT, NumberFormats.DEFAULT);

  /**
   * The style used for hyperlinks
   */
  public static final WritableCellFormat HYPERLINK_STYLE =
    new WritableCellFormat(HYPERLINK_FONT);

  /**
   * A cell format used to hide the cell contents
   */
  public static final WritableCellFormat HIDDEN_STYLE = 
    new WritableCellFormat(new DateFormat(";;;"));

  /**
   * Constructor used by the implemenation class
   */
  protected WritableWorkbook()
  {
  }

  /**
   * Gets the sheets within this workbook.  Use of this method for
   * large worksheets can cause performance problems.
   *
   * @return an array of the individual sheets
   */
  public abstract WritableSheet[] getSheets();

  /**
   * Gets the sheet names
   *
   * @return an array of strings containing the sheet names
   */
  public abstract String[] getSheetNames();

  /**
   * Gets the specified sheet within this workbook
   *
   * @param index the zero based index of the reQuired sheet
   * @return The sheet specified by the index
   * @exception IndexOutOfBoundsException when index refers to a non-existent
   *            sheet
   */
  public abstract WritableSheet getSheet(int index)
    throws IndexOutOfBoundsException;

  /**
   * Gets the sheet with the specified name from within this workbook
   *
   * @param name the sheet name
   * @return The sheet with the specified name, or null if it is not found
   */
  public abstract WritableSheet getSheet(String name);

  /**
   * Returns the cell for the specified location eg. "Sheet1!A4".
   * This is identical to using the CellReferenceHelper with its
   * associated performance overheads, consequently it should
   * be use sparingly
   *
   * @param loc the cell to retrieve
   * @return the cell at the specified location
   */
  public abstract WritableCell getWritableCell(String loc);

  /**
   * Returns the number of sheets in this workbook
   *
   * @return the number of sheets in this workbook
   */
  public abstract int getNumberOfSheets();

  /**
   * Closes this workbook, and makes any memory allocated available
   * for garbage collection.  Also closes the underlying output stream
   * if necessary.
   *
   * @exception IOException
   * @exception WriteException
   */
  public abstract void close() throws IOException, WriteException;

  /**
   * Creates, and returns a worksheet at the specified position
   * with the specified name
   * If the index specified is less than or equal to zero, the new sheet
   * is created at the beginning of the workbook.  If the index is greater
   * than the number of sheet, then the sheet is created at the
   * end of the workbook.
   *
   * @param name the sheet name
   * @param index the index number at which to insert
   * @return the new sheet
   */
  public abstract WritableSheet createSheet(String name, int index);

  /**
   * Imports a sheet from a different workbook.  Does a deep copy on all
   * elements within that sheet
   *
   * @param name the name of the new sheet
   * @param index the position for the new sheet within this workbook
   * @param sheet the sheet (from another workbook) to merge into this one
   * @return the new sheet
   */
  public abstract WritableSheet importSheet(String name, int index, Sheet s);

  /**
   * Copy sheet within the same workbook.  The sheet specified is copied to
   * the new sheet name at the position
   *
   * @param s the index of the sheet to copy
   * @param name the name of the new sheet
   * @param index the position of the new sheet
   */
  public abstract void copySheet(int s, String name, int index);

  /**
   * Copies the specified sheet and places it at the index
   * specified by the parameter
   *
   * @param s the name of the sheet to copy
   * @param name the name of the new sheet
   * @param index the position of the new sheet
   */
  public abstract void copySheet(String s, String name, int index);

  /**
   * Removes the sheet at the specified index from this workbook
   *
   * @param index the sheet index to remove
   */
  public abstract void removeSheet(int index);

  /**
   * Moves the specified sheet within this workbook to another index
   * position.
   *
   * @param fromIndex the zero based index of the required sheet
   * @param toIndex the zero based index of the required sheet
   * @return the sheet that has been moved
   */
  public abstract WritableSheet moveSheet(int fromIndex, int toIndex);

  /**
   * Writes out the data held in this workbook in Excel format
   *
   * @exception IOException
   */
  public abstract void write() throws IOException;

  /**
   * Indicates whether or not this workbook is protected
   *
   * @param prot Protected flag
   */
  public abstract void setProtected(boolean prot);

  /**
   * Sets the RGB value for the specified colour for this workbook
   *
   * @param c the colour whose RGB value is to be overwritten
   * @param r the red portion to set (0-255)
   * @param g the green portion to set (0-255)
   * @param b the blue portion to set (0-255)
   */
  public abstract void setColourRGB(Colour c, int r, int g, int b);

  /**
   * This method can be used to create a writable clone of some other
   * workbook
   *
   * @param w the workdoock to copy
   * @deprecated Copying now occurs implicitly as part of the overloaded
   *   factory method Workbook.createWorkbood
   */
  public void copy(Workbook w)
  {
    // Was an abstract method - leave the method body blank
  }

  /**
   * Gets the named cell from this workbook.  The name refers to a
   * range of cells, then the cell on the top left is returned.  If
   * the name cannot be, null is returned
   *
   * @param name the name of the cell/range to search for
   * @return the cell in the top left of the range if found, NULL
   *         otherwise
   */
  public abstract WritableCell findCellByName(String name);

  /**
   * Gets the named range from this workbook.  The Range object returns
   * contains all the cells from the top left to the bottom right
   * of the range.
   * If the named range comprises an adjacent range,
   * the Range[] will contain one object; for non-adjacent
   * ranges, it is necessary to return an array of length greater than
   * one.
   * If the named range contains a single cell, the top left and
   * bottom right cell will be the same cell
   *
   * @param name the name of the cell/range to search for
   * @return the range of cells
   */
  public abstract Range[] findByName(String name);

  /**
   * Gets the named ranges
   *
   * @return the list of named cells within the workbook
   */
  public abstract String[] getRangeNames();

  /**
   * Removes the specified named range from the workbook.  Note that 
   * removing a name could cause formulas which use that name to
   * calculate their results incorrectly
   *
   * @param name the name to remove
   */
  public abstract void removeRangeName(String name);

  /**
   * Add new named area to this workbook with the given information.
   *
   * @param name name to be created.
   * @param sheet sheet containing the name
   * @param firstCol  first column this name refers to.
   * @param firstRow  first row this name refers to.
   * @param lastCol    last column this name refers to.
   * @param lastRow    last row this name refers to.
   */
  public abstract void addNameArea(String name,
                                   WritableSheet sheet,
                                   int firstCol,
                                   int firstRow,
                                   int lastCol,
                                   int lastRow);

  /**
   * Sets a new output file.  This allows the same workbook to be
   * written to various different output files without having to
   * read in any templates again
   *
   * @param fileName the file name
   * @exception IOException
   */
  public abstract void setOutputFile(java.io.File fileName)
    throws IOException;
}
