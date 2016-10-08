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

import jxl.CellView;
import jxl.Range;
import jxl.Sheet;
import jxl.format.CellFormat;
import jxl.format.PageOrientation;
import jxl.format.PaperSize;
import jxl.write.biff.RowsExceededException;

/**
 * Interface for a worksheet that may be modified.  The most
 * important modification for a sheet is to have cells added to it
 */
public interface WritableSheet extends Sheet
{
  /**
   * Adds a cell to this sheet
   * The RowsExceededException may be caught if client code wishes to
   * explicitly trap the case where too many rows have been written
   * to the current sheet.  If this behaviour is not desired, it is
   * sufficient simply to handle the WriteException, since this is a base
   * class of RowsExceededException
   *
   * @param cell the cell to add
   * @exception jxl.write..WriteException
   * @exception jxl.write.biff.RowsExceededException
   */
  public void addCell(WritableCell cell)
    throws WriteException, RowsExceededException;
  /**
   * Sets the name of this sheet
   *
   * @param name the name of the sheet
   */
  public void setName(String name);
  /**
   * Indicates whether or not this sheet is hidden
   *
   * @param hidden hidden flag
   * @deprecated use the SheetSettings bean instead
   */
  public void setHidden(boolean hidden);
  /**
   * Indicates whether or not this sheet is protected
   *
   * @param prot Protected flag
   * @deprecated use the SheetSettings bean instead
   */
  public void setProtected(boolean prot);

  /**
   * Sets the width of the column on this sheet, in characters.  This causes
   * Excel to resize the entire column.
   * If the columns specified already has view information associated
   * with it, then it is replaced by the new data
   *
   * @param col the column to be formatted
   * @param width the width of the column
   */
  public void setColumnView(int col, int width);

  /**
   * Sets the width and style of every cell in the specified column.
   * If the columns specified already has view information associated
   * with it, then it is replaced by the new data
   *
   * @param col the column to be formatted
   * @param format the format of every cell in the column
   * @param width the width of the column, in characters
   * @deprecated Use the CellView bean instead
   */
  public void setColumnView(int col, int width, CellFormat format);

  /**
   * Sets the view for this column
   *
   * @param col the column on which to set the view
   * @param view the view to set
   */
  public void setColumnView(int col, CellView view);

  /**
   * Sets the height of the specified row, as well as its collapse status
   *
   * @param row the row to be formatted
   * @param height the row height in characters
   * @exception jxl.write.biff.RowsExceededException
   */
  public void setRowView(int row, int height)
    throws RowsExceededException;

  /**
   * Sets the properties of the specified row
   *
   * @param row the row to be formatted
   * @param collapsed indicates whether the row is collapsed
   * @exception jxl.write.biff.RowsExceededException
   */
  public void setRowView(int row, boolean collapsed)
    throws RowsExceededException;

  /**
   * Sets the height of the specified row, as well as its collapse status
   *
   * @param row the row to be formatted
   * @param height the row height in 1/20th of a point
   * @param collapsed indicates whether the row is collapsed
   * @exception jxl.write.biff.RowsExceededException
   */
  public void setRowView(int row, int height,
                         boolean collapsed)
                         throws RowsExceededException;

  /**
   * Sets the view for this column
   *
   * @param row the column on which to set the view
   * @param view the view to set
   * @exception RowsExceededException
   */
  public void setRowView(int row, CellView view) throws RowsExceededException;

  /**
   * Gets the writable cell from this sheet.  Use of this method allows
   * the returned  cell to be modified by the users application
   *
   * @param column the column
   * @param row the row
   * @return the cell at the specified position
   */
  public WritableCell getWritableCell(int column, int row);

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
  public WritableCell getWritableCell(String loc);

  /**
   * Gets the writable hyperlinks from this sheet.  The hyperlinks
   * that are returned may be modified by user applications
   *
   * @return the writable hyperlinks
   */
  public WritableHyperlink[] getWritableHyperlinks();

  /**
   * Inserts a blank row into this spreadsheet.  If the row is out of range
   * of the rows in the sheet, then no action is taken
   *
   * @param row the row to insert
   */
  public void insertRow(int row);

  /**
   * Inserts a blank column into this spreadsheet.  If the column is out of
   * range of the columns in the sheet, then no action is taken
   *
   * @param col the column to insert
   */
  public void insertColumn(int col);

  /**
   * Removes a column from this spreadsheet.  If the column is out of range
   * of the columns in the sheet, then no action is taken
   *
   * @param col the column to remove
   */
  public void removeColumn(int col);

  /**
   * Removes a row from this spreadsheet.  If the row is out of
   * range of the columns in the sheet, then no action is taken
   *
   * @param row the row to remove
   */
  public void removeRow(int row);

  /**
   * Merges the specified cells.  Any clashes or intersections between
   * merged cells are resolved when the spreadsheet is written out
   *
   * @param col1 the column number of the top left cell
   * @param row1 the row number of the top left cell
   * @param col2 the column number of the bottom right cell
   * @param row2 the row number of the bottom right cell
   * @return the Range object representing the merged cells
   * @exception jxl.write..WriteException
   * @exception jxl.write.biff.RowsExceededException
   */
  public Range mergeCells(int col1, int row1, int col2, int row2)
    throws WriteException, RowsExceededException;

  /** 
   * Sets a row grouping
   *
   * @param row1 the first row of the group
   * @param row2 the last row of the group
   * @param collapsed should the group be collapsed?
   * @exception WriteException
   * @exception RowsExceededException
   */
  public void setRowGroup(int row1, int row2, boolean collapsed)
    throws WriteException, RowsExceededException;

  /** 
   * Unsets a row grouping
   *
   * @param row1 the first row to unset
   * @param row2 the last row to unset
   * @exception WriteException
   * @exception RowsExceededException
   */
  public void unsetRowGroup(int row1, int row2)
    throws WriteException, RowsExceededException;

  /** 
   * Sets a column grouping
   *
   * @param col1 the first column of the group
   * @param col2 the last column of the group
   * @param collapsed should the group be collapsed?
   * @exception WriteException
   * @exception RowsExceededException
   */
  public void setColumnGroup(int col1, int col2, boolean collapsed)
    throws WriteException, RowsExceededException;

  /** 
   * Unsets a column grouping
   *
   * @param col1 the first column to unset
   * @param col2 the last column to unset
   * @exception WriteException
   * @exception RowsExceededException
   */
  public void unsetColumnGroup(int col1, int col2)
    throws WriteException, RowsExceededException;
    
  /**
   * Unmerges the specified cells.  The Range passed in should be one that
   * has been previously returned as a result of the getMergedCells method
   *
   * @param r the range of cells to unmerge
   */
  public void unmergeCells(Range r);

  /**
   * Adds the specified hyperlink.  Adding a hyperlink causes any populated
   * cells in the range of the hyperlink to be set to empty
   * If the cells which activate this hyperlink clash with any other cells,
   * they are still added to the worksheet and it is left to Excel to
   * handle this.
   *
   * @param h the hyperlink
   * @exception jxl.write..WriteException
   * @exception jxl.write.biff.RowsExceededException
   */
  public void addHyperlink(WritableHyperlink h)
    throws WriteException, RowsExceededException;;

  /**
   * Removes the specified hyperlink.  Note that if you merely set the
   * cell contents to be an Empty cell, then the cells containing the
   * hyperlink will still be active.  The contents of the cell which
   * activate the hyperlink are removed.
   * The hyperlink passed in must be a hyperlink retrieved using the
   * getHyperlinks method
   *
   * @param h the hyperlink to remove.
   */
  public void removeHyperlink(WritableHyperlink h);

  /**
   * Removes the specified hyperlink.  Note that if you merely set the
   * cell contents to be an Empty cell, then the cells containing the
   * hyperlink will still be active.
   * If the preserveLabel field is set, the cell contents of the
   * hyperlink are preserved, although the hyperlink is deactivated.  If
   * this value is FALSE, the cell contents are removed
   * The hyperlink passed in must be a hyperlink retrieved using the
   * getHyperlinks method
   *
   * @param h the hyperlink to remove.
   * @param preserveLabel if TRUE preserves the label contents, if FALSE
   * removes them
   */
  public void removeHyperlink(WritableHyperlink h, boolean preserveLabel);

  /**
   * Sets the header for this page
   *
   * @param l the print header to print on the left side
   * @param c the print header to print in the centre
   * @param r the print header to print on the right hand side
   * @deprecated use  the SheetSettings bean
   */
  public void setHeader(String l, String c, String r);

  /**
   * Sets the footer for this page
   *
   * @param l the print header to print on the left side
   * @param c the print header to print in the centre
   * @param r the print header to print on the right hand side
   * @deprecated use the SheetSettings bean
   */
  public void setFooter(String l, String c, String r);

  /**
   * Sets the page setup details
   *
   * @param p  the page orientation
   */
  public void setPageSetup(PageOrientation p);

  /**
   * Sets the page setup details
   *
   * @param p  the page orientation
   * @param hm the header margin, in inches
   * @param fm the footer margin, in inches
   */
  public void setPageSetup(PageOrientation p, double hm, double fm);

  /**
   * Sets the page setup details
   *
   * @param p  the page orientation
   * @param ps the paper size
   * @param hm the header margin, in inches
   * @param fm the footer margin, in inches
   */
  public void setPageSetup(PageOrientation p, PaperSize ps,
                           double hm, double fm);

  /**
   * Forces a page break at the specified row
   *
   * @param row the row to break at
   */
  public void addRowPageBreak(int row);
  
  /**
   * Forces a page break at the specified column
   *
   * @param col the column to break at
   */
  public void addColumnPageBreak(int col);

  /**
   * Adds an image to the sheet
   *
   * @param image the image to add
   */
  public void addImage(WritableImage image);

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
  public WritableImage getImage(int i);

  /**
   * Removes the specified image from the sheet.  The image passed in
   * must be the same instance as that previously retrieved using the
   * getImage() method
   *
   * @param wi the image to remove
   */
  public void removeImage(WritableImage wi);

  /**
   * Extend the data validation contained in the specified cell across and 
   * downwards.
   * NOTE:  The source cell (top left) must have been added to the sheet prior
   * to this method being called
   *
   * @param col the number of cells accross to apply this data validation
   * @param row the number of cells downwards to apply this data validation
   */
  public void applySharedDataValidation(WritableCell cell, int col, int row) 
    throws WriteException;

  /**
   * Remove the shared data validation from multiple cells.  The cell passed 
   * in is the top left cell.  The data validation is removed from this 
   * cell and all cells which share the same validation.
   *
   * @param cell the top left cell containing the shared data validation
   */
  public void removeSharedDataValidation(WritableCell cell)
    throws WriteException;
}


