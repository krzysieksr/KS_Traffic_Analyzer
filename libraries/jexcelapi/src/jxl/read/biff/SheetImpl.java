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

package jxl.read.biff;

import java.util.ArrayList;
import java.util.Iterator;
import java.util.regex.Pattern;

import jxl.common.Logger;
import jxl.common.Assert;

import jxl.Cell;
import jxl.CellType;
import jxl.CellView;
import jxl.Hyperlink;
import jxl.Image;
import jxl.LabelCell;
import jxl.Range;
import jxl.Sheet;
import jxl.SheetSettings;
import jxl.WorkbookSettings;
import jxl.biff.BuiltInName;
import jxl.biff.AutoFilter;
import jxl.biff.CellFinder;
import jxl.biff.CellReferenceHelper;
import jxl.biff.ConditionalFormat;
import jxl.biff.DataValidation;
import jxl.biff.EmptyCell;
import jxl.biff.FormattingRecords;
import jxl.biff.Type;
import jxl.biff.WorkspaceInformationRecord;
import jxl.biff.drawing.Chart;
import jxl.biff.drawing.Drawing;
import jxl.biff.drawing.DrawingData;
import jxl.biff.drawing.DrawingGroupObject;
import jxl.format.CellFormat;

/**
 * Represents a sheet within a workbook.  Provides a handle to the individual
 * cells, or lines of cells (grouped by Row or Column)
 * In order to simplify this class due to code bloat, the actual reading
 * logic has been delegated to the SheetReaderClass.  This class' main
 * responsibility is now to implement the API methods declared in the
 * Sheet interface
 */
public class SheetImpl implements Sheet
{
  /**
   * The logger
   */
  private static Logger logger = Logger.getLogger(SheetImpl.class);

  /**
   * The excel file
   */
  private File excelFile;
  /**
   * A handle to the shared string table
   */
  private SSTRecord sharedStrings;

  /**
   * A handle to the sheet BOF record, which indicates the stream type
   */
  private BOFRecord sheetBof;

  /**
   * A handle to the workbook BOF record, which indicates the stream type
   */
  private BOFRecord workbookBof;

  /**
   * A handle to the formatting records
   */
  private FormattingRecords formattingRecords;

  /**
   * The name of this sheet
   */
  private String name;

  /**
   * The  number of rows
   */
  private int numRows;

  /**
   * The number of columns
   */
  private int numCols;

  /**
   * The cells
   */
  private Cell[][] cells;

  /**
   * The start position in the stream of this sheet
   */
  private int startPosition;

  /**
   * The list of specified (ie. non default) column widths
   */
  private ColumnInfoRecord[] columnInfos;

  /**
   * The array of row records
   */
  private RowRecord[] rowRecords;

  /**
   * The list of non-default row properties
   */
  private ArrayList rowProperties;

  /**
   * An array of column info records.  They are held this way before
   * they are transferred to the more convenient array
   */
  private ArrayList columnInfosArray;

  /**
   * A list of shared formula groups
   */
  private ArrayList sharedFormulas;

  /**
   * A list of hyperlinks on this page
   */
  private ArrayList hyperlinks;

  /**
   * A list of charts on this page
   */
  private ArrayList charts;

  /**
   * A list of drawings on this page
   */
  private ArrayList drawings;

  /**
   * A list of drawings (as opposed to comments/validation/charts) on this
   * page
   */
  private ArrayList images;

  /**
   * A list of data validations on this page
   */
  private DataValidation dataValidation;

  /**
   * A list of merged cells on this page
   */
  private Range[] mergedCells;

  /**
   * Indicates whether the columnInfos array has been initialized
   */
  private boolean columnInfosInitialized;

  /**
   * Indicates whether the rowRecords array has been initialized
   */
  private boolean rowRecordsInitialized;

  /**
   * Indicates whether or not the dates are based around the 1904 date system
   */
  private boolean nineteenFour;

  /**
   * The workspace options
   */
  private WorkspaceInformationRecord workspaceOptions;

  /**
   * The hidden flag
   */
  private boolean hidden;

  /**
   * The environment specific print record
   */
  private PLSRecord plsRecord;

  /**
   * The property set record associated with this workbook
   */
  private ButtonPropertySetRecord buttonPropertySet;

  /**
   * The sheet settings
   */
  private SheetSettings settings;

  /**
   * The horizontal page breaks contained on this sheet
   */
  private int[] rowBreaks;

  /**
   * The vertical page breaks contained on this sheet
   */
  private int[] columnBreaks;

  /**
   * The maximum row outline level
   */
  private int maxRowOutlineLevel;

  /**
   * The maximum column outline level
   */
  private int maxColumnOutlineLevel;

  /**
   * The list of local names for this sheet
   */
  private ArrayList localNames;

  /**
   * The list of conditional formats for this sheet
   */
  private ArrayList conditionalFormats;

  /**
   * The autofilter information
   */
  private AutoFilter autoFilter;

  /**
   * A handle to the workbook which contains this sheet.  Some of the records
   * need this in order to reference external sheets
   */
  private WorkbookParser workbook;

  /**
   * A handle to the workbook settings
   */
  private WorkbookSettings workbookSettings;

  /**
   * Constructor
   *
   * @param f the excel file
   * @param sst the shared string table
   * @param fr formatting records
   * @param sb the bof record which indicates the start of the sheet
   * @param wb the bof record which indicates the start of the sheet
   * @param nf the 1904 flag
   * @param wp the workbook which this sheet belongs to
   * @exception BiffException
   */
  SheetImpl(File f,
            SSTRecord sst,
            FormattingRecords fr,
            BOFRecord sb,
            BOFRecord wb,
            boolean nf,
            WorkbookParser wp)
    throws BiffException
  {
    excelFile = f;
    sharedStrings = sst;
    formattingRecords = fr;
    sheetBof = sb;
    workbookBof = wb;
    columnInfosArray = new ArrayList();
    sharedFormulas = new ArrayList();
    hyperlinks = new ArrayList();
    rowProperties = new ArrayList(10);
    columnInfosInitialized = false;
    rowRecordsInitialized = false;
    nineteenFour = nf;
    workbook = wp;
    workbookSettings = workbook.getSettings();

    // Mark the position in the stream, and then skip on until the end
    startPosition = f.getPos();

    if (sheetBof.isChart())
    {
      // Set the start pos to include the bof so the sheet reader can handle it
      startPosition -= (sheetBof.getLength() + 4);
    }

    Record r = null;
    int bofs = 1;

    while (bofs >= 1)
    {
      r = f.next();

      // use this form for quick performance
      if (r.getCode() == Type.EOF.value)
      {
        bofs--;
      }

      if (r.getCode() == Type.BOF.value)
      {
        bofs++;
      }
    }
  }

  /**
   * Returns the cell for the specified location eg. "A4", using the
   * CellReferenceHelper
   *
   * @param loc the cell reference
   * @return the cell at the specified co-ordinates
   */
  public Cell getCell(String loc)
  {
    return getCell(CellReferenceHelper.getColumn(loc),
                   CellReferenceHelper.getRow(loc));
  }

  /**
   * Returns the cell specified at this row and at this column
   *
   * @param row the row number
   * @param column the column number
   * @return the cell at the specified co-ordinates
   */
  public Cell getCell(int column, int row)
  {
    // just in case this has been cleared, but something else holds
    // a reference to it
    if (cells == null)
    {
      readSheet();
    }

    Cell c = cells[row][column];

    if (c == null)
    {
      c = new EmptyCell(column, row);
      cells[row][column] = c;
    }

    return c;
  }

  /**
   * Gets the cell whose contents match the string passed in.
   * If no match is found, then null is returned.  The search is performed
   * on a row by row basis, so the lower the row number, the more
   * efficiently the algorithm will perform
   *
   * @param  contents the string to match
   * @return the Cell whose contents match the paramter, null if not found
   */
  public Cell findCell(String contents)
  {
    CellFinder cellFinder = new CellFinder(this);
    return cellFinder.findCell(contents);
  }

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
                       boolean reverse)
  {
    CellFinder cellFinder = new CellFinder(this);
    return cellFinder.findCell(contents,
                               firstCol, 
                               firstRow, 
                               lastCol,
                               lastRow, 
                               reverse);
  }

  /**
   * Gets the cell whose contents match the regular expressionstring passed in.
   * If no match is found, then null is returned.  The search is performed
   * on a row by row basis, so the lower the row number, the more
   * efficiently the algorithm will perform
   * 
   * @param pattern the regular expression string to match
   * @param firstCol the first column within the range
   * @param firstRow the first row of the range
   * @param lastRow the last row within the range
   * @param lastCol the last column within the ranage
   * @param reverse indicates whether to perform a reverse search or not
   * @return the Cell whose contents match the parameter, null if not found
   */
  public Cell findCell(Pattern pattern, 
                       int firstCol, 
                       int firstRow, 
                       int lastCol, 
                       int lastRow, 
                       boolean reverse)
  {
    CellFinder cellFinder = new CellFinder(this);
    return cellFinder.findCell(pattern,
                               firstCol, 
                               firstRow, 
                               lastCol,
                               lastRow, 
                               reverse);
  }

  /**
   * Gets the cell whose contents match the string passed in.
   * If no match is found, then null is returned.  The search is performed
   * on a row by row basis, so the lower the row number, the more
   * efficiently the algorithm will perform.  This method differs
   * from the findCell methods in that only cells with labels are
   * queried - all numerical cells are ignored.  This should therefore
   * improve performance.
   *
   * @param  contents the string to match
   * @return the Cell whose contents match the paramter, null if not found
   */
  public LabelCell findLabelCell(String contents)
  {
    CellFinder cellFinder = new CellFinder(this);
    return cellFinder.findLabelCell(contents);
  }

  /**
   * Returns the number of rows in this sheet
   *
   * @return the number of rows in this sheet
   */
  public int getRows()
  {
    // just in case this has been cleared, but something else holds
    // a reference to it
    if (cells == null)
    {
      readSheet();
    }

    return numRows;
  }

  /**
   * Returns the number of columns in this sheet
   *
   * @return the number of columns in this sheet
   */
  public int getColumns()
  {
    // just in case this has been cleared, but something else holds
    // a reference to it
    if (cells == null)
    {
      readSheet();
    }

    return numCols;
  }

  /**
   * Gets all the cells on the specified row.  The returned array will
   * be stripped of all trailing empty cells
   *
   * @param row the rows whose cells are to be returned
   * @return the cells on the given row
   */
  public Cell[] getRow(int row)
  {
    // just in case this has been cleared, but something else holds
    // a reference to it
    if (cells == null)
    {
      readSheet();
    }

    // Find the last non-null cell
    boolean found = false;
    int col = numCols - 1;
    while (col >= 0 && !found)
    {
      if (cells[row][col] != null)
      {
        found = true;
      }
      else
      {
        col--;
      }
    }

    // Only create entries for non-null cells
    Cell[] c = new Cell[col + 1];

    for (int i = 0; i <= col; i++)
    {
      c[i] = getCell(i, row);
    }
    return c;
  }

  /**
   * Gets all the cells on the specified column.  The returned array
   * will be stripped of all trailing empty cells
   *
   * @param col the column whose cells are to be returned
   * @return the cells on the specified column
   */
  public Cell[] getColumn(int col)
  {
    // just in case this has been cleared, but something else holds
    // a reference to it
    if (cells == null)
    {
      readSheet();
    }

    // Find the last non-null cell
    boolean found = false;
    int row = numRows - 1;
    while (row >= 0 && !found)
    {
      if (cells[row][col] != null)
      {
        found = true;
      }
      else
      {
        row--;
      }
    }

    // Only create entries for non-null cells
    Cell[] c = new Cell[row + 1];

    for (int i = 0; i <= row; i++)
    {
      c[i] = getCell(col, i);
    }
    return c;
  }

  /**
   * Gets the name of this sheet
   *
   * @return the name of the sheet
   */
  public String getName()
  {
    return name;
  }

  /**
   * Sets the name of this sheet
   *
   * @param s the sheet name
   */
  final void setName(String s)
  {
    name = s;
  }

  /**
   * Determines whether the sheet is hidden
   *
   * @return whether or not the sheet is hidden
   * @deprecated in favour of the getSettings function
   */
  public boolean isHidden()
  {
    return hidden;
  }

  /**
   * Gets the column info record for the specified column.  If no
   * column is specified, null is returned
   *
   * @param col the column
   * @return the ColumnInfoRecord if specified, NULL otherwise
   */
  public ColumnInfoRecord getColumnInfo(int col)
  {
    if (!columnInfosInitialized)
    {
      // Initialize the array
      Iterator i = columnInfosArray.iterator();
      ColumnInfoRecord cir = null;
      while (i.hasNext())
      {
        cir = (ColumnInfoRecord) i.next();

        int startcol = Math.max(0, cir.getStartColumn());
        int endcol = Math.min(columnInfos.length - 1, cir.getEndColumn());

        for (int c = startcol; c <= endcol; c++)
        {
          columnInfos[c] = cir;
        }

        if (endcol < startcol)
        {
          columnInfos[startcol] = cir;
        }
      }

      columnInfosInitialized = true;
    }

    return col < columnInfos.length ? columnInfos[col] : null;
  }

  /**
   * Gets all the column info records
   *
   * @return the ColumnInfoRecordArray
   */
  public ColumnInfoRecord[] getColumnInfos()
  {
    // Just chuck all the column infos we have into an array
    ColumnInfoRecord[] infos = new ColumnInfoRecord[columnInfosArray.size()];
    for (int i = 0; i < columnInfosArray.size(); i++)
    {
      infos[i] = (ColumnInfoRecord) columnInfosArray.get(i);
    }

    return infos;
  }

  /**
   * Sets the visibility of this sheet
   *
   * @param h hidden flag
   */
  final void setHidden(boolean h)
  {
    hidden = h;
  }

  /**
   * Clears out the array of cells.  This is done for memory allocation
   * reasons when reading very large sheets
   */
  final void clear()
  {
    cells = null;
    mergedCells = null;
    columnInfosArray.clear();
    sharedFormulas.clear();
    hyperlinks.clear();
    columnInfosInitialized = false;

    if (!workbookSettings.getGCDisabled())
    {
      System.gc();
    }
  }

  /**
   * Reads in the contents of this sheet
   */
  final void readSheet()
  {
    // If this sheet contains only a chart, then set everything to
    // empty and do not bother parsing the sheet
    // Thanks to steve.brophy for spotting this
    if (!sheetBof.isWorksheet())
    {
      numRows = 0;
      numCols = 0;
      cells = new Cell[0][0];
      //      return;
    }

    SheetReader reader = new SheetReader(excelFile,
                                         sharedStrings,
                                         formattingRecords,
                                         sheetBof,
                                         workbookBof,
                                         nineteenFour,
                                         workbook,
                                         startPosition,
                                         this);
    reader.read();

    // Take stuff that was read in
    numRows = reader.getNumRows();
    numCols = reader.getNumCols();
    cells = reader.getCells();
    rowProperties = reader.getRowProperties();
    columnInfosArray = reader.getColumnInfosArray();
    hyperlinks = reader.getHyperlinks();
    conditionalFormats = reader.getConditionalFormats();
    autoFilter = reader.getAutoFilter();
    charts = reader.getCharts();
    drawings = reader.getDrawings();
    dataValidation = reader.getDataValidation();
    mergedCells = reader.getMergedCells();
    settings = reader.getSettings();
    settings.setHidden(hidden);
    rowBreaks = reader.getRowBreaks();
    columnBreaks = reader.getColumnBreaks();
    workspaceOptions = reader.getWorkspaceOptions();
    plsRecord = reader.getPLS();
    buttonPropertySet = reader.getButtonPropertySet();
    maxRowOutlineLevel = reader.getMaxRowOutlineLevel();
    maxColumnOutlineLevel = reader.getMaxColumnOutlineLevel();

    reader = null;

    if (!workbookSettings.getGCDisabled())
    {
      System.gc();
    }

    if (columnInfosArray.size() > 0)
    {
      ColumnInfoRecord cir = (ColumnInfoRecord)
        columnInfosArray.get(columnInfosArray.size() - 1);
      columnInfos = new ColumnInfoRecord[cir.getEndColumn() + 1];
    }
    else
    {
      columnInfos = new ColumnInfoRecord[0];
    }

    // Add any local names
    if (localNames != null)
    {
      for (Iterator it = localNames.iterator(); it.hasNext() ;)
      {
        NameRecord nr = (NameRecord) it.next();
        if (nr.getBuiltInName() == BuiltInName.PRINT_AREA)
        {
          if(nr.getRanges().length > 0)
          {
            NameRecord.NameRange rng = nr.getRanges()[0];
            settings.setPrintArea(rng.getFirstColumn(),
                                  rng.getFirstRow(),
                                  rng.getLastColumn(),
                                  rng.getLastRow());
          }
        }
        else if (nr.getBuiltInName() == BuiltInName.PRINT_TITLES)
       	{
          // There can be 1 or 2 entries.  
          // Row entries have hardwired column entries (first and last
          //  possible column)
          // Column entries have hardwired row entries (first and last 
          // possible row)
          for (int i = 0 ; i < nr.getRanges().length ; i++)
          {
            NameRecord.NameRange rng = nr.getRanges()[i];
            if (rng.getFirstColumn() == 0 && rng.getLastColumn() == 255)
            {
              settings.setPrintTitlesRow(rng.getFirstRow(),
                                         rng.getLastRow());
            }
            else
            {
              settings.setPrintTitlesCol(rng.getFirstColumn(),
                                         rng.getLastColumn());
            }
          }
        }
      }
    }
  }

  /**
   * Gets the hyperlinks on this sheet
   *
   * @return an array of hyperlinks
   */
  public Hyperlink[] getHyperlinks()
  {
    Hyperlink[] hl = new Hyperlink[hyperlinks.size()];

    for (int i = 0; i < hyperlinks.size(); i++)
    {
      hl[i] = (Hyperlink) hyperlinks.get(i);
    }

    return hl;
  }

  /**
   * Gets the cells which have been merged on this sheet
   *
   * @return an array of range objects
   */
  public Range[] getMergedCells()
  {
    if (mergedCells == null)
    {
      return new Range[0];
    }

    return mergedCells;
  }

  /**
   * Gets the non-default rows.  Used when copying spreadsheets
   *
   * @return an array of row properties
   */
  public RowRecord[] getRowProperties()
  {
    RowRecord[] rp = new RowRecord[rowProperties.size()];
    for (int i = 0; i < rp.length; i++)
    {
      rp[i] = (RowRecord) rowProperties.get(i);
    }

    return rp;
  }

  /**
   * Gets the data validations.  Used when copying sheets
   *
   * @return the data validations
   */
  public DataValidation getDataValidation()
  {
    return dataValidation;
  }

  /**
   * Gets the row record.  Usually called by the cell in the specified
   * row in order to determine its size
   *
   * @param r the row
   * @return the RowRecord for the specified row
   */
  RowRecord getRowInfo(int r)
  {
    if (!rowRecordsInitialized)
    {
      rowRecords = new RowRecord[getRows()];
      Iterator i = rowProperties.iterator();

      int rownum = 0;
      RowRecord rr = null;
      while (i.hasNext())
      {
        rr = (RowRecord) i.next();
        rownum = rr.getRowNumber();
        if (rownum < rowRecords.length)
        {
          rowRecords[rownum] = rr;
        }
      }

      rowRecordsInitialized = true;
    }

    return r < rowRecords.length ? rowRecords[r] : null;
  }

  /**
   * Gets the row breaks.  Called when copying sheets
   *
   * @return the explicit row breaks
   */
  public final int[] getRowPageBreaks()
  {
    return rowBreaks;
  }

  /**
   * Gets the row breaks.  Called when copying sheets
   *
   * @return the explicit row breaks
   */
  public final int[] getColumnPageBreaks()
  {
    return columnBreaks;
  }

  /**
   * Gets the charts.  Called when copying sheets
   *
   * @return the charts on this page
   */
  public final Chart[] getCharts()
  {
    Chart[] ch = new Chart[charts.size()];

    for (int i = 0; i < ch.length; i++)
    {
      ch[i] = (Chart) charts.get(i);
    }
    return ch;
  }

  /**
   * Gets the drawings.  Called when copying sheets
   *
   * @return the drawings on this page
   */
  public final DrawingGroupObject[] getDrawings()
  {
    DrawingGroupObject[] dr = new DrawingGroupObject[drawings.size()];
    dr = (DrawingGroupObject[]) drawings.toArray(dr);
    return dr;
  }

  /**
   * Determines whether the sheet is protected
   *
   * @return whether or not the sheet is protected
   * @deprecated in favour of the getSettings() api
   */
  public boolean isProtected()
  {
    return settings.isProtected();
  }

  /**
   * Gets the workspace options for this sheet.  Called during the copy
   * process
   *
   * @return the workspace options
   */
  public WorkspaceInformationRecord getWorkspaceOptions()
  {
    return workspaceOptions;
  }

  /**
   * Accessor for the sheet settings
   *
   * @return the settings for this sheet
   */
  public SheetSettings getSettings()
  {
    return settings;
  }

  /**
   * Accessor for the workbook.  In addition to be being used by this package,
   * it is also used during the importSheet process
   *
   * @return  the workbook
   */
  public WorkbookParser getWorkbook()
  {
    return workbook;
  }

  /**
   * Gets the column format for the specified column
   *
   * @param col the column number
   * @return the column format, or NULL if the column has no specific format
   * @deprecated use getColumnView instead
   */
  public CellFormat getColumnFormat(int col)
  {
    CellView cv = getColumnView(col);
    return cv.getFormat();
  }

  /**
   * Gets the column width for the specified column
   *
   * @param col the column number
   * @return the column width, or the default width if the column has no
   *         specified format
   */
  public int getColumnWidth(int col)
  {
    return getColumnView(col).getSize() / 256;
  }

  /**
   * Gets the column width for the specified column
   *
   * @param col the column number
   * @return the column format, or the default format if no override is
             specified
   */
  public CellView getColumnView(int col)
  {
    ColumnInfoRecord cir = getColumnInfo(col);
    CellView cv = new CellView();

    if (cir != null)
    {
      cv.setDimension(cir.getWidth() / 256); //deprecated
      cv.setSize(cir.getWidth());
      cv.setHidden(cir.getHidden());
      cv.setFormat(formattingRecords.getXFRecord(cir.getXFIndex()));
    }
    else
    {
      cv.setDimension(settings.getDefaultColumnWidth()); //deprecated
      cv.setSize(settings.getDefaultColumnWidth() * 256);
    }

    return cv;
  }

  /**
   * Gets the row height for the specified column
   *
   * @param row the row number
   * @return the row height, or the default height if the row has no
   *         specified format
   * @deprecated use getRowView instead
   */
  public int getRowHeight(int row)
  {
    return getRowView(row).getDimension();
  }

  /**
   * Gets the row view for the specified row
   *
   * @param row the row number
   * @return the row format, or the default format if no override is
             specified
   */
  public CellView getRowView(int row)
  {
    RowRecord rr = getRowInfo(row);

    CellView cv = new CellView();

    if (rr != null)
    {
      cv.setDimension(rr.getRowHeight()); //deprecated
      cv.setSize(rr.getRowHeight());
      cv.setHidden(rr.isCollapsed());
      if (rr.hasDefaultFormat())
      {
        cv.setFormat(formattingRecords.getXFRecord(rr.getXFIndex()));
      }
    }
    else
    {
      cv.setDimension(settings.getDefaultRowHeight());
      cv.setSize(settings.getDefaultRowHeight()); //deprecated
    }

    return cv;
  }


  /**
   * Used when copying sheets in order to determine the type of this sheet
   *
   * @return the BOF Record
   */
  public BOFRecord getSheetBof()
  {
    return sheetBof;
  }

  /**
   * Used when copying sheets in order to determine the type of the containing
   * workboook
   *
   * @return the workbook BOF Record
   */
  public BOFRecord getWorkbookBof()
  {
    return workbookBof;
  }

  /**
   * Accessor for the environment specific print record, invoked when
   * copying sheets
   *
   * @return the environment specific print record
   */
  public PLSRecord getPLS()
  {
    return plsRecord;
  }

  /**
   * Accessor for the button property set, used during copying
   *
   * @return the button property set
   */
  public ButtonPropertySetRecord getButtonPropertySet()
  {
    return buttonPropertySet;
  }

  /**
   * Accessor for the number of images on the sheet
   *
   * @return the number of images on this sheet
   */
  public int getNumberOfImages()
  {
    if (images == null)
    {
      initializeImages();
    }

    return images.size();
  }

  /**
   * Accessor for the image
   *
   * @param i the 0 based image number
   * @return  the image at the specified position
   */
  public Image getDrawing(int i)
  {
    if (images == null)
    {
      initializeImages();
    }

    return (Image) images.get(i);
  }

  /**
   * Initializes the images
   */
  private void initializeImages()
  {
    if (images != null)
    {
      return;
    }

    images = new ArrayList();
    DrawingGroupObject[] dgos = getDrawings();

    for (int i = 0; i < dgos.length; i++)
    {
      if (dgos[i] instanceof Drawing)
      {
        images.add(dgos[i]);
      }
    }
  }

  /**
   * Used by one of the demo programs for debugging purposes only
   */
  public DrawingData getDrawingData()
  {
    SheetReader reader = new SheetReader(excelFile,
                                         sharedStrings,
                                         formattingRecords,
                                         sheetBof,
                                         workbookBof,
                                         nineteenFour,
                                         workbook,
                                         startPosition,
                                         this);
    reader.read();
    return reader.getDrawingData();
  }

  /**
   * Adds a local name to this shate
   *
   * @param nr the local name to add
   */
  void addLocalName(NameRecord nr)
  {
    if (localNames == null)
    {
      localNames = new ArrayList();
    }

    localNames.add(nr);
  }

  /**
   * Gets the conditional formats
   *
   * @return the conditional formats
   */
  public ConditionalFormat[] getConditionalFormats()
  {
    ConditionalFormat[] formats = 
      new ConditionalFormat[conditionalFormats.size()];
    formats = (ConditionalFormat[]) conditionalFormats.toArray(formats);
    return formats;
  }

  /**
   * Returns the autofilter
   *
   * @return the autofilter
   */
  public AutoFilter getAutoFilter()
  {
    return autoFilter;
  }

  /** 
   * Accessor for the maximum column outline level.  Used during a copy
   *
   * @return the maximum column outline level, or 0 if no outlines/groups
   */
  public int getMaxColumnOutlineLevel() 
  {
    return maxColumnOutlineLevel;
  }

  /** 
   * Accessor for the maximum row outline level.  Used during a copy
   *
   * @return the maximum row outline level, or 0 if no outlines/groups
   */
  public int getMaxRowOutlineLevel() 
  {
    return maxRowOutlineLevel;
  }

}
