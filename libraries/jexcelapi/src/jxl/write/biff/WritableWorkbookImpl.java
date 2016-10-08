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

package jxl.write.biff;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;

import jxl.common.Assert;
import jxl.common.Logger;

import jxl.Range;
import jxl.Sheet;
import jxl.Workbook;
import jxl.WorkbookSettings;
import jxl.biff.BuiltInName;
import jxl.biff.CellReferenceHelper;
import jxl.biff.CountryCode;
import jxl.biff.Fonts;
import jxl.biff.FormattingRecords;
import jxl.biff.IndexMapping;
import jxl.biff.IntegerHelper;
import jxl.biff.RangeImpl;
import jxl.biff.WorkbookMethods;
import jxl.biff.XCTRecord;
import jxl.biff.drawing.Drawing;
import jxl.biff.drawing.DrawingGroup;
import jxl.biff.drawing.DrawingGroupObject;
import jxl.biff.drawing.Origin;
import jxl.biff.formula.ExternalSheet;
import jxl.format.Colour;
import jxl.format.RGB;
import jxl.read.biff.WorkbookParser;
import jxl.write.WritableCell;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;


/**
 * A writable workbook
 */
public class WritableWorkbookImpl extends WritableWorkbook 
  implements ExternalSheet, WorkbookMethods
{
  /**
   * The logger
   */
  private static Logger logger = Logger.getLogger(WritableWorkbookImpl.class);
  /**
   * The list of formats available within this workbook
   */
  private FormattingRecords formatRecords;
  /**
   * The output file to write the workbook to
   */
  private File outputFile;
  /**
   * The list of sheets within this workbook
   */
  private ArrayList sheets;
  /**
   * The list of fonts available within this workbook
   */
  private Fonts fonts;
  /**
   * The list of external sheets, used by cell references in formulas
   */
  private ExternalSheetRecord externSheet;

  /**
   * The supbook records
   */
  private ArrayList supbooks;

  /**
   * The name records
   */
  private ArrayList names;

  /**
   * A lookup hash map of the name records
   */
  private HashMap nameRecords;

  /**
   * The shared strings used by this workbook
   */
  private SharedStrings sharedStrings;

  /**
   * Indicates whether or not the output stream should be closed.  This
   * depends on whether this Workbook was created with an output stream,
   * or a flat file (flat file closes the stream
   */
  private boolean closeStream;

  /**
   * The workbook protection flag
   */
  private boolean wbProtected;

  /**
   * The settings for the workbook
   */
  private WorkbookSettings settings;

  /**
   * The list of cells for the entire workbook which need to be updated 
   * following a row/column insert or remove
   */
  private ArrayList rcirCells;

  /**
   * The drawing group
   */
  private DrawingGroup drawingGroup;

  /**
   * The jxl.common.workbook styles
   */
  private Styles styles;

  /**
   * Contains macros flag
   */
  private boolean containsMacros;

  /**
   * The buttons property set
   */
  private ButtonPropertySetRecord buttonPropertySet;

  /**
   * The country record, initialised when copying a workbook
   */
  private CountryRecord countryRecord;

  // synchronizer object for static unitiatialization
  private static Object SYNCHRONIZER = new Object();
  /**
   * The names of any add in functions
   */
  private String[] addInFunctionNames;

  /**
   * The XCT records
   */
  private XCTRecord[] xctRecords;

  /**
   * Constructor.  Writes the workbook direct to the existing output stream
   * 
   * @exception IOException 
   * @param os the output stream
   * @param cs TRUE if the workbook should close the output stream, FALSE
   * @param ws the configuration for this workbook
   * otherwise
   */
  public WritableWorkbookImpl(OutputStream os, boolean cs, WorkbookSettings ws)
    throws IOException
  {
    super();
    outputFile = new File(os, ws, null);
    sheets = new ArrayList();
    sharedStrings = new SharedStrings();
    nameRecords = new HashMap();
    closeStream = cs;
    wbProtected = false;
    containsMacros = false;
    settings = ws;
    rcirCells = new ArrayList();
    styles = new Styles();

    // Reset the statically declared styles.  These are no longer needed
    // because the Styles class will intercept all calls within 
    // CellValue.setCellDetails and if it detects a standard format, then it
    // will return a clone.  In short, the static cell values will
    // never get initialized anyway.  Still, just to be extra sure...
    synchronized(SYNCHRONIZER)
    {
      WritableWorkbook.ARIAL_10_PT.uninitialize();
      WritableWorkbook.HYPERLINK_FONT.uninitialize();
      WritableWorkbook.NORMAL_STYLE.uninitialize();
      WritableWorkbook.HYPERLINK_STYLE.uninitialize();
      WritableWorkbook.HIDDEN_STYLE.uninitialize();
      DateRecord.defaultDateFormat.uninitialize();
     }

    WritableFonts wf = new WritableFonts(this);
    fonts = wf;

    WritableFormattingRecords wfr = new WritableFormattingRecords(fonts, 
                                                                  styles);
    formatRecords = wfr;
  }

  /**
   * A pseudo copy constructor.  Takes the handles to the font and formatting
   * records
   * 
   * @exception IOException 
   * @param w the workbook to copy
   * @param os the output stream to write the data to
   * @param cs TRUE if the workbook should close the output stream, FALSE
   * @param ws the configuration for this workbook
   */
  public WritableWorkbookImpl(OutputStream os,
                              Workbook w,
                              boolean cs,
                              WorkbookSettings ws) throws IOException
  {
    super();
    WorkbookParser wp = (WorkbookParser) w;

    // Reset the statically declared styles.  These are no longer needed
    // because the Styles class will intercept all calls within 
    // CellValue.setCellDetails and if it detects a standard format, then it
    // will return a clone.  In short, the static cell values will
    // never get initialized anyway.  Still, just to be extra sure...
    synchronized(SYNCHRONIZER)
    {
      WritableWorkbook.ARIAL_10_PT.uninitialize();
      WritableWorkbook.HYPERLINK_FONT.uninitialize();
      WritableWorkbook.NORMAL_STYLE.uninitialize();
      WritableWorkbook.HYPERLINK_STYLE.uninitialize();
      WritableWorkbook.HIDDEN_STYLE.uninitialize();
      DateRecord.defaultDateFormat.uninitialize();
     }

    closeStream = cs;
    sheets = new ArrayList();
    sharedStrings = new SharedStrings();
    nameRecords = new HashMap();
    fonts = wp.getFonts();
    formatRecords = wp.getFormattingRecords();
    wbProtected = false;
    settings = ws;
    rcirCells = new ArrayList();
    styles = new Styles();
    outputFile = new File(os, ws, wp.getCompoundFile());

    containsMacros = false;
    if (!ws.getPropertySetsDisabled())
    {
      containsMacros = wp.containsMacros();
    }

    // Copy the country settings
    if (wp.getCountryRecord() != null)
    {
      countryRecord = new CountryRecord(wp.getCountryRecord());
    }

    // Copy any add in functions
    addInFunctionNames = wp.getAddInFunctionNames();

    // Copy XCT records
    xctRecords = wp.getXCTRecords();

    // Copy any external sheets
    if (wp.getExternalSheetRecord() != null)
    {
      externSheet = new ExternalSheetRecord(wp.getExternalSheetRecord());

      // Get the associated supbooks
      jxl.read.biff.SupbookRecord[] readsr = wp.getSupbookRecords();
      supbooks = new ArrayList(readsr.length);

      for (int i = 0; i < readsr.length; i++)
      {
        jxl.read.biff.SupbookRecord readSupbook = readsr[i];
        if (readSupbook.getType() == readSupbook.INTERNAL ||
            readSupbook.getType() == readSupbook.EXTERNAL)
        {
          supbooks.add(new SupbookRecord(readSupbook, settings));
        }
        else
        {
          if (readSupbook.getType() != readSupbook.ADDIN)
          {
            logger.warn("unsupported supbook type - ignoring");
          }
        }
      }
    }

    // Copy any drawings.  These must be present before we try and copy
    // the images from the read workbook
    if (wp.getDrawingGroup() != null)
    {
      drawingGroup = new DrawingGroup(wp.getDrawingGroup());
    }

    // Copy the property set references
    if (containsMacros && wp.getButtonPropertySet() != null)
    {
      buttonPropertySet = new ButtonPropertySetRecord
        (wp.getButtonPropertySet());
    }

    // Copy any names
    if (!settings.getNamesDisabled())
    {
      jxl.read.biff.NameRecord[] na = wp.getNameRecords();
      names = new ArrayList(na.length);
      
      for (int i = 0; i < na.length; i++)
      {
        if (na[i].isBiff8())
        {
          NameRecord n = new NameRecord(na[i], i);
          names.add(n);
          String name = n.getName();
          nameRecords.put(name, n);
        }
        else
        {
          logger.warn("Cannot copy Biff7 name records - ignoring");
        }
      }
    }
    
    copyWorkbook(w);

    // The copy process may have caused some critical fields in the 
    // read drawing group to change.  Make sure these updates are reflected
    // in the writable drawing group
    if (drawingGroup != null)
    {
      drawingGroup.updateData(wp.getDrawingGroup());
    }
  }

  /**
   * Gets the sheets within this workbook.  Use of this method for
   * large worksheets can cause performance problems.
   * 
   * @return an array of the individual sheets
   */
  public WritableSheet[] getSheets()
  {
    WritableSheet[] sheetArray = new WritableSheet[getNumberOfSheets()];
    
    for (int i = 0 ; i < getNumberOfSheets() ; i++)
    {
      sheetArray[i] = getSheet(i);
    }
    return sheetArray;
  }

  /**
   * Gets the sheet names
   *
   * @return an array of strings containing the sheet names
   */
  public String[] getSheetNames()
  {
    String[] sheetNames = new String[getNumberOfSheets()];

    for (int i = 0 ; i < sheetNames.length ; i++)
    {
      sheetNames[i] = getSheet(i).getName();
    }

    return sheetNames;
  }

  /**
   * Interface method from WorkbookMethods - gets the specified 
   * sheet within this workbook
   *
   * @param index the zero based index of the required sheet
   * @return The sheet specified by the index
   */
  public Sheet getReadSheet(int index)
  {
    return getSheet(index);
  }

  /**
   * Gets the specified sheet within this workbook
   * 
   * @param index the zero based index of the reQuired sheet
   * @return The sheet specified by the index
   */
  public WritableSheet getSheet(int index)
  {
    return (WritableSheet) sheets.get(index);
  }

  /**
   * Gets the sheet with the specified name from within this workbook
   * 
   * @param name the sheet name
   * @return The sheet with the specified name, or null if it is not found
   */
  public WritableSheet getSheet(String name)
  {
    // Iterate through the boundsheet records
    boolean found = false;
    Iterator i = sheets.iterator();
    WritableSheet s = null;

    while (i.hasNext() && !found)
    {
      s = (WritableSheet) i.next();
      
      if (s.getName().equals(name))
      {
        found = true;
      }
    }

    return found ? s : null;
  }

  /**
   * Returns the number of sheets in this workbook
   * 
   * @return the number of sheets in this workbook
   */
  public int getNumberOfSheets()
  {
    return sheets.size();
  }

  /**
   * Closes this workbook, and frees makes any memory allocated available
   * for garbage collection
   * 
   * @exception IOException 
   * @exception JxlWriteException
   */
  public void close() throws IOException, JxlWriteException
  {
    outputFile.close(closeStream);
  }

  /**
   * Sets a new output file.  This allows the smae workbook to be
   * written to various different output files without having to
   * read in any templates again
   *
   * @param fileName the file name
   * @exception IOException
   */
  public void setOutputFile(java.io.File fileName) throws IOException
  {
    FileOutputStream fos = new FileOutputStream(fileName);
    outputFile.setOutputFile(fos);
  }


  /**
   * The internal method implementation for creating new sheets
   *
   * @param name
   * @param index
   * @param handleRefs flag indicating whether or not to handle external
   *                   sheet references
   * @return 
   */
  private WritableSheet createSheet(String name, int index, 
                                    boolean handleRefs)
  {
    WritableSheet w = new WritableSheetImpl(name, 
                                            outputFile, 
                                            formatRecords, 
                                            sharedStrings,
                                            settings,
                                            this);

    int pos = index;

    if (index <= 0)
    {
      pos = 0;
      sheets.add(0, w);
    }
    else if (index > sheets.size())
    {
      pos = sheets.size();
      sheets.add(w);
    }
    else
    {
      sheets.add(index, w);
    }

    if (handleRefs && externSheet != null)
    {
      externSheet.sheetInserted(pos);
    }

    if (supbooks != null && supbooks.size() > 0)
    {
      SupbookRecord supbook = (SupbookRecord) supbooks.get(0);
      if (supbook.getType() == SupbookRecord.INTERNAL)
      {
        supbook.adjustInternal(sheets.size());
      }
    }

    return w;
  }

  /**
   * Creates a new sheet within the workbook, at the specified position.
   * The new sheet is inserted at the specified position, or prepended/appended
   * to the list of sheets if the index specified is somehow inappropriate
   * 
   * @param name the name of the new sheet
   * @param index the index at which to add the sheet
   * @return the created sheet
   */
  public WritableSheet createSheet(String name, int index)
  {
    return createSheet(name, index, true);
  }

  /**
   * Removes a sheet from this workbook, the other sheets indices being
   * altered accordingly. If the sheet referenced by the index
   * does not exist, then no action is taken.
   * 
   * @param index the index of the sheet to remove
   */
  public void removeSheet(int index)
  {
    int pos = index;
    if (index <= 0)
    {
      pos = 0;
      sheets.remove(0);
    }
    else if (index >= sheets.size())
    {
      pos = sheets.size() - 1;
      sheets.remove(sheets.size() - 1);
    }
    else
    {
      sheets.remove(index);
    }

    if (externSheet != null)
    {
      externSheet.sheetRemoved(pos);
    }

    if (supbooks != null && supbooks.size() > 0)
    {
      SupbookRecord supbook = (SupbookRecord) supbooks.get(0);
      if (supbook.getType() == SupbookRecord.INTERNAL)
      {
        supbook.adjustInternal(sheets.size());
      }
    }

    if (names != null && names.size() > 0)
    {
      for (int i=0; i< names.size();i++)
      {
        NameRecord n = (NameRecord) names.get(i);
        int oldRef = n.getSheetRef();
        if(oldRef == (pos+1))
        {
          n.setSheetRef(0); // make a global name reference
        } 
        else if (oldRef > (pos+1))
        {
          if(oldRef < 1)
          {
            oldRef = 1;
          }
          n.setSheetRef(oldRef-1); // move one sheet
        }
      }
    }
  }

  /**
   * Moves the specified sheet within this workbook to another index
   * position.
   * 
   * @param fromIndex the zero based index of the reQuired sheet
   * @param toIndex the zero based index of the reQuired sheet
   * @return the sheet that has been moved
   */
  public WritableSheet moveSheet(int fromIndex, int toIndex)
  {
    // Handle dodgy index 
    fromIndex = Math.max(fromIndex, 0);
    fromIndex = Math.min(fromIndex, sheets.size() - 1);
    toIndex   = Math.max(toIndex, 0);
    toIndex   = Math.min(toIndex, sheets.size() - 1);

    WritableSheet sheet= (WritableSheet)sheets.remove(fromIndex);
    sheets.add(toIndex, sheet);
    
    return sheet;
  }

  /**
   * Writes out this sheet to the output file.  First it writes out
   * the standard workbook information required by excel, before calling
   * the write method on each sheet individually
   * 
   * @exception IOException 
   */
  public void write() throws IOException
  {
    // Perform some preliminary sheet check before we start writing out
    // the workbook
    WritableSheetImpl wsi = null;
    for (int i = 0; i < getNumberOfSheets(); i++)
    {
      wsi = (WritableSheetImpl) getSheet(i);

      // Check the merged records.  This has to be done before the
      // globals are written out because some more XF formats might be created
      wsi.checkMergedBorders();

      // Check to see if there are any predefined names
      Range range = wsi.getSettings().getPrintArea();
      if (range != null)
      {
        addNameArea(BuiltInName.PRINT_AREA,
                    wsi,
                    range.getTopLeft().getColumn(),
                    range.getTopLeft().getRow(),
                    range.getBottomRight().getColumn(),
                    range.getBottomRight().getRow(),
                    false);
      }
      
      // Check to see if print titles by row were set
      Range rangeR = wsi.getSettings().getPrintTitlesRow();
      Range rangeC = wsi.getSettings().getPrintTitlesCol();      
      if (rangeR != null && rangeC != null)
      {
        addNameArea(BuiltInName.PRINT_TITLES, 
                    wsi,
                    rangeR.getTopLeft().getColumn(),
                    rangeR.getTopLeft().getRow(),
                    rangeR.getBottomRight().getColumn(),
                    rangeR.getBottomRight().getRow(),
                    rangeC.getTopLeft().getColumn(),
                    rangeC.getTopLeft().getRow(),
                    rangeC.getBottomRight().getColumn(),
                    rangeC.getBottomRight().getRow(),
                    false);  
      }
      // Check to see if print titles by row were set
      else if (rangeR != null)
      {
    	  addNameArea(BuiltInName.PRINT_TITLES, 
                    wsi,
                    rangeR.getTopLeft().getColumn(),
                    rangeR.getTopLeft().getRow(),
                    rangeR.getBottomRight().getColumn(),
                    rangeR.getBottomRight().getRow(),
                    false);
      }
      // Check to see if print titles by column were set 
      else if (rangeC != null)
      {
        addNameArea(BuiltInName.PRINT_TITLES, 
                    wsi,
                    rangeC.getTopLeft().getColumn(),
                    rangeC.getTopLeft().getRow(),
                    rangeC.getBottomRight().getColumn(),
                    rangeC.getBottomRight().getRow(), 
                    false);
      }
    }
    
    // Rationalize all the XF and number formats
    if (!settings.getRationalizationDisabled())
    {
      rationalize();
    }

    // Write the workbook globals
    BOFRecord bof = new BOFRecord(BOFRecord.workbookGlobals);
    outputFile.write(bof);

    // Must immediatly follow the BOF record
    if (settings.getTemplate())
    {
      // Only write record if we are a template
      TemplateRecord trec= new TemplateRecord();
      outputFile.write(trec);
    }


    InterfaceHeaderRecord ihr = new InterfaceHeaderRecord();
    outputFile.write(ihr);

    MMSRecord mms = new MMSRecord(0,0);
    outputFile.write(mms);

    InterfaceEndRecord ier = new InterfaceEndRecord();
    outputFile.write(ier);

    WriteAccessRecord wr = new WriteAccessRecord(settings.getWriteAccess());
    outputFile.write(wr);

    CodepageRecord cp = new CodepageRecord();
    outputFile.write(cp);

    DSFRecord dsf = new DSFRecord();
    outputFile.write(dsf);

    if (settings.getExcel9File())
    {
      // Only write record if we are a template
      // We are not excel 2000, should we still set the flag
      Excel9FileRecord e9rec= new Excel9FileRecord();
      outputFile.write(e9rec);
    }

    TabIdRecord tabid = new TabIdRecord(getNumberOfSheets());
    outputFile.write(tabid);

    if (containsMacros)
    {
      ObjProjRecord objproj = new ObjProjRecord();
      outputFile.write(objproj);
    }

    if (buttonPropertySet != null)
    {
      outputFile.write(buttonPropertySet);
    }

    FunctionGroupCountRecord fgcr = new FunctionGroupCountRecord();
    outputFile.write(fgcr);

    // do not support password protected workbooks
    WindowProtectRecord wpr = new WindowProtectRecord
      (settings.getWindowProtected());
    outputFile.write(wpr);

    ProtectRecord pr = new ProtectRecord(wbProtected);
    outputFile.write(pr);

    PasswordRecord pw = new PasswordRecord(null);
    outputFile.write(pw);

    Prot4RevRecord p4r = new Prot4RevRecord(false);
    outputFile.write(p4r);

    Prot4RevPassRecord p4rp = new Prot4RevPassRecord();
    outputFile.write(p4rp);

    // If no sheet is identified as being selected, then select
    // the first one
    boolean sheetSelected = false;
    WritableSheetImpl wsheet = null;
    int selectedSheetIndex = 0;
    for (int i = 0 ; i < getNumberOfSheets() && !sheetSelected ; i++)
    {
      wsheet = (WritableSheetImpl) getSheet(i);
      if (wsheet.getSettings().isSelected())
      {
        sheetSelected = true;
        selectedSheetIndex = i;
      }
    }

    if (!sheetSelected)
    {
      wsheet = (WritableSheetImpl) getSheet(0);
      wsheet.getSettings().setSelected(true);
      selectedSheetIndex = 0;
    }

    Window1Record w1r = new Window1Record(selectedSheetIndex);
    outputFile.write(w1r);

    BackupRecord bkr = new BackupRecord(false);
    outputFile.write(bkr);

    HideobjRecord ho = new HideobjRecord(settings.getHideobj());
    outputFile.write(ho);
    
    NineteenFourRecord nf = new NineteenFourRecord(false);
    outputFile.write(nf);

    PrecisionRecord pc = new PrecisionRecord(false);
    outputFile.write(pc);

    RefreshAllRecord rar = new RefreshAllRecord(settings.getRefreshAll());
    outputFile.write(rar);

    BookboolRecord bb = new BookboolRecord(true);
    outputFile.write(bb);

    // Write out all the fonts used
    fonts.write(outputFile);

    // Write out the cell formats used within this workbook
    formatRecords.write(outputFile);

    // Write out the palette, if it exists
    if (formatRecords.getPalette() != null)
    {
      outputFile.write(formatRecords.getPalette());
    }

    // Write out the uses elfs record
    UsesElfsRecord uer = new UsesElfsRecord();
    outputFile.write(uer);
    
    // Write out the boundsheet records.  Keep a handle to each one's
    // position so we can write in the stream offset later
    int[] boundsheetPos = new int[getNumberOfSheets()];
    Sheet sheet = null;

    for (int i = 0; i < getNumberOfSheets(); i++)
    {
      boundsheetPos[i] = outputFile.getPos();
      sheet = getSheet(i);
      BoundsheetRecord br = new BoundsheetRecord(sheet.getName());
      if (sheet.getSettings().isHidden())
      {
        br.setHidden();
      }

      if ( ( (WritableSheetImpl) sheets.get(i)).isChartOnly())
      {
        br.setChartOnly();
      }

      outputFile.write(br);
    }

    if (countryRecord == null)
    {
      CountryCode lang = 
        CountryCode.getCountryCode(settings.getExcelDisplayLanguage());
      if (lang == CountryCode.UNKNOWN)
      {
        logger.warn("Unknown country code " + 
                    settings.getExcelDisplayLanguage() + 
                    " using " + CountryCode.USA.getCode());
        lang = CountryCode.USA;
      }
      CountryCode region = 
        CountryCode.getCountryCode(settings.getExcelRegionalSettings());
      countryRecord = new CountryRecord(lang, region);
      if (region == CountryCode.UNKNOWN)
      {
        logger.warn("Unknown country code " + 
                    settings.getExcelDisplayLanguage() + 
                    " using " + CountryCode.UK.getCode());
        region = CountryCode.UK;
      }
    }

    outputFile.write(countryRecord);
    
    // Write out the names of any add in functions
    if (addInFunctionNames != null && addInFunctionNames.length > 0)
    {
      // Write out the supbook record
      //      SupbookRecord supbook = new SupbookRecord();
      //      outputFile.write(supbook);

      for (int i = 0 ; i < addInFunctionNames.length; i++)
      {
        ExternalNameRecord enr = new ExternalNameRecord(addInFunctionNames[i]);
        outputFile.write(enr);
      }
    }

    if (xctRecords != null)
    {
      for (int i = 0 ; i < xctRecords.length; i++)
      {
        outputFile.write(xctRecords[i]);
      }
    }

    // Write out the external sheet record, if it exists
    if (externSheet != null)
    {
      //Write out all the supbook records
      for (int i = 0; i < supbooks.size() ; i++)
      {
        SupbookRecord supbook = (SupbookRecord) supbooks.get(i);
        outputFile.write(supbook);
      }
      outputFile.write(externSheet);
    }

    // Write out the names, if any exists
    if (names != null)
    {
      for (int i = 0 ; i < names.size() ; i++)
      {
        NameRecord n = (NameRecord) names.get(i);
        outputFile.write(n);
      }
    }
  
    // Write out the mso drawing group, if it exists
    if (drawingGroup != null)
    {
      drawingGroup.write(outputFile);
    }

    sharedStrings.write(outputFile);

    EOFRecord eof = new EOFRecord();
    outputFile.write(eof);


    // Write out the sheets
    for (int i = 0; i < getNumberOfSheets(); i++)
    {
      // first go back and modify the offset we wrote out for the
      // boundsheet record
      outputFile.setData
        (IntegerHelper.getFourBytes(outputFile.getPos()),
         boundsheetPos[i] + 4);

      wsheet = (WritableSheetImpl) getSheet(i);
      wsheet.write();
    }
  }

  /**
   * Produces a writable copy of the workbook passed in by 
   * creating copies of each sheet in the specified workbook and adding 
   * them to its own record
   * 
   * @param w the workbook to copy
   */
  private void copyWorkbook(Workbook w)
  {
    int numSheets = w.getNumberOfSheets();
    wbProtected = w.isProtected();
    Sheet s = null;
    WritableSheetImpl ws = null;
    for (int i = 0 ; i < numSheets; i++)
    {
      s = w.getSheet(i);
      ws = (WritableSheetImpl) createSheet(s.getName(),i, false);
      ws.copy(s);
    }
  }

  /**
   * Copies the specified sheet and places it at the index
   * specified by the parameter
   *
   * @param s the index of the sheet to copy
   * @param name the name of the new sheet
   * @param index the position of the new sheet
   */
  public void copySheet(int s, String name, int index)
  {
    WritableSheet sheet = getSheet(s);
    WritableSheetImpl ws = (WritableSheetImpl) createSheet(name, index);
    ws.copy(sheet);
  }

  /**
   * Copies the specified sheet and places it at the index
   * specified by the parameter
   *
   * @param s the name of the sheet to copy
   * @param name the name of the new sheet
   * @param index the position of the new sheet
   */
  public void copySheet(String s, String name, int index)
  {
    WritableSheet sheet = getSheet(s);
    WritableSheetImpl ws = (WritableSheetImpl) createSheet(name, index);
    ws.copy(sheet);
  }

  /**
   * Indicates whether or not this workbook is protected
   * 
   * @param prot protected flag
   */
  public void setProtected(boolean prot)
  {
    wbProtected = prot;
  }

  /**
   * Rationalizes the cell formats, and then passes the resultant XF index
   * mappings to each sheet in turn
   */
  private void rationalize()
  {
    IndexMapping fontMapping   = formatRecords.rationalizeFonts();
    IndexMapping formatMapping = formatRecords.rationalizeDisplayFormats();
    IndexMapping xfMapping     = formatRecords.rationalize(fontMapping, 
                                                           formatMapping);

    WritableSheetImpl wsi = null;
    for (int i = 0; i < sheets.size(); i++)
    {
      wsi = (WritableSheetImpl) sheets.get(i);
      wsi.rationalize(xfMapping, fontMapping, formatMapping);
    }
  }

  /**
   * Gets the internal sheet index for a sheet name
   *
   * @param name the sheet name
   * @return the internal sheet index
   */
  private int getInternalSheetIndex(String name)
  {
    int index = -1;
    String[] names = getSheetNames();
    for (int i = 0 ; i < names.length; i++)
    {
      if (name.equals(names[i]))
      {
        index = i;
        break;
      }
    }

    return index;
  }

  /**
   * Gets the name of the external sheet specified by the index
   *
   * @param index the external sheet index
   * @return the name of the external sheet
   */
  public String getExternalSheetName(int index)
  {
    int supbookIndex = externSheet.getSupbookIndex(index);
    SupbookRecord sr = (SupbookRecord) supbooks.get(supbookIndex);
    
    int firstTab = externSheet.getFirstTabIndex(index);
    
    if (sr.getType() == SupbookRecord.INTERNAL)
    {
      // It's an internal reference - get the name from the sheets list
      WritableSheet ws = getSheet(firstTab);

      return ws.getName();
    }
    else if (sr.getType() == SupbookRecord.EXTERNAL)
    {
      String name = sr.getFileName() + sr.getSheetName(firstTab);
      return name;
    }
    
    // An unknown supbook - return unkown
    logger.warn("Unknown Supbook 1");
    return "[UNKNOWN]";
  }

  /**
   * Gets the name of the last external sheet specified by the index
   *
   * @param index the external sheet index
   * @return the name of the external sheet
   */
  public String getLastExternalSheetName(int index)
  {
    int supbookIndex = externSheet.getSupbookIndex(index);
    SupbookRecord sr = (SupbookRecord) supbooks.get(supbookIndex);
    
    int lastTab = externSheet.getLastTabIndex(index);
    
    if (sr.getType() == SupbookRecord.INTERNAL)
    {
      // It's an internal reference - get the name from the sheets list
      WritableSheet ws = getSheet(lastTab);

      return ws.getName();
    }
    else if (sr.getType() == SupbookRecord.EXTERNAL)
    {
      Assert.verify(false);
    }
    
    // An unknown supbook - return unkown
    logger.warn("Unknown Supbook 2");
    return "[UNKNOWN]";
  }

  /**
   * Parsing of formulas is only supported for a subset of the available
   * biff version, so we need to test to see if this version is acceptable
   *
   * @return the BOF record, which 
   */
  public jxl.read.biff.BOFRecord getWorkbookBof()
  {
    return null;
  }


  /**
   * Gets the index of the external sheet for the name
   *
   * @param sheetName
   * @return the sheet index of the external sheet index
   */
  public int getExternalSheetIndex(int index)
  {
    if (externSheet == null)
    {
      return index;
    }

    Assert.verify(externSheet != null);

    int firstTab = externSheet.getFirstTabIndex(index);

    return firstTab;
  }

  /**
   * Gets the index of the external sheet for the name
   *
   * @param sheetName
   * @return the sheet index of the external sheet index
   */
  public int getLastExternalSheetIndex(int index)
  {
    if (externSheet == null)
    {
      return index;
    }

    Assert.verify(externSheet != null);

    int lastTab = externSheet.getLastTabIndex(index);

    return lastTab;
  }

  /**
   * Gets the external sheet index for the sheet name
   *
   * @param sheetName 
   * @return the sheet index or -1 if the sheet could not be found
   */
  public int getExternalSheetIndex(String sheetName)
  {
    if (externSheet == null)
    {
      externSheet = new ExternalSheetRecord();
      supbooks = new ArrayList();
      supbooks.add(new SupbookRecord(getNumberOfSheets(), settings));
    }

    // Iterate through the sheets records
    boolean found = false;
    Iterator i = sheets.iterator();
    int sheetpos = 0;
    WritableSheetImpl s = null;

    while (i.hasNext() && !found)
    {
      s = (WritableSheetImpl) i.next();
      
      if (s.getName().equals(sheetName))
      {
        found = true;
      }
      else
      {
        sheetpos++;
      }
    }

    if (found)
    {
      // Check that the supbook record at position zero is internal and
      // contains all the sheets
      SupbookRecord supbook = (SupbookRecord) supbooks.get(0);
      if (supbook.getType() != SupbookRecord.INTERNAL ||
          supbook.getNumberOfSheets() != getNumberOfSheets())
      {
        logger.warn("Cannot find sheet " + sheetName + " in supbook record");
      }
      
      return externSheet.getIndex(0, sheetpos);
    }

    // Check for square brackets
    int closeSquareBracketsIndex = sheetName.lastIndexOf(']');
    int openSquareBracketsIndex = sheetName.lastIndexOf('[');
    
    if (closeSquareBracketsIndex == -1 ||
        openSquareBracketsIndex == -1)
    {
      logger.warn("Square brackets");
      return -1;
    }

    String worksheetName = sheetName.substring(closeSquareBracketsIndex+1);
    String workbookName = sheetName.substring(openSquareBracketsIndex+1, 
                                              closeSquareBracketsIndex);
    String path = sheetName.substring(0, openSquareBracketsIndex);
    String fileName = path + workbookName;

    boolean supbookFound = false;
    SupbookRecord externalSupbook = null;
    int supbookIndex = -1;
    for (int ind = 0; ind < supbooks.size() && !supbookFound ; ind++)
    {
      externalSupbook = (SupbookRecord) supbooks.get(ind);
      if (externalSupbook.getType() == SupbookRecord.EXTERNAL &&
          externalSupbook.getFileName().equals(fileName))
      {
        supbookFound = true;
        supbookIndex = ind;
      }
    }

    if (!supbookFound)
    {
      externalSupbook = new SupbookRecord(fileName, settings);
      supbookIndex = supbooks.size();
      supbooks.add(externalSupbook);
    }
    
    int sheetIndex = externalSupbook.getSheetIndex(worksheetName);

    return externSheet.getIndex(supbookIndex, sheetIndex);
  }

  /**
   * Gets the last external sheet index for the sheet name
   * @param sheetName 
   * @return the sheet index or -1 if the sheet could not be found
   */
  public int getLastExternalSheetIndex(String sheetName)
  {
    if (externSheet == null)
    {
      externSheet = new ExternalSheetRecord();
      supbooks = new ArrayList();
      supbooks.add(new SupbookRecord(getNumberOfSheets(), settings));
    }

    // Iterate through the sheets records
    boolean found = false;
    Iterator i = sheets.iterator();
    int sheetpos = 0;
    WritableSheetImpl s = null;

    while (i.hasNext() && !found)
    {
      s = (WritableSheetImpl) i.next();
      
      if (s.getName().equals(sheetName))
      {
        found = true;
      }
      else
      {
        sheetpos++;
      }
    }

    if (!found)
    {
      return -1;
    }

    // Check that the supbook record at position zero is internal and contains
    // all the sheets
    SupbookRecord supbook = (SupbookRecord) supbooks.get(0);
    Assert.verify(supbook.getType() == SupbookRecord.INTERNAL &&
                  supbook.getNumberOfSheets() == getNumberOfSheets());

    return externSheet.getIndex(0, sheetpos);
  }

  /**
   * Sets the RGB value for the specified colour for this workbook
   *
   * @param c the colour whose RGB value is to be overwritten
   * @param r the red portion to set (0-255)
   * @param g the green portion to set (0-255)
   * @param b the blue portion to set (0-255)
   */
  public void setColourRGB(Colour c, int r, int g, int b)
  {
    formatRecords.setColourRGB(c,r,g,b);
  }

  /**
   * Accessor for the RGB value for the specified colour
   *
   * @return the RGB for the specified colour
   */
  public RGB getColourRGB(Colour c)
  {
    return formatRecords.getColourRGB(c);
  }

  /**
   * Gets the name at the specified index
   *
   * @param index the index into the name table
   * @return the name of the cell
   */
  public String getName(int index)
  {
    Assert.verify(index >= 0 && index < names.size());
    NameRecord n = (NameRecord) names.get(index);
    return n.getName();
  }

  /**
   * Gets the index of the name record for the name
   *
   * @param name 
   * @return the index in the name table
   */
  public int getNameIndex(String name)
  {
    NameRecord nr = (NameRecord) nameRecords.get(name);
    return nr != null ? nr.getIndex() : -1;
  }

  /**
   * Adds a cell to workbook wide range of cells which need adjustment
   * following a row/column insert or remove
   *
   * @param f the cell to add to the list
   */
  void addRCIRCell(CellValue cv)
  {
    rcirCells.add(cv);
  }

  /**
   * Called when a column is inserted on the specified sheet.  Notifies all
   * RCIR cells of this change
   *
   * @param s the sheet on which the column was inserted
   * @param col the column number which was inserted
   */
  void columnInserted(WritableSheetImpl s, int col)
  {
    int externalSheetIndex = getExternalSheetIndex(s.getName());
    for (Iterator i = rcirCells.iterator() ; i.hasNext() ;)
    {
      CellValue cv = (CellValue) i.next();
      cv.columnInserted(s, externalSheetIndex, col);
    }

    // Adjust any named cells
    if (names != null)
    {
      for (Iterator i = names.iterator(); i.hasNext() ;)
      {
        NameRecord nameRecord = (NameRecord) i.next();
        nameRecord.columnInserted(externalSheetIndex, col);
      }
    }
  }

  /**
   * Called when a column is removed on the specified sheet.  Notifies all
   * RCIR cells of this change
   *
   * @param s the sheet on which the column was removed
   * @param col the column number which was removed
   */
  void columnRemoved(WritableSheetImpl s, int col)
  {
    int externalSheetIndex = getExternalSheetIndex(s.getName());
    for (Iterator i = rcirCells.iterator() ; i.hasNext() ;)
    {
      CellValue cv = (CellValue) i.next();
      cv.columnRemoved(s, externalSheetIndex, col);
    }

    // Adjust any named cells
    ArrayList removedNames = new ArrayList();
    if (names != null)
    {
      for (Iterator i = names.iterator(); i.hasNext() ;)
      {
        NameRecord nameRecord = (NameRecord) i.next();
        boolean removeName = nameRecord.columnRemoved(externalSheetIndex,
                                                      col);

        if (removeName)
        {
          removedNames.add(nameRecord);
        }
      }

      // Remove any names which have been deleted
      for (Iterator i = removedNames.iterator(); i.hasNext() ;)
      {
        NameRecord nameRecord = (NameRecord) i.next();
        boolean removed = names.remove(nameRecord);
        Assert.verify(removed, "Could not remove name " + 
                      nameRecord.getName());
      }
    }
  }

  /**
   * Called when a row is inserted on the specified sheet.  Notifies all
   * RCIR cells of this change
   *
   * @param s the sheet on which the row was inserted
   * @param row the row number which was inserted
   */
  void rowInserted(WritableSheetImpl s, int row)
  {
    int externalSheetIndex = getExternalSheetIndex(s.getName());
    
    // Adjust the row infos
    for (Iterator i = rcirCells.iterator() ; i.hasNext() ;)
    {
      CellValue cv = (CellValue) i.next();
      cv.rowInserted(s, externalSheetIndex, row);
    }

    // Adjust any named cells
    if (names != null)
    {
      for (Iterator i = names.iterator(); i.hasNext() ;)
      {
        NameRecord nameRecord = (NameRecord) i.next();
        nameRecord.rowInserted(externalSheetIndex, row);
      }
    }
  }

  /**
   * Called when a row is removed on the specified sheet.  Notifies all
   * RCIR cells of this change
   *
   * @param s the sheet on which the row was removed
   * @param row the row number which was removed
   */
  void rowRemoved(WritableSheetImpl s, int row)
  {
    int externalSheetIndex = getExternalSheetIndex(s.getName());
    for (Iterator i = rcirCells.iterator() ; i.hasNext() ;)
    {
      CellValue cv = (CellValue) i.next();
      cv.rowRemoved(s, externalSheetIndex, row);
    }

    // Adjust any named cells
    ArrayList removedNames = new ArrayList();
    if (names != null)
    {
      for (Iterator i = names.iterator(); i.hasNext() ;)
      {
        NameRecord nameRecord = (NameRecord) i.next();
        boolean removeName = nameRecord.rowRemoved(externalSheetIndex, row);

        if (removeName)
        {
          removedNames.add(nameRecord);
        }
      }

      // Remove any names which have been deleted
      for (Iterator i = removedNames.iterator(); i.hasNext() ;)
      {
        NameRecord nameRecord = (NameRecord) i.next();
        boolean removed = names.remove(nameRecord);
        Assert.verify(removed, "Could not remove name " + 
                      nameRecord.getName());
      }
    }
  }

  /**
   * Gets the named cell from this workbook.  If the name refers to a
   * range of cells, then the cell on the top left is returned.  If
   * the name cannot be found, null is returned
   *
   * @param  the name of the cell/range to search for
   * @return the cell in the top left of the range if found, NULL
   *         otherwise
   */
  public WritableCell findCellByName(String name)
  {
    NameRecord nr = (NameRecord) nameRecords.get(name);

    if (nr == null)
    {
      return null;
    }

    NameRecord.NameRange[] ranges = nr.getRanges();

    // Go and retrieve the first cell in the first range
    int sheetIndex = getExternalSheetIndex(ranges[0].getExternalSheet());
    WritableSheet s    = getSheet(sheetIndex);
    WritableCell  cell = s.getWritableCell(ranges[0].getFirstColumn(), 
                                           ranges[0].getFirstRow());

    return cell;
  }

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
   * @param  the name of the cell/range to search for
   * @return the range of cells
   */
  public Range[] findByName(String name)
  {
    NameRecord nr = (NameRecord) nameRecords.get(name);

    if (nr == null)
    {
      return null;
    }

    NameRecord.NameRange[] ranges = nr.getRanges();

    Range[] cellRanges = new Range[ranges.length];

    for (int i = 0; i < ranges.length ; i++)
    {
      cellRanges[i] = new RangeImpl
        (this, 
         getExternalSheetIndex(ranges[i].getExternalSheet()),
         ranges[i].getFirstColumn(),
         ranges[i].getFirstRow(),
         getLastExternalSheetIndex(ranges[i].getExternalSheet()),
         ranges[i].getLastColumn(),
         ranges[i].getLastRow());
    }

    return cellRanges;
  }

  /**
   * Adds a drawing to this workbook
   *
   * @param d the drawing to add
   */
  void addDrawing(DrawingGroupObject d)
  {
    if (drawingGroup == null)
    {
      drawingGroup = new DrawingGroup(Origin.WRITE);
    }

    drawingGroup.add(d);
  }

  /**
   * Removes a drawing from this workbook
   *
   * @param d the drawing to remove
   */
  void removeDrawing(Drawing d)
  {
    Assert.verify(drawingGroup != null);

    drawingGroup.remove(d);
  }

  /**
   * Accessor for the drawing group
   *
   * @return the drawing group
   */
  DrawingGroup getDrawingGroup()
  {
    return drawingGroup;
  }

  /**
   * Create a drawing group for this workbook - used when importing sheets
   * which contain drawings, but this workbook doesn't.
   * We can't subsume this into the getDrawingGroup() method because the
   * null-ness of the return value is used elsewhere to determine the
   * origin of the workbook
   */
  DrawingGroup createDrawingGroup()
  {
    if (drawingGroup == null)
    {
      drawingGroup = new DrawingGroup(Origin.WRITE);
    }

    return drawingGroup;
  }

  /**
   * Gets the named ranges
   *
   * @return the list of named cells within the workbook
   */
  public String[] getRangeNames()
  {
    if (names == null)
    {
      return new String[0];
    }

    String[] n = new String[names.size()];
    for (int i = 0 ; i < names.size() ; i++)
    {
      NameRecord nr = (NameRecord) names.get(i);
      n[i] = nr.getName();
    }

    return n;
  }

  /**
   * Removes the specified named range from the workbook
   *
   * @param name the name to remove
   */
  public void removeRangeName(String name)
  {
    int pos = 0;
    boolean found = false;
    for (Iterator i = names.iterator(); i.hasNext() && !found ;)
    {
      NameRecord nr = (NameRecord) i.next();
      if (nr.getName().equals(name))
      {
        found = true;
      }
      else
      {
        pos++;
      }
    }

    // Remove the name from the list of names and the associated hashmap
    // of names (used to retrieve the name index).  If the name cannot
    // be found, a warning is displayed
    if (found)
    {
      names.remove(pos);
      if (nameRecords.remove(name) == null)
      {
        logger.warn("Could not remove " + name + " from index lookups");
      }
    }
  }

  /**
   * Accessor for the jxl.common.styles
   *
   * @return the standard styles for this workbook
   */
  Styles getStyles()
  {
    return styles;
  }

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
  public void addNameArea(String name, 
                          WritableSheet sheet,
                          int firstCol, 
                          int firstRow,  
                          int lastCol,
                          int lastRow)
  {
    addNameArea(name, sheet, firstCol, firstRow, lastCol, lastRow, true);
  }

  /**
   * Add new named area to this workbook with the given information.
   * 
   * @param name name to be created.
   * @param sheet sheet containing the name
   * @param firstCol  first column this name refers to.
   * @param firstRow  first row this name refers to.
   * @param lastCol    last column this name refers to.
   * @param lastRow    last row this name refers to.
   * @param global   TRUE if this is a global name, FALSE if this is tied to 
   *                 the sheet
   */
  void addNameArea(String name, 
                   WritableSheet sheet,
                   int firstCol, 
                   int firstRow,  
                   int lastCol,
                   int lastRow,
                   boolean global)
  {
    if (names == null)
    {
      names = new ArrayList();
    }

    int externalSheetIndex = getExternalSheetIndex(sheet.getName());

    // Create a new name record.
    NameRecord nr = 
      new NameRecord(name, 
                     names.size(),
                     externalSheetIndex, 
                     firstRow, lastRow, 
                     firstCol, lastCol, 
                     global);
    
    // Add new name to name array.
    names.add(nr);
    
    // Add new name to name hash table.
    nameRecords.put(name, nr);
  }

  /**
   * Add new named area to this workbook with the given information.
   * 
   * @param name name to be created.
   * @param sheet sheet containing the name
   * @param firstCol  first column this name refers to.
   * @param firstRow  first row this name refers to.
   * @param lastCol    last column this name refers to.
   * @param lastRow    last row this name refers to.
   * @param global   TRUE if this is a global name, FALSE if this is tied to 
   *                 the sheet
   */
  void addNameArea(BuiltInName name, 
                   WritableSheet sheet,
                   int firstCol, 
                   int firstRow,  
                   int lastCol,
                   int lastRow,
                   boolean global)
  {
    if (names == null)
    {
      names = new ArrayList();
    }

    int index = getInternalSheetIndex(sheet.getName());
    int externalSheetIndex = getExternalSheetIndex(sheet.getName());

    // Create a new name record.
    NameRecord nr = 
      new NameRecord(name, 
                     index,
                     externalSheetIndex, 
                     firstRow, lastRow, 
                     firstCol, lastCol, 
                     global);
    
    // Add new name to name array.
    names.add(nr);
    
    // Add new name to name hash table.
    nameRecords.put(name, nr);
  }

  /**
   * Add new named area to this workbook with the given information.
   * 
   * @param name name to be created.
   * @param sheet sheet containing the name
   * @param firstCol  first column this name refers to.
   * @param firstRow  first row this name refers to.
   * @param lastCol   last column this name refers to.
   * @param lastRow   last row this name refers to.
   * @param firstCol2 first column this name refers to.
   * @param firstRow2 first row this name refers to.
   * @param lastCol2  last column this name refers to.
   * @param lastRow2  last row this name refers to.
   * @param global   TRUE if this is a global name, FALSE if this is tied to 
   *                 the sheet
   */
  void addNameArea(BuiltInName name, 
                   WritableSheet sheet,
                   int firstCol, 
                   int firstRow,  
                   int lastCol,
                   int lastRow,
                   int firstCol2, 
                   int firstRow2,  
                   int lastCol2,
                   int lastRow2,
                   boolean global)
  {
    if (names == null)
    {
      names = new ArrayList();
    }

    int index = getInternalSheetIndex(sheet.getName());
    int externalSheetIndex = getExternalSheetIndex(sheet.getName());

    // Create a new name record.
    NameRecord nr = 
      new NameRecord(name, 
                     index,
                     externalSheetIndex, 
                     firstRow2, lastRow2, 
                     firstCol2, lastCol2,
                     firstRow, lastRow, 
                     firstCol, lastCol, 
                     global);
    
    // Add new name to name array.
    names.add(nr);
    
    // Add new name to name hash table.
    nameRecords.put(name, nr);
  }  
  
  /**
   * Accessor for the workbook settings
   */
  WorkbookSettings getSettings()
  {
    return settings;
  }

  /**
   * Returns the cell for the specified location eg. "Sheet1!A4".  
   * This is identical to using the CellReferenceHelper with its
   * associated performance overheads, consequently it should
   * be use sparingly
   *
   * @param loc the cell to retrieve
   * @return the cell at the specified location
   */
  public WritableCell getWritableCell(String loc)
  {
    WritableSheet s = getSheet(CellReferenceHelper.getSheet(loc)); 
    return s.getWritableCell(loc);
  }

  /**
   * Imports a sheet from a different workbook.  Does a deep copy on all
   * elements within that sheet
   *
   * @param name the name of the new sheet
   * @param index the position for the new sheet within this workbook
   * @param sheet the sheet (from another workbook) to merge into this one
   * @return the new sheet
   */
  public WritableSheet importSheet(String name, int index, Sheet sheet)
  {
    WritableSheet ws = createSheet(name, index);
    ((WritableSheetImpl) ws).importSheet(sheet);

    return ws;
  }

}





