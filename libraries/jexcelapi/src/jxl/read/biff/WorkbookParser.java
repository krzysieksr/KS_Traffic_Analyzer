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
import java.util.HashMap;
import java.util.Iterator;

import jxl.common.Assert;
import jxl.common.Logger;

import jxl.Cell;
import jxl.Range;
import jxl.Sheet;
import jxl.Workbook;
import jxl.WorkbookSettings;
import jxl.biff.BuiltInName;
import jxl.biff.CellReferenceHelper;
import jxl.biff.EmptyCell;
import jxl.biff.FontRecord;
import jxl.biff.Fonts;
import jxl.biff.FormatRecord;
import jxl.biff.FormattingRecords;
import jxl.biff.NameRangeException;
import jxl.biff.NumFormatRecordsException;
import jxl.biff.PaletteRecord;
import jxl.biff.RangeImpl;
import jxl.biff.StringHelper;
import jxl.biff.Type;
import jxl.biff.WorkbookMethods;
import jxl.biff.XCTRecord;
import jxl.biff.XFRecord;
import jxl.biff.drawing.DrawingGroup;
import jxl.biff.drawing.MsoDrawingGroupRecord;
import jxl.biff.drawing.Origin;
import jxl.biff.formula.ExternalSheet;

/**
 * Parses the biff file passed in, and builds up an internal representation of
 * the spreadsheet
 */
public class WorkbookParser extends Workbook
  implements ExternalSheet, WorkbookMethods
{
  /**
   * The logger
   */
  private static Logger logger = Logger.getLogger(WorkbookParser.class);

  /**
   * The excel file
   */
  private File excelFile;
  /**
   * The number of open bofs
   */
  private int bofs;
  /**
   * Indicates whether or not the dates are based around the 1904 date system
   */
  private boolean nineteenFour;
  /**
   * The shared string table
   */
  private SSTRecord sharedStrings;
  /**
   * The names of all the worksheets
   */
  private ArrayList boundsheets;
  /**
   * The xf records
   */
  private FormattingRecords formattingRecords;
  /**
   * The fonts used by this workbook
   */
  private Fonts fonts;

  /**
   * The sheets contained in this workbook
   */
  private ArrayList sheets;

  /**
   * The last sheet accessed
   */
  private SheetImpl lastSheet;

  /**
   * The index of the last sheet retrieved
   */
  private int lastSheetIndex;

  /**
   * The named records found in this workbook
   */
  private HashMap namedRecords;

  /**
   * The list of named records
   */
  private ArrayList nameTable;

  /**
   * The list of add in functions
   */
  private ArrayList addInFunctions;

  /**
   * The external sheet record.  Used by formulas, and names
   */
  private ExternalSheetRecord externSheet;

  /**
   * The list of supporting workbooks - used by formulas
   */
  private ArrayList supbooks;

  /**
   * The bof record for this workbook
   */
  private BOFRecord workbookBof;

  /**
   * The Mso Drawing Group record for this workbook
   */
  private MsoDrawingGroupRecord msoDrawingGroup;

  /**
   * The property set record associated with this workbook
   */
  private ButtonPropertySetRecord buttonPropertySet;

  /**
   * Workbook protected flag
   */
  private boolean wbProtected;

  /**
   * Contains macros flag
   */
  private boolean containsMacros;

  /**
   * The workbook settings
   */
  private WorkbookSettings settings;

  /**
   * The drawings contained in this workbook
   */
  private DrawingGroup drawingGroup;

  /**
   * The country record (containing the language and regional settings)
   * for this workbook
   */
  private CountryRecord countryRecord;

  private ArrayList xctRecords;

  /**
   * Constructs this object from the raw excel data
   *
   * @param f the excel 97 biff file
   * @param s the workbook settings
   */
  public WorkbookParser(File f, WorkbookSettings s)
  {
    super();
    excelFile = f;
    boundsheets = new ArrayList(10);
    fonts = new Fonts();
    formattingRecords = new FormattingRecords(fonts);
    sheets = new ArrayList(10);
    supbooks = new ArrayList(10);
    namedRecords = new HashMap();
    lastSheetIndex = -1;
    wbProtected = false;
    containsMacros = false;
    settings = s;
    xctRecords = new ArrayList(10);
  }

 /**
   * Gets the sheets within this workbook.
   * NOTE:  Use of this method for
   * very large worksheets can cause performance and out of memory problems.
   * Use the alternative method getSheet() to retrieve each sheet individually
   *
   * @return an array of the individual sheets
   */
  public Sheet[] getSheets()
  {
    Sheet[] sheetArray = new Sheet[getNumberOfSheets()];
    return (Sheet[]) sheets.toArray(sheetArray);
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
   * @param index the zero based index of the required sheet
   * @return The sheet specified by the index
   */
  public Sheet getSheet(int index)
  {
    // First see if the last sheet index is the same as this sheet index.
    // If so, then the same sheet is being re-requested, so simply
    // return it instead of rereading it
    if ((lastSheet != null) && lastSheetIndex == index)
    {
      return lastSheet;
    }

    // Flush out all of the cached data in the last sheet
    if (lastSheet != null)
    {
      lastSheet.clear();

      if (!settings.getGCDisabled())
      {
        System.gc();
      }
    }

    lastSheet = (SheetImpl) sheets.get(index);
    lastSheetIndex = index;
    lastSheet.readSheet();

    return lastSheet;
  }

  /**
   * Gets the sheet with the specified name from within this workbook
   *
   * @param name the sheet name
   * @return The sheet with the specified name, or null if it is not found
   */
  public Sheet getSheet(String name)
  {
    // Iterate through the boundsheet records
    int pos = 0;
    boolean found = false;
    Iterator i = boundsheets.iterator();
    BoundsheetRecord br = null;

    while (i.hasNext() && !found)
    {
      br = (BoundsheetRecord) i.next();

      if (br.getName().equals(name))
      {
        found = true;
      }
      else
      {
        pos++;
      }
    }

    return found ? getSheet(pos) : null;
  }

  /**
   * Gets the sheet names
   *
   * @return an array of strings containing the sheet names
   */
  public String[] getSheetNames()
  {
    String[] names = new String[boundsheets.size()];

    BoundsheetRecord br = null;
    for (int i = 0; i < names.length; i++)
    {
      br = (BoundsheetRecord) boundsheets.get(i);
      names[i] = br.getName();
    }

    return names;
  }


  /**
   * Package protected function which gets the real internal sheet index
   * based upon  the external sheet reference.  This is used for extern sheet
   * references  which are specified in formulas
   *
   * @param index the external sheet reference
   * @return the actual sheet index
   */
  public int getExternalSheetIndex(int index)
  {
    // For biff7, the whole external reference thing works differently
    // Hopefully for our purposes sheet references will all be local
    if (workbookBof.isBiff7())
    {
      return index;
    }

    Assert.verify(externSheet != null);

    int firstTab = externSheet.getFirstTabIndex(index);

    return firstTab;
  }

  /**
   * Package protected function which gets the real internal sheet index
   * based upon  the external sheet reference.  This is used for extern sheet
   * references  which are specified in formulas
   *
   * @param index the external sheet reference
   * @return the actual sheet index
   */
  public int getLastExternalSheetIndex(int index)
  {
    // For biff7, the whole external reference thing works differently
    // Hopefully for our purposes sheet references will all be local
    if (workbookBof.isBiff7())
    {
      return index;
    }

    Assert.verify(externSheet != null);

    int lastTab = externSheet.getLastTabIndex(index);

    return lastTab;
  }

  /**
   * Gets the name of the external sheet specified by the index
   *
   * @param index the external sheet index
   * @return the name of the external sheet
   */
  public String getExternalSheetName(int index)
  {
    // For biff7, the whole external reference thing works differently
    // Hopefully for our purposes sheet references will all be local
    if (workbookBof.isBiff7())
    {
      BoundsheetRecord br = (BoundsheetRecord) boundsheets.get(index);

      return br.getName();
    }

    int supbookIndex = externSheet.getSupbookIndex(index);
    SupbookRecord sr = (SupbookRecord) supbooks.get(supbookIndex);

    int firstTab = externSheet.getFirstTabIndex(index);
    int lastTab  = externSheet.getLastTabIndex(index);
    String firstTabName = "";
    String lastTabName = "";

    if (sr.getType() == SupbookRecord.INTERNAL)
    {
      // It's an internal reference - get the name from the boundsheets list
      if (firstTab == 65535)
      {
         firstTabName = "#REF";
      }
      else
      {
        BoundsheetRecord br = (BoundsheetRecord) boundsheets.get(firstTab);
        firstTabName = br.getName();
      }
 
      if (lastTab==65535)
      {
        lastTabName = "#REF";
      }
      else
      {
        BoundsheetRecord br = (BoundsheetRecord) boundsheets.get(lastTab);
        lastTabName = br.getName();
      }

      String sheetName = (firstTab == lastTab) ? firstTabName : 
        firstTabName + ':' + lastTabName;

      // if the sheet name contains apostrophes then escape them
      sheetName = sheetName.indexOf('\'') == -1 ? sheetName :
        StringHelper.replace(sheetName, "\'", "\'\'");

      
      // if the sheet name contains spaces, then enclose in quotes
      return sheetName.indexOf(' ') == -1 ? sheetName : 
        '\'' + sheetName + '\'';
    }
    else if (sr.getType() == SupbookRecord.EXTERNAL)
    {
      // External reference - get the sheet name from the supbook record
      StringBuffer sb = new StringBuffer();
      java.io.File fl = new java.io.File(sr.getFileName());
      sb.append("'");
      sb.append(fl.getAbsolutePath());
      sb.append("[");
      sb.append(fl.getName());
      sb.append("]");
      sb.append((firstTab == 65535) ? "#REF" : sr.getSheetName(firstTab));
      if (lastTab != firstTab)
      {
        sb.append(sr.getSheetName(lastTab));
      }
      sb.append("'");
      return sb.toString();
    }

    // An unknown supbook - return unkown
    logger.warn("Unknown Supbook 3");
    return "[UNKNOWN]";
  }

  /**
   * Gets the name of the external sheet specified by the index
   *
   * @param index the external sheet index
   * @return the name of the external sheet
   */
  public String getLastExternalSheetName(int index)
  {
    // For biff7, the whole external reference thing works differently
    // Hopefully for our purposes sheet references will all be local
    if (workbookBof.isBiff7())
    {
      BoundsheetRecord br = (BoundsheetRecord) boundsheets.get(index);

      return br.getName();
    }

    int supbookIndex = externSheet.getSupbookIndex(index);
    SupbookRecord sr = (SupbookRecord) supbooks.get(supbookIndex);

    int lastTab = externSheet.getLastTabIndex(index);

    if (sr.getType() == SupbookRecord.INTERNAL)
    {
      // It's an internal reference - get the name from the boundsheets list
       if (lastTab == 65535)
       {
         return "#REF";
       }
       else
       {
         BoundsheetRecord br = (BoundsheetRecord) boundsheets.get(lastTab);
         return br.getName();
       }
    }
    else if (sr.getType() == SupbookRecord.EXTERNAL)
    {
      // External reference - get the sheet name from the supbook record
      StringBuffer sb = new StringBuffer();
      java.io.File fl = new java.io.File(sr.getFileName());
      sb.append("'");
      sb.append(fl.getAbsolutePath());
      sb.append("[");
      sb.append(fl.getName());
      sb.append("]");
      sb.append((lastTab == 65535) ? "#REF" : sr.getSheetName(lastTab));
      sb.append("'");
      return sb.toString();
    }

    // An unknown supbook - return unkown
    logger.warn("Unknown Supbook 4");
    return "[UNKNOWN]";
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
   */
  public void close()
  {
    if (lastSheet != null)
    {
      lastSheet.clear();
    }
    excelFile.clear();

    if (!settings.getGCDisabled())
    {
      System.gc();
    }
  }

  /**
   * Adds the sheet to the end of the array
   *
   * @param s the sheet to add
   */
  final void addSheet(Sheet s)
  {
    sheets.add(s);
  }

  /**
   * Does the hard work of building up the object graph from the excel bytes
   *
   * @exception BiffException
   * @exception PasswordException if the workbook is password protected
   */
  protected void parse() throws BiffException, PasswordException
  {
    Record r = null;

    BOFRecord bof = new BOFRecord(excelFile.next());
    workbookBof = bof;
    bofs++;

    if (!bof.isBiff8() && !bof.isBiff7())
    {
      throw new BiffException(BiffException.unrecognizedBiffVersion);
    }

    if (!bof.isWorkbookGlobals())
    {
      throw new BiffException(BiffException.expectedGlobals);
    }
    ArrayList continueRecords = new ArrayList();
    ArrayList localNames = new ArrayList();
    nameTable = new ArrayList();
    addInFunctions = new ArrayList();

    // Skip to the first worksheet
    while (bofs == 1)
    {
      r = excelFile.next();

      if (r.getType() == Type.SST)
      {
        continueRecords.clear();
        Record nextrec = excelFile.peek();
        while (nextrec.getType() == Type.CONTINUE)
        {
          continueRecords.add(excelFile.next());
          nextrec = excelFile.peek();
        }

        // cast the array
        Record[] records = new Record[continueRecords.size()];
        records = (Record[]) continueRecords.toArray(records);

        sharedStrings = new SSTRecord(r, records, settings);
      }
      else if (r.getType() == Type.FILEPASS)
      {
        throw new PasswordException();
      }
      else if (r.getType() == Type.NAME)
      {
        NameRecord nr = null;

        if (bof.isBiff8())
        {
          nr = new NameRecord(r, settings, nameTable.size());
        
        }
        else
        {
          nr = new NameRecord(r, settings, nameTable.size(),
                              NameRecord.biff7);
        }

        // Add all local and global names to the name table in order to
        // preserve the indexing
        nameTable.add(nr);

        if (nr.isGlobal())
        {
          namedRecords.put(nr.getName(), nr);
        }
        else
        {
          localNames.add(nr);
        }
      }
      else if (r.getType() == Type.FONT)
      {
        FontRecord fr = null;

        if (bof.isBiff8())
        {
          fr = new FontRecord(r, settings);
        }
        else
        {
          fr = new FontRecord(r, settings, FontRecord.biff7);
        }
        fonts.addFont(fr);
      }
      else if (r.getType() == Type.PALETTE)
      {
        PaletteRecord palette = new PaletteRecord(r);
        formattingRecords.setPalette(palette);
      }
      else if (r.getType() == Type.NINETEENFOUR)
      {
        NineteenFourRecord nr = new NineteenFourRecord(r);
        nineteenFour = nr.is1904();
      }
      else if (r.getType() == Type.FORMAT)
      {
        FormatRecord fr = null;
        if (bof.isBiff8())
        {
          fr = new FormatRecord(r, settings, FormatRecord.biff8);
        }
        else
        {
          fr = new FormatRecord(r, settings, FormatRecord.biff7);
        }
        try
        {
          formattingRecords.addFormat(fr);
        }
        catch (NumFormatRecordsException e)
        {
          // This should not happen.  Bomb out
          Assert.verify(false, e.getMessage());
        }
      }
      else if (r.getType() == Type.XF)
      {
        XFRecord xfr = null;
        if (bof.isBiff8())
        {
          xfr = new XFRecord(r, settings, XFRecord.biff8);
        }
        else
        {
          xfr = new XFRecord(r, settings, XFRecord.biff7);
        }

        try
        {
          formattingRecords.addStyle(xfr);
        }
        catch (NumFormatRecordsException e)
        {
          // This should not happen.  Bomb out
          Assert.verify(false, e.getMessage());
        }
      }
      else if (r.getType() == Type.BOUNDSHEET)
      {
        BoundsheetRecord br = null;

        if (bof.isBiff8())
        {
          br = new BoundsheetRecord(r, settings);
        }
        else
        {
          br = new BoundsheetRecord(r, BoundsheetRecord.biff7);
        }

        if (br.isSheet())
        {
          boundsheets.add(br);
        }
        else if (br.isChart() && !settings.getDrawingsDisabled())
        {
          boundsheets.add(br);
        }
      }
      else if (r.getType() == Type.EXTERNSHEET)
      {
        if (bof.isBiff8())
        {
          externSheet = new ExternalSheetRecord(r, settings);
        }
        else
        {
          externSheet = new ExternalSheetRecord(r, settings,
                                                ExternalSheetRecord.biff7);
        }
      }
      else if (r.getType() == Type.XCT)
      {
        XCTRecord xctr = new XCTRecord(r);
        xctRecords.add(xctr);
      }
      else if (r.getType() == Type.CODEPAGE)
      {
        CodepageRecord cr = new CodepageRecord(r);
        settings.setCharacterSet(cr.getCharacterSet());
      }
      else if (r.getType() == Type.SUPBOOK)
      {
        Record nextrec = excelFile.peek();
        while (nextrec.getType() == Type.CONTINUE)
        {
          r.addContinueRecord(excelFile.next());
          nextrec = excelFile.peek();
        }

        SupbookRecord sr = new SupbookRecord(r, settings);
        supbooks.add(sr);
      }
      else if (r.getType() == Type.EXTERNNAME)
      {
        ExternalNameRecord enr = new ExternalNameRecord(r, settings);

        if (enr.isAddInFunction())
        {
          addInFunctions.add(enr.getName());
        }
      }
      else if (r.getType() == Type.PROTECT)
      {
        ProtectRecord pr = new ProtectRecord(r);
        wbProtected = pr.isProtected();
      }
      else if (r.getType() == Type.OBJPROJ)
      {
        containsMacros = true;
      }
      else if (r.getType() == Type.COUNTRY)
      {
        countryRecord = new CountryRecord(r);
      }
      else if (r.getType() == Type.MSODRAWINGGROUP)
      {
        if (!settings.getDrawingsDisabled())
        {
          msoDrawingGroup = new MsoDrawingGroupRecord(r);

          if (drawingGroup == null)
          {
            drawingGroup = new DrawingGroup(Origin.READ);
          }

          drawingGroup.add(msoDrawingGroup);

          Record nextrec = excelFile.peek();
          while (nextrec.getType() == Type.CONTINUE)
          {
            drawingGroup.add(excelFile.next());
            nextrec = excelFile.peek();
          }
        }
      }
      else if (r.getType() == Type.BUTTONPROPERTYSET)
      {
        buttonPropertySet = new ButtonPropertySetRecord(r);
      }
      else if (r.getType() == Type.EOF)
      {
        bofs--;
      }
      else if (r.getType() == Type.REFRESHALL)
      {
        RefreshAllRecord rfm = new RefreshAllRecord(r);
        settings.setRefreshAll(rfm.getRefreshAll());
      }
      else if (r.getType() == Type.TEMPLATE)
      {
        TemplateRecord rfm = new TemplateRecord(r);
        settings.setTemplate(rfm.getTemplate());
      }
      else if (r.getType() == Type.EXCEL9FILE)
      {
        Excel9FileRecord e9f = new Excel9FileRecord(r);
        settings.setExcel9File(e9f.getExcel9File());
      }
      else if (r.getType() == Type.WINDOWPROTECT)
      {
        WindowProtectedRecord winp = new WindowProtectedRecord(r);
        settings.setWindowProtected(winp.getWindowProtected());
      }
      else if (r.getType() == Type.HIDEOBJ)
      {
        HideobjRecord hobj = new HideobjRecord(r);
        settings.setHideobj(hobj.getHideMode());
      }
      else if (r.getType() == Type.WRITEACCESS)
      {
        WriteAccessRecord war = new WriteAccessRecord(r, bof.isBiff8(),
                                                      settings);
        settings.setWriteAccess(war.getWriteAccess());
      }
      else
      {
        // logger.info("Unsupported record type: " +
        //            Integer.toHexString(r.getCode())+"h");
      }
    }

    bof = null;
    if (excelFile.hasNext())
    {
      r = excelFile.next();

      if (r.getType() == Type.BOF)
      {
        bof = new BOFRecord(r);
      }
    }

    // Only get sheets for which there is a corresponding Boundsheet record
    while (bof != null && getNumberOfSheets() < boundsheets.size())
    {
      if (!bof.isBiff8() && !bof.isBiff7())
      {
        throw new BiffException(BiffException.unrecognizedBiffVersion);
      }

      if (bof.isWorksheet())
      {
        // Read the sheet in
        SheetImpl s = new SheetImpl(excelFile,
                                    sharedStrings,
                                    formattingRecords,
                                    bof,
                                    workbookBof,
                                    nineteenFour,
                                    this);

        BoundsheetRecord br = (BoundsheetRecord) boundsheets.get
          (getNumberOfSheets());
        s.setName(br.getName());
        s.setHidden(br.isHidden());
        addSheet(s);
      }
      else if (bof.isChart())
      {
        // Read the sheet in
        SheetImpl s = new SheetImpl(excelFile,
                                    sharedStrings,
                                    formattingRecords,
                                    bof,
                                    workbookBof,
                                    nineteenFour,
                                    this);

        BoundsheetRecord br = (BoundsheetRecord) boundsheets.get
          (getNumberOfSheets());
        s.setName(br.getName());
        s.setHidden(br.isHidden());
        addSheet(s);
      }
      else
      {
        logger.warn("BOF is unrecognized");


        while (excelFile.hasNext() && r.getType() != Type.EOF)
        {
          r = excelFile.next();
        }
      }

      // The next record will normally be a BOF or empty padding until
      // the end of the block is reached.  In exceptionally unlucky cases,
      // the last EOF  will coincide with a block division, so we have to
      // check there is more data to retrieve.
      // Thanks to liamg for spotting this
      bof = null;
      if (excelFile.hasNext())
      {
        r = excelFile.next();

        if (r.getType() == Type.BOF)
        {
          bof = new BOFRecord(r);
        }
      }
    }

    // Add all the local names to the specific sheets
    for (Iterator it = localNames.iterator() ; it.hasNext() ;)
    {
      NameRecord nr  = (NameRecord) it.next();

      if (nr.getBuiltInName() == null)
      {
        logger.warn("Usage of a local non-builtin name");
      } 
      else if (nr.getBuiltInName() == BuiltInName.PRINT_AREA || 
               nr.getBuiltInName() == BuiltInName.PRINT_TITLES)
      {
        // appears to use the internal tab number rather than the
        // external sheet index
        SheetImpl s = (SheetImpl) sheets.get(nr.getSheetRef() - 1);
        s.addLocalName(nr);
      }
    }
  }

  /**
   * Accessor for the formattingRecords, used by the WritableWorkbook
   * when creating a copy of this
   *
   * @return the formatting records
   */
  public FormattingRecords getFormattingRecords()
  {
    return formattingRecords;
  }

  /**
   * Accessor for the externSheet, used by the WritableWorkbook
   * when creating a copy of this
   *
   * @return the external sheet record
   */
  public ExternalSheetRecord getExternalSheetRecord()
  {
    return externSheet;
  }

  /**
   * Accessor for the MsoDrawingGroup, used by the WritableWorkbook
   * when creating a copy of this
   *
   * @return the Mso Drawing Group record
   */
  public MsoDrawingGroupRecord getMsoDrawingGroupRecord()
  {
    return msoDrawingGroup;
  }

  /**
   * Accessor for the supbook records, used by the WritableWorkbook
   * when creating a copy of this
   *
   * @return the supbook records
   */
  public SupbookRecord[] getSupbookRecords()
  {
    SupbookRecord[] sr = new SupbookRecord[supbooks.size()];
    return (SupbookRecord[]) supbooks.toArray(sr);
  }

  /**
   * Accessor for the name records.  Used by the WritableWorkbook when
   * creating a copy of this
   *
   * @return the array of names
   */
  public NameRecord[] getNameRecords()
  {
    NameRecord[] na = new NameRecord[nameTable.size()];
    return (NameRecord[]) nameTable.toArray(na);
  }

  /**
   * Accessor for the fonts, used by the WritableWorkbook
   * when creating a copy of this
   * @return the fonts used in this workbook
   */
  public Fonts getFonts()
  {
    return fonts;
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
  public Cell getCell(String loc)
  {
    Sheet s = getSheet(CellReferenceHelper.getSheet(loc)); 
    return s.getCell(loc);
  }

  /**
   * Gets the named cell from this workbook.  If the name refers to a
   * range of cells, then the cell on the top left is returned.  If
   * the name cannot be found, null is returned
   *
   * @param  name the name of the cell/range to search for
   * @return the cell in the top left of the range if found, NULL
   *         otherwise
   */
  public Cell findCellByName(String name)
  {
    NameRecord nr = (NameRecord) namedRecords.get(name);

    if (nr == null)
    {
      return null;
    }

    NameRecord.NameRange[] ranges = nr.getRanges();

    // Go and retrieve the first cell in the first range
    Sheet s = getSheet(getExternalSheetIndex(ranges[0].getExternalSheet()));
    int col = ranges[0].getFirstColumn();
    int row = ranges[0].getFirstRow();

    // If the sheet boundaries fall short of the named cell, then return
    // an empty cell to stop an exception being thrown
    if (col > s.getColumns() || row > s.getRows())
    {
      return new EmptyCell(col, row);
    }
    
    Cell cell = s.getCell(col, row);

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
   * @param name the name to find
   * @return the range of cells
   */
  public Range[] findByName(String name)
  {
    NameRecord nr = (NameRecord) namedRecords.get(name);

    if (nr == null)
    {
      return null;
    }

    NameRecord.NameRange[] ranges = nr.getRanges();

    Range[] cellRanges = new Range[ranges.length];

    for (int i = 0; i < ranges.length; i++)
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
   * Gets the named ranges
   *
   * @return the list of named cells within the workbook
   */
  public String[] getRangeNames()
  {
    Object[] keys = namedRecords.keySet().toArray();
    String[] names = new String[keys.length];
    System.arraycopy(keys, 0, names, 0, keys.length);

    return names;
  }

  /**
   * Method used when parsing formulas to make sure we are trying
   * to parse a supported biff version
   *
   * @return the BOF record
   */
  public BOFRecord getWorkbookBof()
  {
    return workbookBof;
  }

  /**
   * Determines whether the sheet is protected
   *
   * @return whether or not the sheet is protected
   */
  public boolean isProtected()
  {
    return wbProtected;
  }

  /**
   * Accessor for the settings
   *
   * @return the workbook settings
   */
  public WorkbookSettings getSettings()
  {
    return settings;
  }

  /**
   * Accessor/implementation method for the external sheet reference
   *
   * @param sheetName the sheet name to look for
   * @return the external sheet index
   */
  public int getExternalSheetIndex(String sheetName)
  {
    return 0;
  }

  /**
   * Accessor/implementation method for the external sheet reference
   *
   * @param sheetName the sheet name to look for
   * @return the external sheet index
   */
  public int getLastExternalSheetIndex(String sheetName)
  {
    return 0;
  }

  /**
   * Gets the name at the specified index
   *
   * @param index the index into the name table
   * @return the name of the cell
   * @exception NameRangeException
   */
  public String getName(int index) throws NameRangeException
  {
    //    Assert.verify(index >= 0 && index < nameTable.size());
    if (index < 0 || index >= nameTable.size())
    {
      throw new NameRangeException();
    }
    return ((NameRecord) nameTable.get(index)).getName();
  }

  /**
   * Gets the index of the name record for the name
   *
   * @param name the name to search for
   * @return the index in the name table
   */
  public int getNameIndex(String name)
  {
    NameRecord nr = (NameRecord) namedRecords.get(name);

    return nr != null ? nr.getIndex() : 0;
  }

  /**
   * Accessor for the drawing group
   *
   * @return  the drawing group
   */
  public DrawingGroup getDrawingGroup()
  {
    return drawingGroup;
  }

  /**
   * Accessor for the CompoundFile.  For this feature to return non-null
   * value, the propertySets feature in WorkbookSettings must be enabled
   * and the workbook must contain additional property sets.  This
   * method is used during the workbook copy
   *
   * @return the base compound file if it contains additional data items
   *         and property sets are enabled
   */
  public CompoundFile getCompoundFile()
  {
    return excelFile.getCompoundFile();
  }

  /**
   * Accessor for the containsMacros
   *
   * @return TRUE if this workbook contains macros, FALSE otherwise
   */
  public boolean containsMacros()
  {
    return containsMacros;
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
   * Accessor for the country record, using during copying
   *
   * @return the country record read in
   */
  public CountryRecord getCountryRecord()
  {
    return countryRecord;
  }

  /**
   * Accessor for addin function names
   *
   * @return list of add in function names
   */
  public String[] getAddInFunctionNames()
  {
    String[] addins = new String[0];
    return (String[]) addInFunctions.toArray(addins);
  }

  /**
   * Gets the sheet index in this workbook.  Used when importing a sheet
   *
   * @param sheet the sheet
   * @return the 0-based sheet index, or -1 if it is not found
   */
  public int getIndex(Sheet sheet)
  {
    String name = sheet.getName();
    int index = -1;
    int pos = 0;

    for (Iterator i = boundsheets.iterator() ; i.hasNext() && index == -1 ;)
    {
      BoundsheetRecord br = (BoundsheetRecord) i.next();

      if (br.getName().equals(name))
      {
        index = pos;
      }
      else
      {
        pos++;
      }
    }

    return index;
  }

  public XCTRecord[] getXCTRecords()
  {
    XCTRecord[] xctr = new XCTRecord[0];
    return (XCTRecord[]) xctRecords.toArray(xctr);
  }
}







