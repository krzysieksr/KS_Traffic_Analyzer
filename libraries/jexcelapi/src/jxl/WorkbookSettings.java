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

import java.io.File;
import java.util.HashMap;
import java.util.Locale;

import jxl.common.Logger;

import jxl.biff.CountryCode;
import jxl.biff.formula.FunctionNames;

/**
 * This is a bean which client applications may use to set various advanced
 * workbook properties.  Use of this bean is not mandatory, and its absence
 * will merely result in workbooks being read/written using the default
 * settings
 */
public final class WorkbookSettings
{
  /**
   * The logger
   */
  private static Logger logger = Logger.getLogger(WorkbookSettings.class);

  /**
   * The amount of memory allocated to store the workbook data when
   * reading a worksheet.  For processeses reading many small workbooks inside
   * a WAS it might be necessary to reduce the default size
   */
  private int initialFileSize;

  /**
   * The amount of memory allocated to the array containing the workbook
   * data when its current amount is exhausted.
   */
  private int arrayGrowSize;

  /**
   * Flag to indicate whether the drawing feature is enabled or not
   * Drawings deactivated using -Djxl.nodrawings=true on the JVM command line
   * Activated by default or by using -Djxl.nodrawings=false on the JVM command
   * line
   */
  private boolean drawingsDisabled;

  /**
   * Flag to indicate whether the name feature is enabled or not
   * Names deactivated using -Djxl.nonames=true on the JVM command line
   * Activated by default or by using -Djxl.nonames=false on the JVM command
   * line
   */
  private boolean namesDisabled;

  /**
   * Flag to indicate whether formula cell references should be adjusted
   * following row/column insertion/deletion
   */
  private boolean formulaReferenceAdjustDisabled;

  /**
   * Flag to indicate whether the system hint garbage collection
   * is enabled or not.
   * As a rule of thumb, it is desirable to enable garbage collection
   * when reading large spreadsheets from  a batch process or from the
   * command line, but better to deactivate the feature when reading
   * large spreadsheets within a WAS, as the calls to System.gc() not
   * only garbage collect the junk in JExcelApi, but also in the
   * webservers JVM and can cause significant slowdown
   * GC deactivated using -Djxl.nogc=true on the JVM command line
   * Activated by default or by using -Djxl.nogc=false on the JVM command line
   */
  private boolean gcDisabled;

  /**
   * Flag to indicate whether the rationalization of cell formats is
   * disabled or not.
   * Rationalization is enabled by default, but may be disabled for
   * performance reasons.  It can be deactivated using -Djxl.norat=true on
   * the JVM command line
   */
  private boolean rationalizationDisabled;

  /**
   * Flag to indicate whether or not the merged cell checking has been
   * disabled
   */
  private boolean mergedCellCheckingDisabled;

  /**
   * Flag to indicate whether the copying of additional property sets
   * are disabled
   */
  private boolean propertySetsDisabled;

  /**
   * Flag to indicate that cell validation criteria are ignored
   */
  private boolean cellValidationDisabled;

  /**
   * Flag to indicate whether or not to ignore blank cells when processing
   * sheets.  Cells which are identified as blank can still have associated
   * cell formats which the processing program may still need to read
   */
  private boolean ignoreBlankCells;

  /**
   * Flag to indicate whether auto filtering should be read/copied
   */
  private boolean autoFilterDisabled;

  /**
   * Flag to indicate whether a temporary file should be used when
   * writing out the workbook
   */
  private boolean useTemporaryFileDuringWrite;

  /**
   * The directory for used for the temporary file during write.  If this
   * is NULL, the default system directory is used
   */
  private File temporaryFileDuringWriteDirectory;

  /**
   * The locale.  Normally this is the same as the system locale, but there
   * may be cases (eg. where you are uploading many spreadsheets from foreign
   * sources) where you may want to specify the locale on an individual
   * worksheet basis
   * The locale may also be specified on the command line using the lang and
   * country System properties eg. -Djxl.lang=en -Djxl.country=UK for UK
   * English
   */
  private Locale locale;

  /**
   * The locale specific function names for this workbook
   */
  private FunctionNames functionNames;

  /**
   * The character encoding used for reading non-unicode strings.  This can
   * be different from the default platform encoding if processing spreadsheets
   * from abroad.  This may also be set using the system property jxl.encoding
   */
  private String encoding;

  /**
   * The character set used by the readable spreadsheeet
   */
  private int characterSet;

  /**
   * The display language used by Excel (ISO 3166 mnemonic)
   */
  private String excelDisplayLanguage;

  /**
   * The regional settings used by Excel (ISO 3166 mnemonic)
   */
  private String excelRegionalSettings;

  /**
   * A hash map of function names keyed on locale
   */
  private HashMap localeFunctionNames;

  /**
   * Flag to indicate whether all external data and pivot stuff should 
   * refreshed
   */
  private boolean refreshAll;

  /**
   * Flag to indicate whether the file is a template or not (Usually with .xlt
   * file name extension)
   */
  private boolean template;

  /**
   * Flag to indicate whether the file has been written by excel 2000.
   *
   * The EXCEL9FILE record indicates the file was written by Excel 2000. It has
   * no record data field and is C0010000h. Any application other than Excel
   * 2000 that edits the file should not write out this record.
   *
   * However, it seemas that excel 2003 + 2007 still set this flag....
   */
  private boolean excel9file= false;

  /**
   * The WINDOWPROTECT record stores an option from the Protect Workbook
   * dialog box.
   *
   * =1 if the workbook windows are protected
   */
  private boolean windowProtected;

  /**
   * Write access user name.
   * When not set (null) then we set it to  Java Excel API + Version number
   */
  private String writeAccess;

  /**
   * The HIDEOBJ record stores options selected in the Options dialog,View tab.
   */
  private int hideobj;

  /**
   * The HIDEOBJ record stores options selected in the Options dialog,View tab.
   */
  public final static int HIDEOBJ_HIDE_ALL= 2;

  /**
   * The HIDEOBJ record stores options selected in the Options dialog,View tab.
   */
  public final static int HIDEOBJ_SHOW_PLACEHOLDERS= 1;
  /**
   * The HIDEOBJ record stores options selected in the Options dialog,View tab.
   */
  public final static int HIDEOBJ_SHOW_ALL= 0;

  // **
  // The default values
  // **
  private static final int DEFAULT_INITIAL_FILE_SIZE = 5 * 1024 * 1024;
    // 5 megabytes
  private static final int DEFAULT_ARRAY_GROW_SIZE = 1024 * 1024; // 1 megabyte
  
  /**
   * Default constructor
   */
  public WorkbookSettings()
  {
    initialFileSize = DEFAULT_INITIAL_FILE_SIZE;
    arrayGrowSize = DEFAULT_ARRAY_GROW_SIZE;
    localeFunctionNames = new HashMap();
    excelDisplayLanguage = CountryCode.USA.getCode();
    excelRegionalSettings = CountryCode.UK.getCode();
    refreshAll = false;
    template = false;
    excel9file = false;
    windowProtected = false;
    hideobj = HIDEOBJ_SHOW_ALL;

    // Initialize other properties from the system properties
    try
    {
      boolean suppressWarnings = Boolean.getBoolean("jxl.nowarnings");
      setSuppressWarnings(suppressWarnings);
      drawingsDisabled        = Boolean.getBoolean("jxl.nodrawings");
      namesDisabled           = Boolean.getBoolean("jxl.nonames");
      gcDisabled              = Boolean.getBoolean("jxl.nogc");
      rationalizationDisabled = Boolean.getBoolean("jxl.norat");
      mergedCellCheckingDisabled =
        Boolean.getBoolean("jxl.nomergedcellchecks");
      formulaReferenceAdjustDisabled =
                                Boolean.getBoolean("jxl.noformulaadjust");
      propertySetsDisabled = Boolean.getBoolean("jxl.nopropertysets");
      ignoreBlankCells = Boolean.getBoolean("jxl.ignoreblanks");
      cellValidationDisabled = Boolean.getBoolean("jxl.nocellvalidation");
      autoFilterDisabled = !Boolean.getBoolean("jxl.autofilter"); 
             // autofilter currently disabled by default
      useTemporaryFileDuringWrite = 
        Boolean.getBoolean("jxl.usetemporaryfileduringwrite");
      String tempdir =
        System.getProperty("jxl.temporaryfileduringwritedirectory");

      if (tempdir != null)
      {
        temporaryFileDuringWriteDirectory = new File(tempdir);
      }
      
      encoding = System.getProperty("file.encoding");
    }
    catch (SecurityException e)
    {
      logger.warn("Error accessing system properties.", e);
    }

    // Initialize the locale to the system locale
    try
    {
      if (System.getProperty("jxl.lang")    == null ||
          System.getProperty("jxl.country") == null)
      {
        locale = Locale.getDefault();
      }
      else
      {
        locale = new Locale(System.getProperty("jxl.lang"),
                            System.getProperty("jxl.country"));
      }

      if (System.getProperty("jxl.encoding") != null)
      {
        encoding = System.getProperty("jxl.encoding");
      }
    } 
    catch (SecurityException e)
    {
      logger.warn("Error accessing system properties.", e);
      locale = Locale.getDefault();
    }
  }

  /**
   * Sets the amount of memory by which to increase the amount of
   * memory allocated to storing the workbook data.
   * For processeses reading many small workbooks
   * inside  a WAS it might be necessary to reduce the default size
   * Default value is 1 megabyte
   *
   * @param sz the file size in bytes
   */
  public void setArrayGrowSize(int sz)
  {
    arrayGrowSize = sz;
  }

  /**
   * Accessor for the array grow size property
   *
   * @return the array grow size
   */
  public int getArrayGrowSize()
  {
    return arrayGrowSize;
  }

  /**
   * Sets the initial amount of memory allocated to store the workbook data
   * when reading a worksheet.  For processeses reading many small workbooks
   * inside  a WAS it might be necessary to reduce the default size
   * Default value is 5 megabytes
   *
   * @param sz the file size in bytes
   */
  public void setInitialFileSize(int sz)
  {
    initialFileSize = sz;
  }

  /**
   * Accessor for the initial file size property
   *
   * @return the initial file size
   */
  public int getInitialFileSize()
  {
    return initialFileSize;
  }

  /**
   * Gets the drawings disabled flag
   *
   * @return TRUE if drawings are disabled, FALSE otherwise
   */
  public boolean getDrawingsDisabled()
  {
    return drawingsDisabled;
  }

  /**
   * Accessor for the disabling of garbage collection
   *
   * @return FALSE if JExcelApi hints for garbage collection, TRUE otherwise
   */
  public boolean getGCDisabled()
  {
    return gcDisabled;
  }

  /**
   * Accessor for the disabling of interpretation of named ranges
   *
   * @return FALSE if named cells are interpreted, TRUE otherwise
   */
  public boolean getNamesDisabled()
  {
    return namesDisabled;
  }

  /**
   * Disables the handling of names
   *
   * @param b TRUE to disable the names feature, FALSE otherwise
   */
  public void setNamesDisabled(boolean b)
  {
    namesDisabled = b;
  }

  /**
   * Disables the handling of drawings
   *
   * @param b TRUE to disable the names feature, FALSE otherwise
   */
  public void setDrawingsDisabled(boolean b)
  {
    drawingsDisabled = b;
  }

  /**
   * Sets whether or not to rationalize the cell formats before
   * writing out the sheet.  The default value is true
   *
   * @param r the rationalization flag
   */
  public void setRationalization(boolean r)
  {
    rationalizationDisabled = !r;
  }

  /**
   * Accessor to retrieve the rationalization flag
   *
   * @return TRUE if rationalization is off, FALSE if rationalization is on
   */
  public boolean getRationalizationDisabled()
  {
    return rationalizationDisabled;
  }

  /**
   * Accessor to retrieve the merged cell checking flag
   *
   * @return TRUE if merged cell checking is off, FALSE if it is on
   */
  public boolean getMergedCellCheckingDisabled()
  {
    return mergedCellCheckingDisabled;
  }

  /**
   * Accessor to set the merged cell checking
   *
   * @param b - TRUE to enable merged cell checking, FALSE otherwise
   */
  public void setMergedCellChecking(boolean b)
  {
    mergedCellCheckingDisabled = !b;
  }

  /**
   * Sets whether or not to enable any property sets (such as macros)
   * to be copied along with the workbook
   * Leaving this feature enabled will result in the JXL process using
   * more memory
   *
   * @param r the property sets flag
   */
  public void setPropertySets(boolean r)
  {
    propertySetsDisabled = !r;
  }

  /**
   * Accessor to retrieve the property sets disabled flag
   *
   * @return TRUE if property sets are disabled, FALSE otherwise
   */
  public boolean getPropertySetsDisabled()
  {
    return propertySetsDisabled;
  }

  /**
   * Accessor to set the suppress warnings flag.  Due to the change
   * in logging in version 2.4, this will now set the warning
   * behaviour across the JVM (depending on the type of logger used)
   *
   * @param w the flag
   */
  public void setSuppressWarnings(boolean w)
  {
    logger.setSuppressWarnings(w);
  }

  /**
   * Accessor for the formula adjust disabled
   *
   * @return TRUE if formulas are adjusted following row/column inserts/deletes
   *         FALSE otherwise
   */
  public boolean getFormulaAdjust()
  {
    return !formulaReferenceAdjustDisabled;
  }

  /**
   * Setter for the formula adjust disabled property
   *
   * @param b TRUE to adjust formulas, FALSE otherwise
   */
  public void setFormulaAdjust(boolean b)
  {
    formulaReferenceAdjustDisabled = !b;
  }

  /**
   * Sets the locale used by JExcelApi to generate the spreadsheet.
   * Setting this value has no effect on the language or region of
   * the generated excel file
   *
   * @param l the locale
   */
  public void setLocale(Locale l)
  {
    locale = l;
  }

  /**
   * Returns the locale used by JExcelAPI to read the spreadsheet
   *
   * @return the locale
   */
  public Locale getLocale()
  {
    return locale;
  }

  /**
   * Accessor for the character encoding
   *
   * @return the character encoding for this workbook
   */
  public String getEncoding()
  {
    return encoding;
  }

  /**
   * Sets the encoding for this workbook
   *
   * @param enc the encoding
   */
  public void setEncoding(String enc)
  {
    encoding = enc;
  }

  /**
   * Gets the function names.  This is used by the formula parsing package
   * in order to get the locale specific function names for this particular
   * workbook
   *
   * @return the list of function names
   */
  public FunctionNames getFunctionNames()
  {
    if (functionNames == null)
    {
      functionNames = (FunctionNames) localeFunctionNames.get(locale);

      // have not previously accessed function names for this locale,
      // so create a brand new one and add it to the list
      if (functionNames == null)
      {
        functionNames = new FunctionNames(locale);
        localeFunctionNames.put(locale, functionNames);
      }
    }

    return functionNames;
  }

  /**
   * Accessor for the character set.   This value is only used for reading
   * and has no effect when writing out the spreadsheet
   *
   * @return the character set used by this spreadsheet
   */
  public int getCharacterSet()
  {
    return characterSet;
  }

  /**
   * Sets the character set.  This is only used when the spreadsheet is
   * read, and has no effect when the spreadsheet is written
   *
   * @param cs the character set encoding value
   */
  public void setCharacterSet(int cs)
  {
    characterSet = cs;
  }

  /**
   * Sets the garbage collection disabled
   *
   * @param disabled TRUE to disable garbage collection, FALSE to enable it
   */
  public void setGCDisabled(boolean disabled)
  {
    gcDisabled = disabled;
  }

  /**
   * Sets the ignore blanks flag
   *
   * @param ignoreBlanks TRUE to ignore blanks, FALSE to take them into account
   */
  public void setIgnoreBlanks(boolean ignoreBlanks)
  {
    ignoreBlankCells = ignoreBlanks;
  }

  /**
   * Accessor for the ignore blanks flag
   *
   * @return TRUE if blank cells are being ignored, FALSE otherwise
   */
  public boolean getIgnoreBlanks()
  {
    return ignoreBlankCells;
  }

  /**
   * Sets the ignore cell validation flag
   *
   * @param cv TRUE to disable cell validation, FALSE to enable it
   */
  public void setCellValidationDisabled(boolean cv)
  {
    cellValidationDisabled = cv;
  }

  /**
   * Accessor for the ignore cell validation
   *
   * @return TRUE if cell validation is disabled
   */
  public boolean getCellValidationDisabled()
  {
    return cellValidationDisabled;
  }

  /**
   * Returns the two character ISO 3166 mnemonic used by excel for user
   * language displayto display
   * @return the display language
   */
  public String getExcelDisplayLanguage()
  {
    return excelDisplayLanguage;
  }

  /**
   * Returns the two character ISO 3166 mnemonic used by excel for
   * its regional settings
   * @return the regional settings
   */
  public String getExcelRegionalSettings()
  {
    return excelRegionalSettings;
  }

  /**
   * Sets the language in which the generated file will display
   *
   * @param code the two character ISO 3166 country code
   */
  public void setExcelDisplayLanguage(String code)
  {
    excelDisplayLanguage = code;
  }

  /**
   * Sets the regional settings for the generated excel file
   *
   * @param code the two character ISO 3166 country code
   */
  public void setExcelRegionalSettings(String code)
  {
    excelRegionalSettings = code;
  }

  /**
   * Accessor for the autofilter disabled feature
   *
   * @return TRUE if autofilter is disabled, FALSE otherwise
   */
  public boolean getAutoFilterDisabled()
  {
    return autoFilterDisabled;
  }

  /**
   * Sets the autofilter disabled 
   *
   * @param disabled 
   */
  public void setAutoFilterDisabled(boolean disabled)
  {
    autoFilterDisabled = disabled;
  }

  /**
   * Accessor for the temporary file during write.  If this is set, then
   * when the workbook is written a temporary file will be used to store
   * the interim binary data, otherwise it will take place in memory.  Setting
   * this flag involves an assessment of the trade-offs between memory usage
   * and performance
   *
   * @return TRUE if a temporary is file is used during writing, 
   * FALSE otherwise
   */
  public boolean getUseTemporaryFileDuringWrite()
  {
    return useTemporaryFileDuringWrite;
  }

  /**
   * Sets whether a temporary file is used during the generation of
   * the workbook.  If not set, the workbook will take place entirely in
   * memory.   Setting
   * this flag involves an assessment of the trade-offs between memory usage
   * and performance
   *
   * @return TRUE if a temporary is file is used during writing, 
   * FALSE otherwise
   */
  public void setUseTemporaryFileDuringWrite(boolean temp)
  {
    useTemporaryFileDuringWrite = temp;
  }

  /**
   * Used in conjunction with the UseTemporaryFileDuringWrite setting to
   * set the target directory for the temporary files.   If this is not set,
   * the system default temporary directory is used.
   * This has no effect unless the useTemporaryFileDuringWrite setting
   * is TRUE
   *
   * @param dir the directory to which temporary files should be written
   */
  public void setTemporaryFileDuringWriteDirectory(File dir)
  {
    temporaryFileDuringWriteDirectory = dir;
  }

  /**
   * Used in conjunction with the UseTemporaryFileDuringWrite setting to
   * set the target directory for the temporary files.  This value can 
   * be NULL, in which case the normal system default temporary directory
   * is used instead
   *
   * @return the temporary directory used during write, or NULL if it is 
   *         not set
   */
  public File getTemporaryFileDuringWriteDirectory()
  {
    return temporaryFileDuringWriteDirectory;
  }


  /**
   * When true then Refresh All should be done on all external data ranges and
   * PivotTables when loading the workbook (the default is =0)
   *
   * @param refreshAll the refreshAll to set
   */
  public void setRefreshAll(boolean refreshAll)
  {
    this.refreshAll = refreshAll;
  }

  /**
   * When true then Refresh All should be done on all external data ranges and
   * PivotTables when loading the workbook (the default is =0)
   *
   * @return the refreshAll value
   */
  public boolean getRefreshAll()
  {
    return refreshAll;
  }

  /**
   * Workbook Is a Template
   * @return the template
   */
  public boolean getTemplate()
  {
    return template;
  }

  /**
   * Workbook Is a Template
   * @param template the template to set
   */
  public void setTemplate(boolean template)
  {
    this.template = template;
  }

  /**
   * Has this file been written by excel 2000?
   * 
   * @return the excel9file
   */
  public boolean getExcel9File()
  {
    return excel9file;
  }

  /**
   * @param excel9file the excel9file to set
   */
  public void setExcel9File(boolean excel9file)
  {
    this.excel9file = excel9file;
  }

  /**
   * @return the windowprotected
   */
  public boolean getWindowProtected()
  {
    return windowProtected;
  }

  /**
   * @param windowprotected the windowprotected to set
   */
  public void setWindowProtected(boolean windowprotected)
  {
    this.windowProtected = windowProtected;
  }

  /**
   * The HIDEOBJ record stores options selected in the Options dialog,View tab
   *
   * Possible values are:
   * HIDEOBJ_HIDE_ALL, HIDEOBJ_SHOW_ALL and HIDEOBJ_SHOW_PLACEHOLDERS
   * @return the hideobj
   */
  public int getHideobj()
  {
    return hideobj;
  }

  /**
   * The HIDEOBJ record stores options selected in the Options dialog,View tab
   *
   * Possible values are:
   * HIDEOBJ_HIDE_ALL, HIDEOBJ_SHOW_ALL and HIDEOBJ_SHOW_PLACEHOLDERS
   * @param hideobj the hideobj to set
   */
  public void setHideobj(int hideobj)
  {
    this.hideobj = hideobj;
  }

  /**
   * @return the writeAccess
   */
  public String getWriteAccess()
  {
    return writeAccess;
  }

  /**
   * @param writeAccess the writeAccess to set
   */
  public void setWriteAccess(String writeAccess)
  {
    this.writeAccess = writeAccess;
  }
}

