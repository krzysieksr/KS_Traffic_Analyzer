/*********************************************************************
*
*      Copyright (C) 2004 Andrew Khan
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

package jxl.biff;

import jxl.common.Assert;
import jxl.common.Logger;

import java.text.MessageFormat;
import java.text.DecimalFormat;
import java.util.Collection;
import java.util.Iterator;

import jxl.WorkbookSettings;
import jxl.biff.formula.ExternalSheet;
import jxl.biff.formula.FormulaException;
import jxl.biff.formula.FormulaParser;
import jxl.biff.formula.ParseContext;

/**
 * Class which parses the binary data associated with Data Validity (DV)
 * setting
 */
public class DVParser
{
  /**
   * The logger
   */
  private static Logger logger = Logger.getLogger(DVParser.class);

  // DV Type
  public static class DVType 
  {
    private int value;
    private String desc;
    
    private static DVType[] types = new DVType[0];
   
    DVType(int v, String d) 
    {
      value = v;
      desc = d;
      DVType[] oldtypes = types;
      types = new DVType[oldtypes.length+1];
      System.arraycopy(oldtypes, 0, types, 0, oldtypes.length);
      types[oldtypes.length] = this;
    }

    static DVType getType(int v)
    {
      DVType found = null;
      for (int i = 0 ; i < types.length && found == null ; i++)
      {
        if (types[i].value == v)
        {
          found = types[i];
        }
      }
      return found;
    }

    public int getValue() 
    {
      return value;
    }

    public String getDescription()
    {
      return desc;
    }
  }

  // Error Style
  public static class ErrorStyle
  {
    private int value;
    
    private static ErrorStyle[] types = new ErrorStyle[0];
   
    ErrorStyle(int v) 
    {
      value = v;
      ErrorStyle[] oldtypes = types;
      types = new ErrorStyle[oldtypes.length+1];
      System.arraycopy(oldtypes, 0, types, 0, oldtypes.length);
      types[oldtypes.length] = this;
    }

    static ErrorStyle getErrorStyle(int v)
    {
      ErrorStyle found = null;
      for (int i = 0 ; i < types.length && found == null ; i++)
      {
        if (types[i].value == v)
        {
          found = types[i];
        }
      }
      return found;
    }

    public int getValue() 
    {
      return value;
    }
  }

  // Conditions
  public static class Condition
  {
    private int value;
    private MessageFormat format;
    
    private static Condition[] types = new Condition[0];
   
    Condition(int v, String pattern) 
    {
      value = v;
      format = new MessageFormat(pattern);
      Condition[] oldtypes = types;
      types = new Condition[oldtypes.length+1];
      System.arraycopy(oldtypes, 0, types, 0, oldtypes.length);
      types[oldtypes.length] = this;
    }

    static Condition getCondition(int v)
    {
      Condition found = null;
      for (int i = 0 ; i < types.length && found == null ; i++)
      {
        if (types[i].value == v)
        {
          found = types[i];
        }
      }
      return found;
    }

    public int getValue() 
    {
      return value;
    }

    public String getConditionString(String s1, String s2)
    {
      return format.format(new String[] {s1, s2});
    }
  }

  // The values
  public static final DVType ANY = new DVType(0, "any");
  public static final DVType INTEGER = new DVType(1, "int");
  public static final DVType DECIMAL = new DVType(2, "dec");
  public static final DVType LIST = new DVType(3, "list");
  public static final DVType DATE = new DVType(4, "date");
  public static final DVType TIME = new DVType(5, "time");
  public static final DVType TEXT_LENGTH = new DVType(6, "strlen");
  public static final DVType FORMULA = new DVType(7, "form");

  // The error styles
  public static final ErrorStyle STOP = new ErrorStyle(0);
  public static final ErrorStyle WARNING = new ErrorStyle(1);
  public static final ErrorStyle INFO = new ErrorStyle(2);

  // The conditions
  public static final Condition BETWEEN = new Condition(0, "{0} <= x <= {1}");
  public static final Condition NOT_BETWEEN = 
    new Condition(1, "!({0} <= x <= {1}");
  public static final Condition EQUAL = new Condition(2, "x == {0}");
  public static final Condition NOT_EQUAL = new Condition(3, "x != {0}");
  public static final Condition GREATER_THAN = new Condition(4, "x > {0}");
  public static final Condition LESS_THAN = new Condition(5, "x < {0}");
  public static final Condition GREATER_EQUAL = new Condition(6, "x >= {0}");
  public static final Condition LESS_EQUAL = new Condition(7, "x <= {0}");

  // The masks
  private static final int STRING_LIST_GIVEN_MASK = 0x80;
  private static final int EMPTY_CELLS_ALLOWED_MASK = 0x100;
  private static final int SUPPRESS_ARROW_MASK = 0x200;
  private static final int SHOW_PROMPT_MASK = 0x40000;
  private static final int SHOW_ERROR_MASK = 0x80000;

  // The decimal format
  private static DecimalFormat DECIMAL_FORMAT = new DecimalFormat("#.#");

  // The maximum string length for a data validation list
  private static final int MAX_VALIDATION_LIST_LENGTH = 254;

  // The maximum number of rows and columns
  private static final int MAX_ROWS=0xffff;
  private static final int MAX_COLUMNS=0xff;

  /**
   * The type
   */
  private DVType type;

  /**
   * The error style
   */
  private ErrorStyle errorStyle;

  /**
   * The condition
   */
  private Condition condition;

  /**
   * String list option
   */
  private boolean stringListGiven;

  /**
   * Empty cells allowed
   */
  private boolean emptyCellsAllowed;

  /**
   * Suppress arrow
   */
  private boolean suppressArrow;

  /**
   * Show prompt
   */
  private boolean showPrompt;

  /**
   * Show error
   */
  private boolean showError;

  /**
   * The title of the prompt box
   */
  private String promptTitle;

  /**
   * The title of the error box
   */
  private String errorTitle;

  /**
   * The text of the prompt box
   */
  private String promptText;

  /**
   * The text of the error box
   */
  private String errorText;

  /**
   * The first formula
   */
  private FormulaParser formula1;

  /**
   * The first formula string
   */
  private String formula1String;

  /**
   * The second formula
   */
  private FormulaParser formula2;

  /**
   * The second formula string
   */
  private String formula2String;

  /**
   * The column number of the cell at the top left of the range
   */
  private int column1;

  /**
   * The row number of the cell at the top left of the range
   */
  private int row1;

  /**
   * The column index of the cell at the bottom right
   */
  private int column2;

  /**
   * The row index of the cell at the bottom right
   */
  private int row2;

  /**
   * Flag to indicate that this DV Parser is shared amongst a group
   * of cells
   */
  private boolean extendedCellsValidation;

  /**
   * Flag indicated whether this has been copied
   */
  private boolean copied;

  /**
   * Constructor
   */
  public DVParser(byte[] data, 
                  ExternalSheet es, 
                  WorkbookMethods nt,
                  WorkbookSettings ws)
  {
    Assert.verify(nt != null);

    copied = false;
    int options = IntegerHelper.getInt(data[0], data[1], data[2], data[3]);

    int typeVal = options & 0xf;
    type = DVType.getType(typeVal);

    int errorStyleVal = (options & 0x70) >> 4;
    errorStyle = ErrorStyle.getErrorStyle(errorStyleVal);

    int conditionVal = (options & 0xf00000) >> 20;
    condition = Condition.getCondition(conditionVal);

    stringListGiven = (options & STRING_LIST_GIVEN_MASK) != 0;
    emptyCellsAllowed = (options & EMPTY_CELLS_ALLOWED_MASK) != 0;
    suppressArrow = (options & SUPPRESS_ARROW_MASK) != 0;
    showPrompt = (options & SHOW_PROMPT_MASK) != 0;
    showError = (options & SHOW_ERROR_MASK) != 0;

    int pos = 4;
    int length = IntegerHelper.getInt(data[pos], data[pos+1]);
    if (length > 0 && data[pos + 2] == 0)
    {
      promptTitle = StringHelper.getString(data, length, pos + 3, ws);
      pos += length + 3;
    }
    else if (length > 0)
    {
      promptTitle = StringHelper.getUnicodeString(data, length, pos + 3);
      pos += length * 2 + 3;
    }
    else
    {
      pos += 3;
    }

    length = IntegerHelper.getInt(data[pos], data[pos+1]);
    if (length > 0 && data[pos + 2] == 0)
    {
      errorTitle = StringHelper.getString(data, length, pos + 3, ws);
      pos += length + 3;
    }
    else if (length > 0)
    {
      errorTitle = StringHelper.getUnicodeString(data, length, pos + 3);
      pos += length * 2 + 3;
    }
    else
    {
      pos += 3;
    }

    length = IntegerHelper.getInt(data[pos], data[pos+1]);
    if (length > 0 && data[pos + 2] == 0)
    {
      promptText = StringHelper.getString(data, length, pos + 3, ws);
      pos += length + 3;
    }
    else if (length > 0)
    {
      promptText = StringHelper.getUnicodeString(data, length, pos + 3);
      pos += length * 2 + 3;
    }
    else
    {
      pos += 3;
    }

    length = IntegerHelper.getInt(data[pos], data[pos+1]);
    if (length > 0 && data[pos + 2] == 0)
    {
      errorText = StringHelper.getString(data, length, pos + 3, ws);
      pos += length + 3;
    }
    else if (length > 0)
    {
      errorText = StringHelper.getUnicodeString(data, length, pos + 3);
      pos += length * 2 + 3;
    }
    else
    {
      pos += 3;
    }

    int formula1Length = IntegerHelper.getInt(data[pos], data[pos+1]);
    pos += 4;
    int formula1Pos = pos;
    pos += formula1Length;

    int formula2Length = IntegerHelper.getInt(data[pos], data[pos+1]);
    pos += 4;
    int formula2Pos = pos;
    pos += formula2Length;

    pos += 2;

    row1 = IntegerHelper.getInt(data[pos], data[pos+1]);
    pos += 2;

    row2 = IntegerHelper.getInt(data[pos], data[pos+1]);
    pos += 2;

    column1 = IntegerHelper.getInt(data[pos], data[pos+1]);
    pos += 2;

    column2 = IntegerHelper.getInt(data[pos], data[pos+1]);
    pos += 2;

    extendedCellsValidation = (row1 == row2 && column1 == column2) ? 
      false : true;

    // Do the formulas
    try
    {
      // First, create a temporary  blank cell for any formula relative 
      // references
      EmptyCell tmprt = new EmptyCell(column1, row1);

      if (formula1Length != 0)
      {
        byte[] tokens = new byte[formula1Length];
        System.arraycopy(data, formula1Pos, tokens, 0, formula1Length);
        formula1 = new FormulaParser(tokens, tmprt, es, nt,ws, 
                                     ParseContext.DATA_VALIDATION);
        formula1.parse();
      }

      if (formula2Length != 0)
      {
        byte[] tokens = new byte[formula2Length];
        System.arraycopy(data, formula2Pos, tokens, 0, formula2Length);
        formula2 = new FormulaParser(tokens, tmprt, es, nt, ws, 
                                     ParseContext.DATA_VALIDATION);
        formula2.parse();
      }
    }
    catch (FormulaException e)
    {
      logger.warn(e.getMessage() + " for cells " + 
      CellReferenceHelper.getCellReference(column1, row1)+ "-" + 
      CellReferenceHelper.getCellReference(column2, row2));
    }
  }

  /**
   * Constructor called when creating a data validation from the API
   */
  public DVParser(Collection strings)
  {
    copied = false;
    type = LIST;
    errorStyle = STOP;
    condition = BETWEEN;
    extendedCellsValidation = false;
    
    // the options
    stringListGiven = true;
    emptyCellsAllowed = true;
    suppressArrow = false;
    showPrompt = true;
    showError = true;

    promptTitle = "\0";
    errorTitle = "\0";
    promptText = "\0";
    errorText = "\0";
    if (strings.size() == 0)
    {
      logger.warn("no validation strings - ignoring");
    }

    Iterator i = strings.iterator();
    StringBuffer formulaString = new StringBuffer();

    formulaString.append(i.next().toString());
    while (i.hasNext())
    {
      formulaString.append('\0');
      formulaString.append(' ');
      formulaString.append(i.next().toString());
    }

    // If the formula string exceeds
    // the maximum validation list length, then truncate and stop there
    if (formulaString.length() > MAX_VALIDATION_LIST_LENGTH)
    {
      logger.warn("Validation list exceeds maximum number of characters - " +
                  "truncating");
      formulaString.delete(MAX_VALIDATION_LIST_LENGTH, 
                           formulaString.length());
    }

    // Put the string in quotes
    formulaString.insert(0, '\"');
    formulaString.append('\"');
    formula1String = formulaString.toString();
  }

  /**
   * Constructor called when creating a data validation from the API
   */
  public DVParser(String namedRange)
  {
    // Handle the case for an empty string
    if (namedRange.length() == 0)
    {
      copied = false;
      type = FORMULA;
      errorStyle = STOP;
      condition = EQUAL;
      extendedCellsValidation = false;
      // the options
      stringListGiven = false;
      emptyCellsAllowed = false;
      suppressArrow = false;
      showPrompt = true;
      showError = true;
      
      promptTitle = "\0";
      errorTitle = "\0";
      promptText = "\0";
      errorText = "\0";
      formula1String = "\"\"";
      return;
    }

    copied = false;
    type = LIST;
    errorStyle = STOP;
    condition = BETWEEN;
    extendedCellsValidation = false;
    
    // the options
    stringListGiven = false;
    emptyCellsAllowed = true;
    suppressArrow = false;
    showPrompt = true;
    showError = true;

    promptTitle = "\0";
    errorTitle = "\0";
    promptText = "\0";
    errorText = "\0";
    formula1String = namedRange;
  }

  /**
   * Constructor called when creating a data validation from the API
   */
  public DVParser(int c1, int r1, int c2, int r2)
  {
    copied = false;
    type = LIST;
    errorStyle = STOP;
    condition = BETWEEN;
    extendedCellsValidation = false;
    
    // the options
    stringListGiven = false;
    emptyCellsAllowed = true;
    suppressArrow = false;
    showPrompt = true;
    showError = true;

    promptTitle = "\0";
    errorTitle = "\0";
    promptText = "\0";
    errorText = "\0";
    StringBuffer formulaString = new StringBuffer();
    CellReferenceHelper.getCellReference(c1,r1,formulaString);
    formulaString.append(':');
    CellReferenceHelper.getCellReference(c2,r2,formulaString);
    formula1String = formulaString.toString();
  }

  /**
   * Constructor called when creating a data validation from the API
   */
  public DVParser(double val1, double val2, Condition c)
  {
    copied = false;
    type = DECIMAL;
    errorStyle = STOP;
    condition = c;
    extendedCellsValidation = false;
    
    // the options
    stringListGiven = false;
    emptyCellsAllowed = true;
    suppressArrow = false;
    showPrompt = true;
    showError = true;

    promptTitle = "\0";
    errorTitle = "\0";
    promptText = "\0";
    errorText = "\0";
    formula1String = DECIMAL_FORMAT.format(val1);

    if (!Double.isNaN(val2))
    {
      formula2String = DECIMAL_FORMAT.format(val2);
    }
  }

  /**
   * Constructor called when doing a cell deep copy
   */
  public DVParser(DVParser copy)
  {
    copied = true;
    type = copy.type;
    errorStyle = copy.errorStyle;
    condition = copy.condition;
    stringListGiven = copy.stringListGiven;
    emptyCellsAllowed = copy.emptyCellsAllowed;
    suppressArrow = copy.suppressArrow;
    showPrompt = copy.showPrompt;
    showError = copy.showError;
    promptTitle = copy.promptTitle;
    promptText = copy.promptText;
    errorTitle = copy.errorTitle;
    errorText = copy.errorText;
    extendedCellsValidation = copy.extendedCellsValidation;

    row1 = copy.row1;
    row2 = copy.row2;
    column1 = copy.column1;
    column2 = copy.column2;

    // Don't copy the formula parsers - just take their string equivalents
    if (copy.formula1String != null)
    {
      formula1String = copy.formula1String;
      formula2String = copy.formula2String;
    }
    else
    {
      try
      {
        formula1String = copy.formula1.getFormula();
        formula2String = (copy.formula2 != null) ? 
          copy.formula2.getFormula() : null;
      }
      catch (FormulaException e)
      {
        logger.warn("Cannot parse validation formula:  " + e.getMessage());
      }
    }
    // Don't copy the cell references - these will be added later
  }

  /**
   * Gets the data
   */
  public byte[] getData()
  {
    // Compute the length of the data
    byte[] f1Bytes = formula1 != null ? formula1.getBytes() : new byte[0];
    byte[] f2Bytes = formula2 != null ? formula2.getBytes() : new byte[0];
    int dataLength = 
      4 + // the options
      promptTitle.length() * 2 + 3 + // the prompt title
      errorTitle.length() * 2 + 3 + // the error title
      promptText.length() * 2 + 3 + // the prompt text
      errorText.length() * 2 + 3 + // the error text
      f1Bytes.length + 2 + // first formula
      f2Bytes.length + 2 + // second formula
      + 4 + // unused bytes
      10; // cell range

    byte[] data = new byte[dataLength];

    // The position
    int pos = 0;

    // The options
    int options = 0;
    options |= type.getValue();
    options |= errorStyle.getValue() << 4;
    options |= condition.getValue() << 20;

    if (stringListGiven) 
    {
      options |= STRING_LIST_GIVEN_MASK;
    }

    if (emptyCellsAllowed) 
    {
      options |= EMPTY_CELLS_ALLOWED_MASK;
    }

    if (suppressArrow)
    {
      options |= SUPPRESS_ARROW_MASK;
    }

    if (showPrompt)
    {
      options |= SHOW_PROMPT_MASK;
    }

    if (showError)
    {
      options |= SHOW_ERROR_MASK;
    }

    // The text
    IntegerHelper.getFourBytes(options, data, pos);
    pos += 4;
    
    IntegerHelper.getTwoBytes(promptTitle.length(), data, pos);
    pos += 2;

    data[pos] = (byte) 0x1; // unicode indicator
    pos++;

    StringHelper.getUnicodeBytes(promptTitle, data, pos);
    pos += promptTitle.length() * 2;

    IntegerHelper.getTwoBytes(errorTitle.length(), data, pos);
    pos += 2;

    data[pos] = (byte) 0x1; // unicode indicator
    pos++;

    StringHelper.getUnicodeBytes(errorTitle, data, pos);
    pos += errorTitle.length() * 2;

    IntegerHelper.getTwoBytes(promptText.length(), data, pos);
    pos += 2;

    data[pos] = (byte) 0x1; // unicode indicator
    pos++;

    StringHelper.getUnicodeBytes(promptText, data, pos);
    pos += promptText.length() * 2;

    IntegerHelper.getTwoBytes(errorText.length(), data, pos);
    pos += 2;

    data[pos] = (byte) 0x1; // unicode indicator
    pos++;

    StringHelper.getUnicodeBytes(errorText, data, pos);
    pos += errorText.length() * 2;

    // Formula 1
    IntegerHelper.getTwoBytes(f1Bytes.length, data, pos);
    pos += 4;

    System.arraycopy(f1Bytes, 0, data, pos, f1Bytes.length);
    pos += f1Bytes.length;

    // Formula 2
    IntegerHelper.getTwoBytes(f2Bytes.length, data, pos);
    pos += 4;
    
    System.arraycopy(f2Bytes, 0, data, pos, f2Bytes.length);
    pos += f2Bytes.length;

    // The cell ranges
    IntegerHelper.getTwoBytes(1, data, pos);
    pos += 2;

    IntegerHelper.getTwoBytes(row1, data, pos);
    pos += 2;

    IntegerHelper.getTwoBytes(row2, data, pos);
    pos += 2;

    IntegerHelper.getTwoBytes(column1, data, pos);
    pos += 2;

    IntegerHelper.getTwoBytes(column2, data, pos);
    pos += 2;

    return data;
  }

  /**
   * Inserts a row
   *
   * @param row the row to insert
   */
  public void insertRow(int row)
  {
    if (formula1 != null)
    {
      formula1.rowInserted(0, row, true);
    }

    if (formula2 != null)
    {
      formula2.rowInserted(0, row, true);
    }

    if (row1 >= row)
    {
      row1++;
    }

    if (row2 >= row && row2 != MAX_ROWS)
    {
      row2++;
    }
  }

  /**
   * Inserts a column
   *
   * @param col the column to insert
   */
  public void insertColumn(int col)
  {
    if (formula1 != null)
    {
      formula1.columnInserted(0, col, true);
    }

    if (formula2 != null)
    {
      formula2.columnInserted(0, col, true);
    }

    if (column1 >= col)
    {
      column1++;
    }

    if (column2 >= col && column2 != MAX_COLUMNS)
    {
      column2++;
    }
  }

  /**
   * Removes a row
   *
   * @param row the row to insert
   */
  public void removeRow(int row)
  {
    if (formula1 != null)
    {
      formula1.rowRemoved(0, row, true);
    }

    if (formula2 != null)
    {
      formula2.rowRemoved(0, row, true);
    }

    if (row1 > row)
    {
      row1--;
    }

    if (row2 >= row)
    {
      row2--;
    }
  }

  /**
   * Removes a column
   *
   * @param col the row to remove
   */
  public void removeColumn(int col)
  {
    if (formula1 != null)
    {
      formula1.columnRemoved(0, col, true);
    }

    if (formula2 != null)
    {
      formula2.columnRemoved(0, col, true);
    }

    if (column1 > col)
    {
      column1--;
    }

    if (column2 >= col && column2 != MAX_COLUMNS)
    {
      column2--;
    }
  }

  /**
   * Accessor for first column
   *
   * @return the first column
   */
  public int getFirstColumn()
  {
    return column1;
  }

  /**
   * Accessor for the last column
   *
   * @return the last column
   */
  public int getLastColumn()
  {
    return column2;
  }

  /**
   * Accessor for first row
   *
   * @return the first row
   */
  public int getFirstRow()
  {
    return row1;
  }

  /**
   * Accessor for the last row
   *
   * @return the last row
   */
  public int getLastRow()
  {
    return row2;
  }

  /**
   * Gets the formula present in the validation
   *
   * @return the validation formula as a string
   * @exception FormulaException
   */
  String getValidationFormula() throws FormulaException
  {
    if (type == LIST)
    {
      return formula1.getFormula();
    }

    String s1 = formula1.getFormula();
    String s2 = formula2 != null ? formula2.getFormula() : null;
    return condition.getConditionString(s1, s2) + 
      "; x " + type.getDescription();
  }

  /**
   * Called by the cell value when the cell features are added to the sheet
   */
  public void setCell(int col, 
                      int row, 
                      ExternalSheet es, 
                      WorkbookMethods nt,
                      WorkbookSettings ws) throws FormulaException
  {
    // If this is part of an extended cells validation, then do nothing
    // as this will already have been called and parsed when the top left
    // cell was added
    if (extendedCellsValidation)
    {
      return;
    }

    row1 = row;
    row2 = row;
    column1 = col;
    column2 = col;

    formula1 = new FormulaParser(formula1String,
                                 es, nt, ws, 
                                 ParseContext.DATA_VALIDATION);
    formula1.parse();

    if (formula2String != null)
    {
      formula2 = new FormulaParser(formula2String,
                                   es, nt, ws, 
                                   ParseContext.DATA_VALIDATION);
      formula2.parse();
    }
  }

  /**
   * Indicates that the data validation extends across several more cells
   *
   * @param cols - the number of extra columns
   * @param rows - the number of extra rows
   */
  public void extendCellValidation(int cols, int rows)
  {
    row2 = row1 + rows;
    column2 = column1 + cols;
    extendedCellsValidation = true;
  }

  /**
   * Accessor which indicates whether this validation applies across
   * multiple cels
   */
  public boolean extendedCellsValidation()
  {
    return extendedCellsValidation;
  }

  public boolean copied()
  {
    return copied;
  }
}
