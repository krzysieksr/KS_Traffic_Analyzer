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

import java.util.Collection;

import jxl.common.Assert;
import jxl.common.Logger;

import jxl.CellReferenceHelper;
import jxl.Range;
import jxl.biff.drawing.ComboBox;
import jxl.biff.drawing.Comment;
import jxl.write.biff.CellValue;

/**
 * Container for any additional cell features
 */
public class BaseCellFeatures
{
  /**
   * The logger
   */
  public static Logger logger = Logger.getLogger(BaseCellFeatures.class);

  /**
   * The comment
   */
  private String comment;

  /**
   * The comment width in cells
   */
  private double commentWidth;

  /**
   * The comment height in cells
   */
  private double commentHeight;
  
  /**
   * A handle to the drawing object
   */
  private Comment commentDrawing;

  /**
   * A handle to the combo box object
   */
  private ComboBox comboBox;

  /**
   * The data validation settings
   */
  private DataValiditySettingsRecord validationSettings;

  /**
   * The DV Parser used to contain the validation details
   */
  private DVParser dvParser;

  /**
   * Indicates whether a drop down is required
   */
  private boolean dropDown;

  /**
   * Indicates whether this cell features has data validation
   */
  private boolean dataValidation;

  /**
   * The cell to which this is attached, and which may need to be notified
   */
  private CellValue writableCell;

  // Constants
  private final static double defaultCommentWidth = 3;
  private final static double defaultCommentHeight = 4;

  // Validation conditions
  protected static class ValidationCondition
  {
    private DVParser.Condition condition;
    
    private static ValidationCondition[] types = new ValidationCondition[0];
   
    ValidationCondition(DVParser.Condition c) 
    {
      condition = c;
      ValidationCondition[] oldtypes = types;
      types = new ValidationCondition[oldtypes.length+1];
      System.arraycopy(oldtypes, 0, types, 0, oldtypes.length);
      types[oldtypes.length] = this;
    }

    public DVParser.Condition getCondition()
    {
      return condition;
    }
  }

  public static final ValidationCondition BETWEEN = 
    new ValidationCondition(DVParser.BETWEEN);
  public static final ValidationCondition NOT_BETWEEN = 
    new ValidationCondition(DVParser.NOT_BETWEEN);
  public static final ValidationCondition EQUAL = 
    new ValidationCondition(DVParser.EQUAL);
  public static final ValidationCondition NOT_EQUAL = 
    new ValidationCondition(DVParser.NOT_EQUAL);
  public static final ValidationCondition GREATER_THAN = 
    new ValidationCondition(DVParser.GREATER_THAN);
  public static final ValidationCondition LESS_THAN = 
    new ValidationCondition(DVParser.LESS_THAN);
  public static final ValidationCondition GREATER_EQUAL = 
    new ValidationCondition(DVParser.GREATER_EQUAL);
  public static final ValidationCondition LESS_EQUAL = 
    new ValidationCondition(DVParser.LESS_EQUAL);

  /**
   * Constructor
   */
  protected BaseCellFeatures()
  {
  }

  /**
   * Copy constructor
   *
   * @param the cell to copy
   */
  public BaseCellFeatures(BaseCellFeatures cf)
  {
    // The comment stuff
    comment = cf.comment;
    commentWidth = cf.commentWidth;
    commentHeight = cf.commentHeight;

    // The data validation stuff.  
    dropDown = cf.dropDown;
    dataValidation = cf.dataValidation;

    validationSettings = cf.validationSettings; // ?

    if (cf.dvParser != null)
    {
      dvParser = new DVParser(cf.dvParser);
    }
  }

  /**
   * Accessor for the cell comment
   */
  protected String getComment()
  {
    return comment;
  }

  /**
   * Accessor for the comment width
   */
  public double getCommentWidth()
  {
    return commentWidth;
  }

  /**
   * Accessor for the comment height
   */
  public double getCommentHeight()
  {
    return commentHeight;
  }

  /** 
   * Called by the cell when the features are added
   *
   * @param wc the writable cell
   */
  public final void setWritableCell(CellValue wc)
  {
    writableCell = wc;
  } 

  /**
   * Internal method to set the cell comment.  Used when reading
   */
  public void setReadComment(String s, double w, double h)
  {
    comment = s;
    commentWidth = w;
    commentHeight = h;
  }

  /**
   * Internal method to set the data validation.  Used when reading
   */
  public void setValidationSettings(DataValiditySettingsRecord dvsr)
  {
    Assert.verify(dvsr != null);

    validationSettings = dvsr;
    dataValidation = true;
  }

  /**
   * Sets the cell comment
   *
   * @param s the comment
   */
  public void setComment(String s)
  {
    setComment(s, defaultCommentWidth, defaultCommentHeight);
  }

  /**
   * Sets the cell comment
   *
   * @param s the comment
   * @param height the height of the comment box in cells
   * @param width the width of the comment box in cells
   */
  public void setComment(String s, double width, double height)
  {
    comment = s;
    commentWidth = width;
    commentHeight = height;

    if (commentDrawing != null)
    {
      commentDrawing.setCommentText(s);
      commentDrawing.setWidth(width);
      commentDrawing.setWidth(height);
      // commentDrawing is set up when trying to modify a copied cell
    }
  }

  /**
   * Removes the cell comment, if present
   */
  public void removeComment()
  {
    // Set the comment string to be empty
    comment = null;

    // Remove the drawing from the drawing group
    if (commentDrawing != null)
    {
      // do not call DrawingGroup.remove() because comments are not present
      // on the Workbook DrawingGroup record
      writableCell.removeComment(commentDrawing);
      commentDrawing = null;
    }
  }

  /**
   * Public function which removes any data validation, if present
   */
  public void removeDataValidation()
  {
    if (!dataValidation)
    {
      return;
    }

    // If the data validation is shared, then generate a warning
    DVParser dvp = getDVParser();
    if (dvp.extendedCellsValidation())
    {
      logger.warn("Cannot remove data validation from " + 
                  CellReferenceHelper.getCellReference(writableCell) + 
                  " as it is part of the shared reference " +
                  CellReferenceHelper.getCellReference(dvp.getFirstColumn(),
                                                       dvp.getFirstRow()) +
                  "-" +
                  CellReferenceHelper.getCellReference(dvp.getLastColumn(),
                                                       dvp.getLastRow()));
      return;
    }

    // Remove the validation from the WritableSheet object if present
    writableCell.removeDataValidation();
    clearValidationSettings();
  }

  /**
   * Internal function which removes any data validation, including
   * shared ones, if present.  This is called from WritableSheetImpl
   * in response to a call to removeDataValidation
   */
  public void removeSharedDataValidation()
  {
    if (!dataValidation)
    {
      return;
    }

    // Remove the validation from the WritableSheet object if present
    writableCell.removeDataValidation();
    clearValidationSettings();
  }

  /**
   * Sets the comment drawing object
   */
  public final void setCommentDrawing(Comment c)
  {
    commentDrawing = c;
  }

  /**
   * Accessor for the comment drawing
   */
  public final Comment getCommentDrawing()
  {
    return commentDrawing;
  }

  /**
   * Gets the data validation list as a formula.  Used only when reading
   *
   * @return the validation formula as a list
   */
  public String getDataValidationList()
  {
    if (validationSettings == null)
    {
      return null;
    }

    return validationSettings.getValidationFormula();
  }

  /**
   * The list of items to validate for this cell.  For each object in the 
   * collection, the toString() method will be called and the data entered
   * will be validated against that string
   *
   * @param c the list of valid values
   */
  public void setDataValidationList(Collection c)
  {
    if (dataValidation && getDVParser().extendedCellsValidation())
    {
      logger.warn("Cannot set data validation on " + 
                  CellReferenceHelper.getCellReference(writableCell) + 
                  " as it is part of a shared data validation");
      return;
    }
    clearValidationSettings();
    dvParser = new DVParser(c);
    dropDown = true;
    dataValidation = true;
  }

  /**
   * The list of items to validate for this cell in the form of a cell range.
   *
   * @param c the list of valid values
   */
  public void setDataValidationRange(int col1, int r1, int col2, int r2)
  {
    if (dataValidation && getDVParser().extendedCellsValidation())
    {
      logger.warn("Cannot set data validation on " + 
                  CellReferenceHelper.getCellReference(writableCell) + 
                  " as it is part of a shared data validation");
      return;
    }
    clearValidationSettings();
    dvParser = new DVParser(col1, r1, col2, r2);
    dropDown = true;
    dataValidation = true;
  }

  /**
   * Sets the data validation based upon a named range
   */
  public void setDataValidationRange(String namedRange)
  {
    if (dataValidation && getDVParser().extendedCellsValidation())
    {
      logger.warn("Cannot set data validation on " + 
                  CellReferenceHelper.getCellReference(writableCell) + 
                  " as it is part of a shared data validation");
      return;
    }
    clearValidationSettings();
    dvParser = new DVParser(namedRange);
    dropDown = true;
    dataValidation = true;
  }

  /**
   * Sets the data validation based upon a numerical condition
   */
  public void setNumberValidation(double val, ValidationCondition c)
  {
    if (dataValidation && getDVParser().extendedCellsValidation())
    {
      logger.warn("Cannot set data validation on " + 
                  CellReferenceHelper.getCellReference(writableCell) + 
                  " as it is part of a shared data validation");
      return;
    }
    clearValidationSettings();
    dvParser = new DVParser(val, Double.NaN, c.getCondition());
    dropDown = false;
    dataValidation = true;
  }

  public void setNumberValidation(double val1, double val2, 
                                  ValidationCondition c)
  {
    if (dataValidation && getDVParser().extendedCellsValidation())
    {
      logger.warn("Cannot set data validation on " + 
                  CellReferenceHelper.getCellReference(writableCell) + 
                  " as it is part of a shared data validation");
      return;
    }
    clearValidationSettings();
    dvParser = new DVParser(val1, val2, c.getCondition());
    dropDown = false;
    dataValidation = true;
  }

  /**
   * Accessor for the data validation
   *
   * @return TRUE if this has a data validation associated with it,
             FALSE otherwise
  */
  public boolean hasDataValidation()
  {
    return dataValidation;
  }

  /**
   * Clears out any existing validation settings
   */
  private void clearValidationSettings()
  {
    validationSettings = null;
    dvParser = null;
    dropDown = false;
    comboBox = null;
    dataValidation = false;
  }

  /**
   * Accessor for whether a drop down is required
   *
   * @return TRUE if this requires a drop down, FALSE otherwise
   */
  public boolean hasDropDown()
  {
    return dropDown;
  }

  /**
   * Sets the combo box drawing object for list validations
   *
   * @param cb the combo box
   */
  public void setComboBox(ComboBox cb)
  {
    comboBox = cb;
  }

  /**
   * Gets the dv parser
   */
  public DVParser getDVParser()
  {
    // straightforward - this was created as  a writable cell
    if (dvParser != null)
    {
      return dvParser;
    }

    // this was copied from a readable cell, and then copied again
    if (validationSettings != null)
    {
      dvParser = new DVParser(validationSettings.getDVParser());
      return dvParser;
    }

    return null; // keep the compiler happy
  }

  /**
   * Use the same data validation logic as the specified cell features
   *
   * @param cf the data validation to reuse
   */
  public void shareDataValidation(BaseCellFeatures source)
  {
    if (dataValidation)
    {
      logger.warn("Attempting to share a data validation on cell " + 
                  CellReferenceHelper.getCellReference(writableCell) + 
                  " which already has a data validation");
      return;
    }
    clearValidationSettings();
    dvParser = source.getDVParser();
    validationSettings = null;
    dataValidation = true;
    dropDown = source.dropDown;
    comboBox = source.comboBox;
  }

  /**
   * Gets the range of cells to which the data validation applies.  If the
   * validation applies to just this cell, this will be reflected in the 
   * returned range
   *
   * @return the range to which the same validation extends, or NULL if this
   *         cell doesn't have a validation
   */
  public Range getSharedDataValidationRange()
  {
    if (!dataValidation)
    {
      return null;
    }
    
    DVParser dvp = getDVParser();

    return new SheetRangeImpl(writableCell.getSheet(),
                              dvp.getFirstColumn(),
                              dvp.getFirstRow(),
                              dvp.getLastColumn(),
                              dvp.getLastRow());
  }
}
