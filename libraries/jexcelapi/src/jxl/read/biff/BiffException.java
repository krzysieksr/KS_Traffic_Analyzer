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

import jxl.JXLException;

/**
 * Exception thrown when reading a biff file
 */
public class BiffException extends JXLException
{
  /**
   * Inner class containing the various error messages
   */
  private static class BiffMessage
  {
    /**
     * The formatted message
     */
    public String message;
    /**
     * Constructs this exception with the specified message
     *
     * @param m the messageA
     */
    BiffMessage(String m)
    {
      message = m;
    }
  }

  /**
   */
  static final BiffMessage unrecognizedBiffVersion =
    new BiffMessage("Unrecognized biff version");

  /**
   */
  static final BiffMessage expectedGlobals =
    new BiffMessage("Expected globals");

  /**
   */
  static final BiffMessage excelFileTooBig =
    new BiffMessage("Not all of the excel file could be read");

  /**
   */
  static final BiffMessage excelFileNotFound =
    new BiffMessage("The input file was not found");

  /**
   */
  static final BiffMessage unrecognizedOLEFile =
    new BiffMessage("Unable to recognize OLE stream");

  /**
   */
  static final BiffMessage streamNotFound =
    new BiffMessage("Compound file does not contain the specified stream");

  /**
   */
  static final BiffMessage passwordProtected =
    new BiffMessage("The workbook is password protected");

  /**
   */
  static final BiffMessage corruptFileFormat =
    new BiffMessage("The file format is corrupt");

  /**
   * Constructs this exception with the specified message
   *
   * @param m the message
   */
  public BiffException(BiffMessage m)
  {
    super(m.message);
  }
}
