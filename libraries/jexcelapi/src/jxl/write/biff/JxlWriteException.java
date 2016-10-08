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

import jxl.write.WriteException;

/**
 * Exception thrown when reading a biff file
 */
public class JxlWriteException extends WriteException
{
  private static class WriteMessage
  {
    /**
     */
    public String message;
    /**
     * Constructs this exception with the specified message
     * 
     * @param m the messageA
     */
    WriteMessage(String m) {message = m;}
  }

  /**
   */
  static WriteMessage formatInitialized = 
    new WriteMessage("Attempt to modify a referenced format");
  /**
   */
  static WriteMessage cellReferenced = 
    new WriteMessage("Cell has already been added to a worksheet");

  static WriteMessage maxRowsExceeded =
    new WriteMessage("The maximum number of rows permitted on a worksheet " +
                     "been exceeded");

  static WriteMessage maxColumnsExceeded =
    new WriteMessage("The maximum number of columns permitted on a " +
                     "worksheet has been exceeded");

  static WriteMessage copyPropertySets =
    new WriteMessage("Error encounted when copying additional property sets");

  /**
   * Constructs this exception with the specified message
   * 
   * @param m the message
   */
  public JxlWriteException(WriteMessage m)
  {
    super(m.message);
  }
}
