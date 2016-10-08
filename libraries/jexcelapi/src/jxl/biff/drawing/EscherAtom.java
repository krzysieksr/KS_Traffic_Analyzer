/*********************************************************************
*
*      Copyright (C) 2003 Andrew Khan
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

package jxl.biff.drawing;

import jxl.common.Logger;

/**
 * Class for atoms.  This may be instantiated as is for unknown/uncared about
 * atoms, or subclassed if we have some semantic interest in the contents
 */
class EscherAtom extends EscherRecord
{
  /**
   * The logger
   */
  private static Logger logger = Logger.getLogger(EscherAtom.class);

  /**
   * Constructor
   *
   * @param erd the escher record data
   */
  public EscherAtom(EscherRecordData erd)
  {
    super(erd);
  }

  /**
   * Constructor
   *
   * @param type the type
   */
  protected EscherAtom(EscherRecordType type)
  {
    super(type);
  }

  /**
   * Gets the data for writing
   *
   * @return the data
   */
  byte[] getData()
  {
    logger.warn("escher atom getData called on object of type " +
                getClass().getName() + " code " +
                Integer.toString(getType().getValue(), 16));
    return null;
  }
}
