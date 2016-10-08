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

package jxl.biff.formula;

import jxl.common.Logger;

import jxl.WorkbookSettings;

/**
 * Class used to hold a function when reading it in from a string.  At this
 * stage it is unknown whether it is a BuiltInFunction or a VariableArgFunction
 */
class StringFunction extends StringParseItem
{
  /**
   * The logger
   */
  private static Logger logger = Logger.getLogger(StringFunction.class);

  /**
   * The function
   */
  private Function function;

  /**
   * The function string
   */
  private String functionString;

  /**
   * Constructor
   *
   * @param s the lexically parsed stirng
   */
  StringFunction(String s)
  {
    functionString = s.substring(0, s.length() - 1); 
  }

  /**
   * Accessor for the function
   *
   * @param ws the workbook settings
   * @return the function
   */
  Function getFunction(WorkbookSettings ws)
  {
    if (function == null)
    {
      function = Function.getFunction(functionString, ws);
    }
    return function;
  }
}
