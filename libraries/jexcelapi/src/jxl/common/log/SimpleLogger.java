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

package jxl.common.log;

import jxl.common.Logger;

/**
 * The default logger.  Simple writes everything out to stdout or stderr
 */
public class SimpleLogger extends Logger
{
  /**
   * Flag to indicate whether or not warnings should be suppressed
   */
  private boolean suppressWarnings;

  /**
   * Constructor
   */
  public SimpleLogger()
  {
    suppressWarnings = false;
  }

  /**
   *  Log a debug message
   */
  public void debug(Object message)
  {
    if (!suppressWarnings)
    {
      System.out.print("Debug: ");
      System.out.println(message);
    }
  }

  /**
   * Log a debug message and exception
   */
  public void debug(Object message, Throwable t)
  {
    if (!suppressWarnings)
    {
      System.out.print("Debug: ");
      System.out.println(message);
      t.printStackTrace();
    }
  }

  /**
   *  Log an error message
   */
  public void error(Object message)
  {
    System.err.print("Error: ");
    System.err.println(message);
  }

  /**
   * Log an error message object and exception
   */
  public void error(Object message, Throwable t)
  {
    System.err.print("Error: ");
    System.err.println(message);
    t.printStackTrace();
  }

  /**
   * Log a fatal message
   */
  public void fatal(Object message)
  {
    System.err.print("Fatal: ");
    System.err.println(message);
  }

  /**
   * Log a fatal message and exception
   */
  public void fatal(Object message, Throwable t)
  {
    System.err.print("Fatal:  ");
    System.err.println(message);
    t.printStackTrace();
  }

  /**
   * Log an information message
   */
  public void info(Object message)
  {
    if (!suppressWarnings)
    {
      System.out.println(message);
    }
  }

  /**
   * Logs an information message and an exception
   */

  public void info(Object message, Throwable t)
  {
    if (!suppressWarnings)
    {
      System.out.println(message);
      t.printStackTrace();
    }
  }

  /**
   * Log a warning message object
   */
  public void warn(Object message)
  {
    if (!suppressWarnings)
    {
      System.err.print("Warning:  ");
      System.err.println(message);
    }
  }

  /**
   * Log a warning message with exception
   */
  public void warn(Object message, Throwable t)
  {
    if (!suppressWarnings)
    {
      System.err.print("Warning:  ");
      System.err.println(message);
      t.printStackTrace();
    }
  }

  /**
   * Accessor to the logger implementation
   */
  protected Logger getLoggerImpl(Class c)
  {
    return this;
  }

  /**
   * Overrides the method in the base class to suppress warnings - it can
   * be set using the system property jxl.nowarnings.  
   * This method was originally present in the WorkbookSettings bean,
   * but has been moved to the logger class.  This means it is now present
   * when the JVM is initialized, and subsequent to change it on 
   * a Workbook by Workbook basis will prove fruitless
   *
   * @param w suppression flag
   */
  public void setSuppressWarnings(boolean w)
  {
    suppressWarnings = w;
  }
}
