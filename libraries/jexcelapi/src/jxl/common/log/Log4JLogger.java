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

import org.apache.log4j.Logger;

/**
 * A logger which uses the log4j library from jakarta.  Each instance
 * of this class acts as a wrapper to the log4j Logger class
 */
public class Log4JLogger extends jxl.common.Logger
{
  /**
   * The log4j logger
   */
  private Logger log4jLogger;

  /**
   * Default constructor. This constructor is 
   */
  public Log4JLogger()
  {
    super();
  }

  /**
   * Constructor invoked by the getLoggerImpl method to return a logger
   * for a particular class
   */
  private Log4JLogger(Logger l)
  {
    super();
    log4jLogger = l;
  }

  /**
   *  Log a debug message
   */
  public void debug(Object message)
  {
    log4jLogger.debug(message);
  }

  /**
   * Log a debug message and exception
   */
  public void debug(Object message, Throwable t)
  {
    log4jLogger.debug(message, t);
  }

  /**
   *  Log an error message
   */
  public void error(Object message)
  {
    log4jLogger.error(message);
  }

  /**
   * Log an error message object and exception
   */
  public void error(Object message, Throwable t)
  {
    log4jLogger.error(message, t);
  }

  /**
   * Log a fatal message
   */
  public void fatal(Object message)
  {
    log4jLogger.fatal(message);
  }

  /**
   * Log a fatal message and exception
   */
  public void fatal(Object message, Throwable t)
  {
    log4jLogger.fatal(message,t);
  }

  /**
   * Log an information message
   */
  public void info(Object message)
  {
    log4jLogger.info(message);
  }

  /**
   * Logs an information message and an exception
   */

  public void info(Object message, Throwable t)
  {
    log4jLogger.info(message, t);
  }

  /**
   * Log a warning message object
   */
  public void warn(Object message)
  {
    log4jLogger.warn(message);
  }

  /**
   * Log a warning message with exception
   */
  public void warn(Object message, Throwable t)
  {
    log4jLogger.warn(message, t);
  }

  /**
   * Accessor to the logger implementation
   */
  protected jxl.common.Logger getLoggerImpl(Class cl)
  {
    Logger l = Logger.getLogger(cl);
    return new Log4JLogger(l);
  }
}
