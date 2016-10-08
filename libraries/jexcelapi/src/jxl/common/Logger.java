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

package jxl.common;

import java.security.AccessControlException;

/**
 * Abstract wrapper class for the logging interface of choice.  
 * The methods declared here are the same as those for the log4j  
 */
public abstract class Logger
{
  /**
   * The singleton logger
   */
  private static Logger logger = null;

  /**
   * Factory method to return the logger
   */
  public static final Logger getLogger(Class cl)
  {
    if (logger == null)
    {
      initializeLogger();
    }

    return logger.getLoggerImpl(cl);
  }

  /**
   * Initializes the logger in a thread safe manner
   */
  private synchronized static void initializeLogger()
  {
    if (logger != null)
    {
      return;
    }

    String loggerName = jxl.common.log.LoggerName.NAME;

    try
    {
      // First see if there was anything defined at run time
      loggerName = System.getProperty("logger");

      if (loggerName == null)
      {
        // Get the logger name from the compiled in logger 
        loggerName = jxl.common.log.LoggerName.NAME;
      }

      logger = (Logger) Class.forName(loggerName).newInstance();
    }
    catch(IllegalAccessException e)
    {
      logger = new jxl.common.log.SimpleLogger();
      logger.warn("Could not instantiate logger " + loggerName + 
                  " using default");
    }
    catch(InstantiationException e)
    {
      logger = new jxl.common.log.SimpleLogger();
      logger.warn("Could not instantiate logger " + loggerName + 
                  " using default");
    }
    catch (AccessControlException e)
    {
      logger = new jxl.common.log.SimpleLogger();
      logger.warn("Could not instantiate logger " + loggerName + 
                  " using default");
    }
    catch(ClassNotFoundException e)
    {
      logger = new jxl.common.log.SimpleLogger();
      logger.warn("Could not instantiate logger " + loggerName + 
                  " using default");
    }
  }

  /**
   * Constructor
   */
  protected Logger()
  {
  }

  /**
   *  Log a debug message
   */
  public abstract void debug(Object message);

  /**
   * Log a debug message and exception
   */
  public abstract void debug(Object message, Throwable t);

  /**
   *  Log an error message
   */
  public abstract void error(Object message);

  /**
   * Log an error message object and exception
   */
  public abstract void error(Object message, Throwable t);

  /**
   * Log a fatal message
   */
  public abstract void fatal(Object message);

  /**
   * Log a fatal message and exception
   */
  public abstract void fatal(Object message, Throwable t);

  /**
   * Log an information message
   */
  public abstract void info(Object message);

  /**
   * Logs an information message and an exception
   */
  public abstract void info(Object message, Throwable t);

  /**
   * Log a warning message object
   */
  public abstract void warn(Object message);

  /**
   * Log a warning message with exception
   */
  public abstract void warn(Object message, Throwable t);

  /**
   * Accessor to the logger implementation
   */
  protected abstract Logger getLoggerImpl(Class cl);

  /**
   * Empty implementation of the suppressWarnings.  Subclasses may 
   * or may not override this method.  This method is included
   * primarily for backwards support of the jxl.nowarnings property, and
   * is used only by the SimpleLogger
   *
   * @param w suppression flag
   */
  public void setSuppressWarnings(boolean w)
  {
    // default implementation does nothing
  }
}
