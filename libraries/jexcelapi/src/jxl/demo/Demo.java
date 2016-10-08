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

package jxl.demo;

import java.io.File;
import java.io.FileOutputStream;
import java.io.OutputStream;

import jxl.common.Logger;

import jxl.Cell;
import jxl.Range;
import jxl.Workbook;

/**
 * The main demo class which interprets the command line switches in order
 * to determine how to call the demo programs
 * The demo program uses stdout as its default output stream
 */
public class Demo
{
  private static final int CSVFormat = 13;
  private static final int XMLFormat = 14;

  /**
   * The logger
   */
  private static Logger logger = Logger.getLogger(Demo.class);

  /**
   * Displays the acceptable command line arguments
   */
  private static void displayHelp()
  {
      System.err.println
        ("Command format:  Demo [-unicode] [-csv] [-hide] excelfile");
      System.err.println("                 Demo -xml [-format]  excelfile");
      System.err.println("                 Demo -readwrite|-rw excelfile output");
      System.err.println("                 Demo -biffdump | -bd | -wa | -write | -formulas | -features | -escher | -escherdg excelfile");
      System.err.println("                 Demo -ps excelfile [property] [output]");
      System.err.println("                 Demo -version | -logtest | -h | -help");

  }

  /**
   * The main method.  Gets the worksheet and then uses the API 
   * within a simple loop to print out the spreadsheet contents as
   * comma separated values
   * 
   * @param args the command line arguments
   */
  public static void main(String args[])
  {
    if (args.length == 0)
    {
      displayHelp();
      System.exit(1);
    }

    if (args[0].equals("-help") || args[0].equals("-h"))
    {
      displayHelp();
      System.exit(1);
    }

    if (args[0].equals("-version"))
    {
      System.out.println("v"+Workbook.getVersion());
      System.exit(0);
    }

    if (args[0].equals("-logtest"))
    {
      logger.debug("A sample \"debug\" message");
      logger.info("A sample \"info\" message");
      logger.warn("A sample \"warning\" message");
      logger.error("A sample \"error\" message");
      logger.fatal("A sample \"fatal\" message");
      System.exit(0);
    }

    boolean write = false;
    boolean readwrite = false;
    boolean formulas = false;
    boolean biffdump = false;
    boolean jxlversion = false;
    boolean propertysets = false;
    boolean features = false;
    boolean escher = false;
    boolean escherdg = false;
    String file = args[0];
    String outputFile = null;
    String propertySet = null;

    if (args[0].equals("-write"))
    {
      write = true;
      file = args[1];
    }
    else if (args[0].equals("-formulas"))
    {
      formulas = true;
      file = args[1];
    }
    else if (args[0].equals("-features"))
    {
      features = true;
      file = args[1];
    }
    else if (args[0].equals("-escher"))
    {
      escher = true;
      file = args[1];
    }
    else if (args[0].equals("-escherdg"))
    {
      escherdg = true;
      file = args[1];
    }
    else if (args[0].equals("-biffdump") || args[0].equals("-bd"))
    {
      biffdump = true;
      file = args[1];
    }
    else if (args[0].equals("-wa"))
    {
      jxlversion = true;
      file = args[1];
    }
    else if (args[0].equals("-ps"))
    {
      propertysets = true;
      file = args[1];

      if (args.length > 2)
      {
        propertySet = args[2];
      }

      if (args.length == 4)
      {
        outputFile = args[3];
      }
    }
    else if (args[0].equals("-readwrite") || args[0].equals("-rw"))
    {
      readwrite = true;
      file = args[1];
      outputFile = args[2];
    }
    else
    {
      file = args[args.length - 1];
    }
    
    String encoding = "UTF8";
    int format = CSVFormat;
    boolean formatInfo = false;
    boolean hideCells = false;

    if (write == false && 
        readwrite == false && 
        formulas == false &&
        biffdump == false && 
        jxlversion == false &&
        propertysets == false &&
        features == false && 
        escher == false &&
        escherdg == false)
    {
      for (int i = 0; i < args.length - 1; i++)
      {
        if (args[i].equals("-unicode"))
        {
          encoding="UnicodeBig";
        }
        else if (args[i].equals("-xml"))
        {
          format = XMLFormat;
        }
        else if (args[i].equals("-csv"))
        {
          format = CSVFormat;
        }
        else if (args[i].equals("-format"))
        {
          formatInfo = true;
        }
        else if (args[i].equals("-hide"))
        {
          hideCells = true;
        }
        else
        {
          System.err.println
            ("Command format:  CSV [-unicode] [-xml|-csv] excelfile");
          System.exit(1);
        }
      }
    }

    try
    {
      if (write)
      {
        Write w = new Write(file);
        w.write();
      }
      else if (readwrite)
      {
        ReadWrite rw = new ReadWrite(file, outputFile);
        rw.readWrite();
      }
      else if (formulas)
      {
        Workbook w = Workbook.getWorkbook(new File(file));
        Formulas f = new Formulas(w, System.out, encoding);
        w.close();
      }
      else if (features)
      {
        Workbook w = Workbook.getWorkbook(new File(file));
        Features f = new Features(w, System.out, encoding);
        w.close();
      }
      else if (escher)
      {
        Workbook w = Workbook.getWorkbook(new File(file));
        Escher f = new Escher(w, System.out, encoding);
        w.close();
      }
      else if (escherdg)
      {
        Workbook w = Workbook.getWorkbook(new File(file));
        EscherDrawingGroup f = new EscherDrawingGroup(w, System.out, encoding);
        w.close();
      }
      else if (biffdump)
      {
        BiffDump bd = new BiffDump(new File(file), System.out);
      }
      else if (jxlversion)
      {
        WriteAccess bd = new WriteAccess(new File(file));
      }
      else if (propertysets)
      {
        OutputStream os = System.out;
        if (outputFile != null)
        {
          os = new FileOutputStream(outputFile);
        }
        PropertySetsReader psr = new PropertySetsReader(new File(file),
                                                        propertySet,
                                                        os);
      }
      else
      {
        Workbook w = Workbook.getWorkbook(new File(file));
        
        //        findTest(w);

        if (format == CSVFormat)
        {  
          CSV csv = new CSV(w, System.out, encoding, hideCells);
        }
        else if (format == XMLFormat)
        {
          XML xml = new XML(w, System.out, encoding, formatInfo);
        }
        
        w.close();
      }
    }
    catch (Throwable t)
    {
      System.out.println(t.toString());
      t.printStackTrace();
    }
  }

  /**
   * A private method to test the various find functions
   */
  private static void findTest(Workbook w)
  {
    logger.info("Find test");

    Cell c = w.findCellByName("named1");
    if (c != null)
    {
      logger.info("named1 contents:  " + c.getContents());
    }

    c = w.findCellByName("named2");
    if (c != null)
    {
      logger.info("named2 contents:  " + c.getContents());
    }

    c = w.findCellByName("namedrange");
    if (c != null)
    {
      logger.info("named2 contents:  " + c.getContents());
    }

    Range[] range = w.findByName("namedrange");
    if (range != null)
    {
      c = range[0].getTopLeft();
      logger.info("namedrange top left contents:  " + c.getContents());

      c = range[0].getBottomRight();
      logger.info("namedrange bottom right contents:  " + c.getContents());
    }

    range = w.findByName("nonadjacentrange");
    if (range != null)
    {
      for (int i = 0; i < range.length; i++)
      {
        c = range[i].getTopLeft();
        logger.info("nonadjacent top left contents:  " + c.getContents());
        
        c = range[i].getBottomRight();
        logger.info("nonadjacent bottom right contents:  " + c.getContents());
      }
    }

    range = w.findByName("horizontalnonadjacentrange");
    if (range != null)
    {
      for (int i = 0; i < range.length; i++)
      {
        c = range[i].getTopLeft();
        logger.info("horizontalnonadjacent top left contents:  " + 
                           c.getContents());
        
        c = range[i].getBottomRight();
        logger.info("horizontalnonadjacent bottom right contents:  " + 
                 c.getContents());
      }
    }

  }
}



