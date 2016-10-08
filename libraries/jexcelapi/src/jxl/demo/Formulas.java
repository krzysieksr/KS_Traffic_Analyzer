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

import java.io.BufferedWriter;
import java.io.IOException;
import java.io.OutputStream;
import java.io.OutputStreamWriter;
import java.io.UnsupportedEncodingException;

import java.util.ArrayList;
import java.util.Iterator;

import jxl.Cell;
import jxl.CellType;
import jxl.FormulaCell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.biff.CellReferenceHelper;
import jxl.biff.formula.FormulaException;

/**
 * Goes through each cell in the workbook, and if the contents of that
 * cell is a formula, it prints out the last calculated value and
 * the formula string
 */
public class Formulas
{
  /**
   * Constructor
   *
   * @param w The workbook to interrogate
   * @param out The output stream to which the CSV values are written
   * @param encoding The encoding used by the output stream.  Null or 
   * unrecognized values cause the encoding to default to UTF8
   * @exception java.io.IOException
   */
  public Formulas(Workbook w, OutputStream out, String encoding)
    throws IOException
  {
    if (encoding == null || !encoding.equals("UnicodeBig"))
    {
      encoding = "UTF8";
    }

    try
    {
      OutputStreamWriter osw = new OutputStreamWriter(out, encoding);
      BufferedWriter bw = new BufferedWriter(osw);

      ArrayList parseErrors = new ArrayList();
      
      for (int sheet = 0; sheet < w.getNumberOfSheets(); sheet++)
      {
        Sheet s = w.getSheet(sheet);

        bw.write(s.getName());
        bw.newLine();
      
        Cell[] row = null;
        Cell c = null;
      
        for (int i = 0 ; i < s.getRows() ; i++)
        {
          row = s.getRow(i);

          for (int j = 0; j < row.length; j++)
          {
            c = row[j];
            if (c.getType() == CellType.NUMBER_FORMULA || 
                c.getType() == CellType.STRING_FORMULA || 
                c.getType() == CellType.BOOLEAN_FORMULA ||
                c.getType() == CellType.DATE_FORMULA ||
                c.getType() == CellType.FORMULA_ERROR)
            {
              FormulaCell nfc = (FormulaCell) c;
              StringBuffer sb = new StringBuffer();
              CellReferenceHelper.getCellReference
                 (c.getColumn(), c.getRow(), sb);

              try
              {
                bw.write("Formula in "  + sb.toString() + 
                         " value:  " + c.getContents());
                bw.flush();
                bw.write(" formula: " + nfc.getFormula());
                bw.flush();
                bw.newLine();
              }
              catch (FormulaException e)
              {
                bw.newLine();
                parseErrors.add(s.getName() + '!' +
                                sb.toString() + ": " + e.getMessage());
              }
            }
          }
        }
      }
      bw.flush();
      bw.close();

      if (parseErrors.size() > 0)
      {
        System.err.println();
        System.err.println("There were " + parseErrors.size() + " errors");

        Iterator i = parseErrors.iterator();
        while (i.hasNext())
        {
          System.err.println(i.next());
        }
      }
    }
    catch (UnsupportedEncodingException e)
    {
      System.err.println(e.toString());
    }
  }

}

