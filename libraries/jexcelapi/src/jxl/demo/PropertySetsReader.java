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
import java.io.FileInputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.io.OutputStreamWriter;

import jxl.WorkbookSettings;
import jxl.biff.BaseCompoundFile;
import jxl.read.biff.BiffException;
import jxl.read.biff.CompoundFile;


/**
 * Generates a biff dump of the specified excel file
 */
class PropertySetsReader
{
  private BufferedWriter writer;
  private CompoundFile compoundFile;

  /**
   * Constructor
   *
   * @param file the file
   * @param propertySet the property set to read
   * @param os the output stream
   * @exception IOException 
   * @exception BiffException
   */
  public PropertySetsReader(java.io.File file, String propertySet, 
                            OutputStream os) 
    throws IOException, BiffException
  {
    writer = new BufferedWriter(new OutputStreamWriter(os));
    FileInputStream fis = new FileInputStream(file);

    int initialFileSize = 1024*1024; // 1mb
    int arrayGrowSize = 1024*1024;// 1mb

    byte[] d = new byte[initialFileSize];  
    int bytesRead = fis.read(d);
    int pos = bytesRead;

    while (bytesRead != -1)
    {
      if (pos >= d.length)
      {
        // Grow the array
        byte newArray[] = new byte[d.length + arrayGrowSize];
        System.arraycopy(d, 0, newArray, 0, d.length);
        d = newArray;
      }
      bytesRead = fis.read(d, pos, d.length - pos);
      pos += bytesRead;
    }

    bytesRead = pos + 1;

    compoundFile = new CompoundFile(d, new WorkbookSettings());
    fis.close();

    if (propertySet == null)
    {
      displaySets();
    }
    else
    {
      displayPropertySet(propertySet, os);
    }
  }

  /**
   * Displays the properties to the output stream
   */
  void displaySets() throws IOException
  {
    int numSets = compoundFile.getNumberOfPropertySets();

    for (int i = 0; i < numSets ; i++)
    {
      BaseCompoundFile.PropertyStorage ps = compoundFile.getPropertySet(i);
      writer.write(Integer.toString(i));
      writer.write(") ");
      writer.write(ps.name);
      writer.write("(type ");
      writer.write(Integer.toString(ps.type));
      writer.write(" size ");
      writer.write(Integer.toString(ps.size));
      writer.write(" prev " );
      writer.write(Integer.toString(ps.previous));
      writer.write(" next " );
      writer.write(Integer.toString(ps.next));
      writer.write(" child " );
      writer.write(Integer.toString(ps.child));
      writer.write(" start block " );
      writer.write(Integer.toString(ps.startBlock));
      writer.write(")");
      writer.newLine();
    }

    writer.flush();
    writer.close();
  }

  /**
   * Write the property stream to the output stream
   */
  void displayPropertySet(String ps, OutputStream os) 
    throws IOException,BiffException
  {
    if (ps.equalsIgnoreCase("SummaryInformation"))
    {
      ps = BaseCompoundFile.SUMMARY_INFORMATION_NAME;
    }
    else if (ps.equalsIgnoreCase("DocumentSummaryInformation"))
    {
      ps = BaseCompoundFile.DOCUMENT_SUMMARY_INFORMATION_NAME;
    }
    else if (ps.equalsIgnoreCase("CompObj"))
    {
      ps = BaseCompoundFile.COMP_OBJ_NAME;
    }

    byte[] stream = compoundFile.getStream(ps);
    os.write(stream);
  }

}
