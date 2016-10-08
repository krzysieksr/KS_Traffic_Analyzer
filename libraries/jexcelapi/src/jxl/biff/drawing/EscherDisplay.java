/*********************************************************************
*
*      Copyright (C) 2005 Andrew Khan
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

import java.io.BufferedWriter;
import java.io.IOException;

/**
 * Class used to display a complete hierarchically organized Escher stream
 * The whole thing is dumped to System.out
 *
 * This class is only used as a debugging tool
 */
public class EscherDisplay
{
  /**
   * The escher stream
   */
  private EscherStream stream;

  /**
   * The writer
   */
  private BufferedWriter writer;

  /**
   * Constructor
   *
   * @param s the stream
   * @param bw the writer
   */
  public EscherDisplay(EscherStream s, BufferedWriter bw)
  {
    stream = s;
    writer = bw;
  }

  /**
   * Display the formatted escher stream
   *
   * @exception IOException
   */
  public void display() throws IOException
  {
    EscherRecordData er = new EscherRecordData(stream, 0);
    EscherContainer ec = new EscherContainer(er);
    displayContainer(ec, 0);
  }

  /**
   * Displays the escher container as text
   *
   * @param ec the escher container
   * @param level the indent level
   * @exception IOException
   */
  private void displayContainer(EscherContainer ec, int level)
    throws IOException
  {
    displayRecord(ec, level);

    // Display the contents of the container
    level++;

    EscherRecord[] children = ec.getChildren();

    for (int i = 0; i < children.length; i++)
    {
      EscherRecord er = children[i];
      if (er.getEscherData().isContainer())
      {
        displayContainer((EscherContainer) er, level);
      }
      else
      {
        displayRecord(er, level);
      }
    }
  }

  /**
   * Displays an escher record
   *
   * @param er the record to display
   * @param level the amount of indentation
   * @exception IOException
   */
  private void displayRecord(EscherRecord er, int level)
    throws IOException
  {
    indent(level);

    EscherRecordType type = er.getType();

    // Display the code
    writer.write(Integer.toString(type.getValue(), 16));
    writer.write(" - ");

    // Display the name
    if (type == EscherRecordType.DGG_CONTAINER)
    {
      writer.write("Dgg Container");
      writer.newLine();
    }
    else if (type == EscherRecordType.BSTORE_CONTAINER)
    {
      writer.write("BStore Container");
      writer.newLine();
    }
    else if (type == EscherRecordType.DG_CONTAINER)
    {
      writer.write("Dg Container");
      writer.newLine();
    }
    else if (type == EscherRecordType.SPGR_CONTAINER)
    {
      writer.write("Spgr Container");
      writer.newLine();
    }
    else if (type == EscherRecordType.SP_CONTAINER)
    {
      writer.write("Sp Container");
      writer.newLine();
    }
    else if (type == EscherRecordType.DGG)
    {
      writer.write("Dgg");
      writer.newLine();
    }
    else if (type == EscherRecordType.BSE)
    {
      writer.write("Bse");
      writer.newLine();
    }
    else if (type == EscherRecordType.DG)
    {
      Dg dg = new Dg(er.getEscherData());
      writer.write("Dg:  drawing id " + dg.getDrawingId() + 
                   " shape count " + dg.getShapeCount());
      writer.newLine();
    }
    else if (type == EscherRecordType.SPGR)
    {
      writer.write("Spgr");
      writer.newLine();
    }
    else if (type == EscherRecordType.SP)
    {
      Sp sp = new Sp(er.getEscherData());
      writer.write("Sp:  shape id " + sp.getShapeId() + 
                   " shape type " + sp.getShapeType());
      writer.newLine();
    }
    else if (type == EscherRecordType.OPT)
    {
      Opt opt = new Opt(er.getEscherData());
      Opt.Property p260 = opt.getProperty(260);
      Opt.Property p261 = opt.getProperty(261);
      writer.write("Opt (value, stringValue): ");
      if (p260 != null)
      {
        writer.write("260: " + 
                     p260.value + ", " + 
                     p260.stringValue +
                     ";");
      }
      if (p261 != null)
      {
        writer.write("261: " + 
                     p261.value + ", " + 
                     p261.stringValue +
                     ";");
      }
      writer.newLine();
    }
    else if (type == EscherRecordType.CLIENT_ANCHOR)
    {
      writer.write("Client Anchor");
      writer.newLine();
    }
    else if (type == EscherRecordType.CLIENT_DATA)
    {
      writer.write("Client Data");
      writer.newLine();
    }
    else if (type == EscherRecordType.CLIENT_TEXT_BOX)
    {
      writer.write("Client Text Box");
      writer.newLine();
    }
    else if (type == EscherRecordType.SPLIT_MENU_COLORS)
    {
      writer.write("Split Menu Colors");
      writer.newLine();
    }
    else
    {
      writer.write("???");
      writer.newLine();
    }
  }

  /**
   * Indents to the amount specified by the level
   *
   * @param level the level
   * @exception IOException
   */
  private void indent(int level) throws IOException
  {
    for (int i = 0; i < level * 2; i++)
    {
      writer.write(' ');
    }
  }
}
