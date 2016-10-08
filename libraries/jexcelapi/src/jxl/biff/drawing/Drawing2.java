/*********************************************************************
*
*      Copyright (C) 2006 Andrew Khan
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

import java.io.FileInputStream;
import java.io.IOException;

import jxl.common.Assert;
import jxl.common.Logger;

import jxl.write.biff.File;


/**
 * Contains the various biff records used to insert a drawing into a
 * worksheet.  This type of image does not have an associated object
 * record
 */
public class Drawing2 implements DrawingGroupObject
{
  /**
   * The logger
   */
  private static Logger logger = Logger.getLogger(Drawing.class);

  /**
   * The spContainer that was read in
   */
  private EscherContainer readSpContainer;

  /**
   * The MsoDrawingRecord associated with the drawing
   */
  private MsoDrawingRecord msoDrawingRecord;

  /**
   * Initialized flag
   */
  private boolean initialized = false;

  /**
   * The file containing the image
   */
  private java.io.File imageFile;

  /**
   * The raw image data, used instead of an image file
   */
  private byte[] imageData;

  /**
   * The object id, assigned by the drawing group
   */
  private int objectId;

  /**
   * The blip id
   */
  private int blipId;

  /**
   * The column position of the image
   */
  private double x;

  /**
   * The row position of the image
   */
  private double y;

  /**
   * The width of the image in cells
   */
  private double width;

  /**
   * The height of the image in cells
   */
  private double height;

  /**
   * The number of places this drawing is referenced
   */
  private int referenceCount;

  /**
   * The top level escher container
   */
  private EscherContainer escherData;

  /**
   * Where this image came from (read, written or a copy)
   */
  private Origin origin;

  /**
   * The drawing group for all the images
   */
  private DrawingGroup drawingGroup;

  /**
   * The drawing data
   */
  private DrawingData drawingData;

  /**
   * The type of this drawing object
   */
  private ShapeType type;

  /**
   * The shape id
   */
  private int shapeId;

  /**
   * The drawing position on the sheet
   */
  private int drawingNumber;


  /**
   * Constructor used when reading images
   *
   * @param mso the drawing record
   * @param dd the drawing data for all drawings on this sheet
   * @param dg the drawing group
   */
  public Drawing2(MsoDrawingRecord mso,
                  DrawingData dd,
                  DrawingGroup dg)
  {
    drawingGroup = dg;
    msoDrawingRecord = mso;
    drawingData = dd;
    initialized = false;
    origin = Origin.READ;
    // there is no drawing number associated with this drawing
    drawingData.addRawData(msoDrawingRecord.getData());
    drawingGroup.addDrawing(this);

    Assert.verify(mso != null);

    initialize();
  }

  /**
   * Copy constructor used to copy drawings from read to write
   *
   * @param dgo the drawing group object
   * @param dg the drawing group
   */
  protected Drawing2(DrawingGroupObject dgo, DrawingGroup dg)
  {
    Drawing2 d = (Drawing2) dgo;
    Assert.verify(d.origin == Origin.READ);
    msoDrawingRecord = d.msoDrawingRecord;
    initialized = false;
    origin = Origin.READ;
    drawingData = d.drawingData;
    drawingGroup = dg;
    drawingNumber = d.drawingNumber;
    drawingGroup.addDrawing(this);
  }

  /**
   * Constructor invoked when writing the images
   *
   * @param x the column
   * @param y the row
   * @param w the width in cells
   * @param h the height in cells
   * @param image the image file
   */
  public Drawing2(double x,
                  double y,
                  double w,
                  double h,
                  java.io.File image)
  {
    imageFile = image;
    initialized = true;
    origin = Origin.WRITE;
    this.x = x;
    this.y = y;
    this.width = w;
    this.height = h;
    referenceCount = 1;
    type = ShapeType.PICTURE_FRAME;
  }

  /**
   * Constructor invoked when writing the images
   *
   * @param x the column
   * @param y the row
   * @param w the width in cells
   * @param h the height in cells
   * @param image the image data
   */
  public Drawing2(double x,
                  double y,
                  double w,
                  double h,
                  byte[] image)
  {
    imageData = image;
    initialized = true;
    origin = Origin.WRITE;
    this.x = x;
    this.y = y;
    this.width = w;
    this.height = h;
    referenceCount = 1;
    type = ShapeType.PICTURE_FRAME;
  }

  /**
   * Initializes the member variables from the Escher stream data
   */
  private void initialize()
  {
    initialized = true;
  }

  /**
   * Sets the object id.  Invoked by the drawing group when the object is
   * added to id
   *
   * @param objid the object id
   * @param bip the blip id
   * @param sid the shape id
   */
  public final void setObjectId(int objid, int bip, int sid)
  {
    objectId = objid;
    blipId = bip;
    shapeId = sid;

    if (origin == Origin.READ)
    {
      origin = Origin.READ_WRITE;
    }
  }

  /**
   * Accessor for the object id
   *
   * @return the object id
   */
  public final int getObjectId()
  {
    if (!initialized)
    {
      initialize();
    }

    return objectId;
  }

  /**
   * Accessor for the shape id
   *
   * @return the shape id
   */
  public int getShapeId()
  {
    if (!initialized)
    {
      initialize();
    }

    return shapeId;
  }

  /**
   * Accessor for the blip id
   *
   * @return the blip id
   */
  public final int getBlipId()
  {
    if (!initialized)
    {
      initialize();
    }

    return blipId;
  }

  /**
   * Gets the drawing record which was read in
   *
   * @return the drawing record
   */
  public MsoDrawingRecord  getMsoDrawingRecord()
  {
    return msoDrawingRecord;
  }

  /**
   * Creates the main Sp container for the drawing
   *
   * @return the SP container
   */
  public EscherContainer getSpContainer()
  {
    if (!initialized)
    {
      initialize();
    }

    Assert.verify(origin == Origin.READ);

    return getReadSpContainer();
  }

  /**
   * Sets the drawing group for this drawing.  Called by the drawing group
   * when this drawing is added to it
   *
   * @param dg the drawing group
   */
  public void setDrawingGroup(DrawingGroup dg)
  {
    drawingGroup = dg;
  }

  /**
   * Accessor for the drawing group
   *
   * @return the drawing group
   */
  public DrawingGroup getDrawingGroup()
  {
    return drawingGroup;
  }

  /**
   * Gets the origin of this drawing
   *
   * @return where this drawing came from
   */
  public Origin getOrigin()
  {
    return origin;
  }

  /**
   * Accessor for the reference count on this drawing
   *
   * @return the reference count
   */
  public int getReferenceCount()
  {
    return referenceCount;
  }

  /**
   * Sets the new reference count on the drawing
   *
   * @param r the new reference count
   */
  public void setReferenceCount(int r)
  {
    referenceCount = r;
  }

  /**
   * Accessor for the column of this drawing
   *
   * @return the column
   */
  public double getX()
  {
    if (!initialized)
    {
      initialize();
    }
    return x;
  }

  /**
   * Sets the column position of this drawing
   *
   * @param x the column
   */
  public void setX(double x)
  {
    if (origin == Origin.READ)
    {
      if (!initialized)
      {
        initialize();
      }
      origin = Origin.READ_WRITE;
    }

    this.x = x;
  }

  /**
   * Accessor for the row of this drawing
   *
   * @return the row
   */
  public double getY()
  {
    if (!initialized)
    {
      initialize();
    }

    return y;
  }

  /**
   * Accessor for the row of the drawing
   *
   * @param y the row
   */
  public void setY(double y)
  {
    if (origin == Origin.READ)
    {
      if (!initialized)
      {
        initialize();
      }
      origin = Origin.READ_WRITE;
    }

    this.y = y;
  }


  /**
   * Accessor for the width of this drawing
   *
   * @return the number of columns spanned by this image
   */
  public double getWidth()
  {
    if (!initialized)
    {
      initialize();
    }

    return width;
  }

  /**
   * Accessor for the width
   *
   * @param w the number of columns to span
   */
  public void setWidth(double w)
  {
    if (origin == Origin.READ)
    {
      if (!initialized)
      {
        initialize();
      }
      origin = Origin.READ_WRITE;
    }

    width = w;
  }

  /**
   * Accessor for the height of this drawing
   *
   * @return the number of rows spanned by this image
   */
  public double getHeight()
  {
    if (!initialized)
    {
      initialize();
    }

    return height;
  }

  /**
   * Accessor for the height of this drawing
   *
   * @param h the number of rows spanned by this image
   */
  public void setHeight(double h)
  {
    if (origin == Origin.READ)
    {
      if (!initialized)
      {
        initialize();
      }
      origin = Origin.READ_WRITE;
    }

    height = h;
  }


  /**
   * Gets the SpContainer that was read in
   *
   * @return the read sp container
   */
  private EscherContainer getReadSpContainer()
  {
    if (!initialized)
    {
      initialize();
    }

    return readSpContainer;
  }

  /**
   * Accessor for the image data
   *
   * @return the image data
   */
  public byte[] getImageData()
  {
    Assert.verify(false);
    Assert.verify(origin == Origin.READ || origin == Origin.READ_WRITE);

    if (!initialized)
    {
      initialize();
    }

    return drawingGroup.getImageData(blipId);
  }

  /**
   * Accessor for the image data
   *
   * @return the image data
   */
  public byte[] getImageBytes() throws IOException
  {
    Assert.verify(false);
    if (origin == Origin.READ || origin == Origin.READ_WRITE)
    {
      return getImageData();
    }

    Assert.verify(origin == Origin.WRITE);

    if (imageFile == null)
    {
      Assert.verify(imageData != null);
      return imageData;
    }

    byte[] data = new byte[(int) imageFile.length()];
    FileInputStream fis = new FileInputStream(imageFile);
    fis.read(data, 0, data.length);
    fis.close();
    return data;
  }

  /**
   * Accessor for the type
   *
   * @return the type
   */
  public ShapeType getType()
  {
    return type;
  }

  /**
   * Writes any other records associated with this drawing group object
   *
   * @param outputFile the output file
   * @exception IOException
   */
  public void writeAdditionalRecords(File outputFile) throws IOException
  {
    // no records to write
  }

  /**
   * Writes any records that need to be written after all the drawing group
   * objects have been written
   * Does nothing here
   *
   * @param outputFile the output file
   * @exception IOException
   */
  public void writeTailRecords(File outputFile) throws IOException
  {
    // does nothing
  }

  /**
   * Interface method
   *
   * @return the column number at which the image is positioned
   */
  public double getColumn()
  {
    return getX();
  }

  /**
   * Interface method
   *
   * @return the row number at which the image is positions
   */
  public double getRow()
  {
    return getY();
  }

  /**
   * Accessor for the first drawing on the sheet.  This is used when
   * copying unmodified sheets to indicate that this drawing contains
   * the first time Escher gubbins
   *
   * @return TRUE if this MSORecord is the first drawing on the sheet
   */
  public boolean isFirst()
  {
    return msoDrawingRecord.isFirst();
  }

  /**
   * Queries whether this object is a form object.  Form objects have their
   * drawings records spread over TXO and CONTINUE records and
   * require special handling
   *
   * @return TRUE if this is a form object, FALSE otherwise
   */
  public boolean isFormObject()
  {
    return false;
  }

  /**
   * Removes a row
   *
   * @param r the row to be removed
   */
  public void removeRow(int r)
  {
    if (y > r)
    {
      setY(r);
    }
  }

  /**
   * Accessor for the image file path.  Normally this is the absolute path
   * of a file on the directory system, but if this drawing was constructed
   * using an byte[] then the blip id is returned
   *
   * @return the image file path, or the blip id
   */
  public String getImageFilePath()
  {
    Assert.verify(false);
    return null;
  }
}



