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

package jxl.biff.drawing;

import java.io.FileInputStream;
import java.io.IOException;

import jxl.common.Assert;
import jxl.common.Logger;
import jxl.common.LengthUnit;
import jxl.common.LengthConverter;

import jxl.Image;
import jxl.Sheet;
import jxl.CellView;
import jxl.write.biff.File;


/**
 * Contains the various biff records used to insert a drawing into a
 * worksheet
 */
public class Drawing implements DrawingGroupObject, Image
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
   * The ObjRecord associated with the drawing
   */
  private ObjRecord objRecord;

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
   * A reference to the sheet containing this drawing.  Used to calculate
   * the drawing dimensions in pixels
   */
  private Sheet sheet;

  /**
   * Reader for the raw image data
   */
  private PNGReader pngReader;

  /**
   * The client anchor properties
   */
  private ImageAnchorProperties imageAnchorProperties;

  // Enumeration type for the image anchor properties
  protected static class ImageAnchorProperties
  {
    private int value;
    private static ImageAnchorProperties[] o = new ImageAnchorProperties[0];

    ImageAnchorProperties(int v)
    {
      value = v;
      
      ImageAnchorProperties[] oldArray = o;
      o = new ImageAnchorProperties[oldArray.length + 1];
      System.arraycopy(oldArray, 0, o, 0, oldArray.length);
      o[oldArray.length] = this;
    }

    int getValue()
    {
      return value;
    }

    static ImageAnchorProperties getImageAnchorProperties(int val)
    {
      ImageAnchorProperties iap = MOVE_AND_SIZE_WITH_CELLS;
      int pos = 0;
      while (pos < o.length)
      {
        if (o[pos].getValue()== val)
        {
          iap = o[pos];
          break;
        }
        else
        {
          pos++;
        }
      }
      return iap;
    }
  }

  // The image anchor properties
  public static ImageAnchorProperties MOVE_AND_SIZE_WITH_CELLS = 
    new ImageAnchorProperties(1);
  public static ImageAnchorProperties MOVE_WITH_CELLS = 
    new ImageAnchorProperties(2);
  public static ImageAnchorProperties NO_MOVE_OR_SIZE_WITH_CELLS = 
    new ImageAnchorProperties(3);

  /**
   * The default font size for columns
   */
  private static final double DEFAULT_FONT_SIZE = 10;

  /**
   * Constructor used when reading images
   *
   * @param mso the drawing record
   * @param obj the object record
   * @param dd the drawing data for all drawings on this sheet
   * @param dg the drawing group
   */
  public Drawing(MsoDrawingRecord mso,
                 ObjRecord obj,
                 DrawingData dd,
                 DrawingGroup dg,
                 Sheet s)
  {
    drawingGroup = dg;
    msoDrawingRecord = mso;
    drawingData = dd;
    objRecord = obj;
    sheet = s;
    initialized = false;
    origin = Origin.READ;
    drawingData.addData(msoDrawingRecord.getData());
    drawingNumber = drawingData.getNumDrawings() - 1;
    drawingGroup.addDrawing(this);

    Assert.verify(mso != null && obj != null);

    initialize();
  }

  /**
   * Copy constructor used to copy drawings from read to write
   *
   * @param dgo the drawing group object
   * @param dg the drawing group
   */
  protected Drawing(DrawingGroupObject dgo, DrawingGroup dg)
  {
    Drawing d = (Drawing) dgo;
    Assert.verify(d.origin == Origin.READ);
    msoDrawingRecord = d.msoDrawingRecord;
    objRecord = d.objRecord;
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
  public Drawing(double x,
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
    imageAnchorProperties = MOVE_WITH_CELLS;
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
  public Drawing(double x,
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
    imageAnchorProperties = MOVE_WITH_CELLS;
    type = ShapeType.PICTURE_FRAME;
  }

  /**
   * Initializes the member variables from the Escher stream data
   */
  private void initialize()
  {
    readSpContainer = drawingData.getSpContainer(drawingNumber);
    Assert.verify(readSpContainer != null);

    EscherRecord[] children = readSpContainer.getChildren();

    Sp sp = (Sp) readSpContainer.getChildren()[0];
    shapeId = sp.getShapeId();
    objectId = objRecord.getObjectId();
    type = ShapeType.getType(sp.getShapeType());

    if (type == ShapeType.UNKNOWN)
    {
      logger.warn("Unknown shape type");
    }

    Opt opt = (Opt) readSpContainer.getChildren()[1];

    if (opt.getProperty(260) != null)
    {
      blipId = opt.getProperty(260).value;
    }

    if (opt.getProperty(261) != null)
    {
      imageFile = new java.io.File(opt.getProperty(261).stringValue);
    }
    else
    {
      if (type == ShapeType.PICTURE_FRAME)
      {
        logger.warn("no filename property for drawing");
        imageFile = new java.io.File(Integer.toString(blipId));
      }
    }

    ClientAnchor clientAnchor = null;
    for (int i = 0; i < children.length && clientAnchor == null; i++)
    {
      if (children[i].getType() == EscherRecordType.CLIENT_ANCHOR)
      {
        clientAnchor = (ClientAnchor) children[i];
      }
    }

    if (clientAnchor == null)
    {
      logger.warn("client anchor not found");
    }
    else
    {
      x = clientAnchor.getX1();
      y = clientAnchor.getY1();
      width = clientAnchor.getX2() - x;
      height = clientAnchor.getY2() - y;
      imageAnchorProperties = ImageAnchorProperties.getImageAnchorProperties
        (clientAnchor.getProperties());
    }

    if (blipId == 0)
    {
      logger.warn("linked drawings are not supported");
    }

    initialized = true;
  }

  /**
   * Accessor for the image file
   *
   * @return the image file
   */
  public java.io.File getImageFile()
  {
    return imageFile;
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
    if (imageFile == null)
    {
      // return the blip id, if it exists
      return blipId != 0 ? Integer.toString(blipId) : "__new__image__";
    }

    return imageFile.getPath();
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

    if (origin == Origin.READ)
    {
      return getReadSpContainer();
    }

    SpContainer spContainer = new SpContainer();
    Sp sp = new Sp(type, shapeId, 2560);
    spContainer.add(sp);
    Opt opt = new Opt();
    opt.addProperty(260, true, false, blipId);

    if (type == ShapeType.PICTURE_FRAME)
    {
      String filePath = imageFile != null ? imageFile.getPath() : "";
      opt.addProperty(261, true, true, filePath.length() * 2, filePath);
      opt.addProperty(447, false, false, 65536);
      opt.addProperty(959, false, false, 524288);
      spContainer.add(opt);
    }

    ClientAnchor clientAnchor = new ClientAnchor
      (x, y, x + width, y + height, 
       imageAnchorProperties.getValue());
    spContainer.add(clientAnchor);
    ClientData clientData = new ClientData();
    spContainer.add(clientData);

    return spContainer;
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
    if (origin == Origin.READ)
    {
      outputFile.write(objRecord);
      return;
    }

    // Create the obj record
    ObjRecord objrec = new ObjRecord(objectId,
                                     ObjRecord.PICTURE);
    outputFile.write(objrec);
  }

  /**
   * Writes any records that need to be written after all the drawing group
   * objects have been written
   * Does nothing here
   *
   * @param outputFile the output file
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
   * Accessor for the image dimensions.  See technotes for Bill's explanation
   * of the calculation logic
   *
   * @return  approximate drawing size in pixels
   */
  private double getWidthInPoints()
  {
    if (sheet == null)
    {
      logger.warn("calculating image width:  sheet is null");
      return 0;
    }

    // The start and end row numbers
    int firstCol = (int) x;
    int lastCol = (int) Math.ceil(x + width) - 1;

    // **** MAGIC NUMBER ALERT ***
    // multiply the point size of the font by 0.59 to give the point size
    // I know of no explanation for this yet, other than that it seems to
    // give the right answer

    // Get the width of the image within the first col, allowing for 
    // fractional offsets
    CellView cellView = sheet.getColumnView(firstCol);
    int firstColWidth = cellView.getSize();
    double firstColImageWidth =  (1 - (x - firstCol)) * firstColWidth;
    double pointSize = (cellView.getFormat() != null) ? 
      cellView.getFormat().getFont().getPointSize() : DEFAULT_FONT_SIZE;
    double firstColWidthInPoints = firstColImageWidth * 0.59 * pointSize / 256;

    // Get the height of the image within the last row, allowing for
    // fractional offsets
    int lastColWidth = 0;
    double lastColImageWidth = 0;
    double lastColWidthInPoints = 0;
    if (lastCol != firstCol)
    {
      cellView = sheet.getColumnView(lastCol);
      lastColWidth = cellView.getSize();
      lastColImageWidth = (x + width - lastCol) * lastColWidth;
      pointSize = (cellView.getFormat() != null) ? 
        cellView.getFormat().getFont().getPointSize() : DEFAULT_FONT_SIZE;
      lastColWidthInPoints = lastColImageWidth * 0.59 * pointSize / 256;
    }
    
    // Now get all the columns in between
    double width = 0;
    for (int i = 0 ; i < lastCol - firstCol - 1 ; i++)
    {
      cellView = sheet.getColumnView(firstCol + 1 +i);
      pointSize = (cellView.getFormat() != null) ? 
        cellView.getFormat().getFont().getPointSize() : DEFAULT_FONT_SIZE;
      width += cellView.getSize() * 0.59 * pointSize / 256;
    }

    // Add on the first and last row contributions to get the height in twips
    double widthInPoints = width + 
      firstColWidthInPoints + lastColWidthInPoints;
    
    return widthInPoints;
  }

  /**
   * Accessor for the image dimensions.  See technotes for Bill's explanation
   * of the calculation logic
   *
   * @return approximate drawing size in pixels
   */
  private double getHeightInPoints()
  {
    if (sheet == null)
    {
      logger.warn("calculating image height:  sheet is null");
      return 0;
    }

    // The start and end row numbers
    int firstRow = (int) y;
    int lastRow = (int) Math.ceil(y + height) - 1;

    // Get the height of the image within the first row, allowing for 
    // fractional offsets
    int firstRowHeight = sheet.getRowView(firstRow).getSize();
    double firstRowImageHeight =  (1 - (y - firstRow)) * firstRowHeight;

    // Get the height of the image within the last row, allowing for
    // fractional offsets
    int lastRowHeight = 0;
    double lastRowImageHeight = 0;
    if (lastRow != firstRow)
    {
      lastRowHeight = sheet.getRowView(lastRow).getSize();
      lastRowImageHeight = (y + height - lastRow) * lastRowHeight;
    }
    
    // Now get all the rows in between
    double height = 0;
    for (int i = 0 ; i < lastRow - firstRow - 1 ; i++)
    {
      height += sheet.getRowView(firstRow + 1 + i).getSize();
    }

    // Add on the first and last row contributions to get the height in twips
    double heightInTwips = height + firstRowHeight + lastRowHeight;
    
    // Now divide by the magic number to converts twips into pixels and 
    // return the value
    double heightInPoints = heightInTwips / 20.0;

    return heightInPoints;
  }

  /**
   * Get the width of this image as rendered within Excel
   *
   * @param unit the unit of measurement
   * @return the width of the image within Excel
   */
  public double getWidth(LengthUnit unit)
  {
    double widthInPoints = getWidthInPoints();
    return widthInPoints * LengthConverter.getConversionFactor
      (LengthUnit.POINTS, unit);
  }

  /**
   * Get the height of this image as rendered within Excel
   *
   * @param unit the unit of measurement
   * @return the height of the image within Excel
   */
  public double getHeight(LengthUnit unit)
  {
    double heightInPoints = getHeightInPoints();
    return heightInPoints * LengthConverter.getConversionFactor
      (LengthUnit.POINTS, unit);
  }

  /**
   * Gets the width of the image.  Note that this is the width of the 
   * underlying image, and does not take into account any size manipulations
   * that may have occurred when the image was added into Excel
   *
   * @return the image width in pixels
   */
  public int getImageWidth()
  {
    return getPngReader().getWidth();
  }

  /**
   * Gets the height of the image.  Note that this is the height of the 
   * underlying image, and does not take into account any size manipulations
   * that may have occurred when the image was added into Excel
   *
   * @return the image width in pixels
   */
  public int getImageHeight()
  {
    return getPngReader().getHeight();
  }


  /**
   * Gets the horizontal resolution of the image, if that information
   * is available.
   *
   * @return the number of dots per unit specified, if available, 0 otherwise
   */
  public double getHorizontalResolution(LengthUnit unit)
  {
    int res = getPngReader().getHorizontalResolution();
    return res / LengthConverter.getConversionFactor(LengthUnit.METRES, unit);
  }

  /**
   * Gets the vertical resolution of the image, if that information
   * is available.
   *
   * @return the number of dots per unit specified, if available, 0 otherwise
   */
  public double getVerticalResolution(LengthUnit unit)
  {
    int res = getPngReader().getVerticalResolution();
    return res / LengthConverter.getConversionFactor(LengthUnit.METRES, unit);
  }

  private PNGReader getPngReader()
  {
    if (pngReader != null)
    {
      return pngReader;
    }

    byte[] imdata = null;
    if (origin == Origin.READ || origin == Origin.READ_WRITE)
    {
      imdata = getImageData();
    }
    else
    {
      try
      {
        imdata = getImageBytes();
      }
      catch (IOException e)
      {
        logger.warn("Could not read image file");
        imdata = new byte[0];
      }
    }

    pngReader = new PNGReader(imdata);
    pngReader.read();
    return pngReader;
  }

  /**
   * Accessor for the anchor properties
   */
  protected void setImageAnchor(ImageAnchorProperties iap)
  {
    imageAnchorProperties = iap;

    if (origin == Origin.READ)
    {
      origin = Origin.READ_WRITE;
    }
  }

  /**
   * Accessor for the anchor properties
   */
  protected ImageAnchorProperties getImageAnchor()
  {
    if (!initialized)
    {
      initialize();
    }

    return imageAnchorProperties;
  }
}



