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

package jxl.write.biff;

import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;

import jxl.common.Assert;
import jxl.common.Logger;

import jxl.biff.BaseCompoundFile;
import jxl.biff.IntegerHelper;
import jxl.read.biff.BiffException;

/**
 * Writes out a compound file
 * 
 * Header block is -1
 * Excel data is e..n (where e is the head extension blocks, normally 0 and
 * n is at least 8)
 * Summary information (8 blocks)
 * Document summary (8 blocks)
 * BBD is block p..q (where p=e+n+16 and q-p+1 is the number of BBD blocks)
 * Property storage block is q+b...r (normally 1 block) (where b is the number
 * of BBD blocks)
 */
final class CompoundFile extends BaseCompoundFile
{
  /**
   * The logger
   */
  private static Logger logger = Logger.getLogger(CompoundFile.class);

  /**
   * The stream to which the jumbled up data is written to
   */
  private OutputStream out;
  /**
   * The organized biff records which form the actual excel data
   */
  private ExcelDataOutput excelData;

  /**
   * The size of the array
   */
  private int   size;

  /**
   * The size the excel data should be in order to comply with the
   * general compound file format
   */
  private int    requiredSize;

  /**
   * The number of blocks it takes to store the big block depot
   */
  private int numBigBlockDepotBlocks;

  /**
   * The number of blocks it takes to store the small block depot chain
   */
  private int numSmallBlockDepotChainBlocks;

  /**
   * The number of blocks it takes to store the small block depot
   */
  private int numSmallBlockDepotBlocks;

  /**
   * The number of extension blocks required for the header to describe
   * the BBD
   */
  private int numExtensionBlocks;

  /**
   * The extension block for the header
   */
  private int extensionBlock;

  /**
   * The number of blocks it takes to store the excel data
   */
  private int excelDataBlocks;

  /**
   * The start block of the root entry
   */
  private int rootStartBlock;

  /**
   * The start block of the excel data
   */
  private int excelDataStartBlock;

  /**
   * The start block of the big block depot
   */
  private int bbdStartBlock;

  /**
   * The start block of the small block depot
   */
  private int sbdStartBlockChain;

  /**
   * The start block of the small block depot
   */
  private int sbdStartBlock;

  /**
   * The number of big blocks required for additional property sets
   */
  private int additionalPropertyBlocks;

  /**
   * The number of small blocks
   */
  private int numSmallBlocks;

  /**
   * The total number of property sets in this compound file
   */
  private int numPropertySets;

  /**
   * The number of blocks required to store the root entry property sets
   * and small block depot
   */
  private int numRootEntryBlocks;

  /**
   * The list of additional, non standard property sets names
   */
  private ArrayList additionalPropertySets;

  /**
   * The map of standard property sets, keyed on name
   */
  private HashMap standardPropertySets;

  /**
   * Structure used to store the property set and the data
   */
  private static final class ReadPropertyStorage
  {
    PropertyStorage propertyStorage;
    byte[] data;
    int number;

    ReadPropertyStorage(PropertyStorage ps, byte[] d, int n)
    {
      propertyStorage = ps;
      data = d;
      number = n;
    }
  }


  // The following member variables are used across methods when
  // writing out the big block depot
  /**
   * The current position within the bbd.  Used when writing out the
   * BBD
   */
  private int bbdPos;

  /**
   * The current bbd block
   */
  private byte[] bigBlockDepot;


  /**
   * Constructor
   * 
   * @param l the length of the data
   * @param os the output stream to write to
   * @param data the excel data
   * @param rcf the read compound
   */
  public CompoundFile(ExcelDataOutput data, int l, OutputStream os, 
                      jxl.read.biff.CompoundFile rcf) 
    throws CopyAdditionalPropertySetsException, IOException
  {
    super();
    size = l;
    excelData = data;

    readAdditionalPropertySets(rcf);

    numRootEntryBlocks = 1;
    numPropertySets = 4 + 
      (additionalPropertySets != null ? additionalPropertySets.size() : 0);


    if (additionalPropertySets != null)
    {
      numSmallBlockDepotChainBlocks = getBigBlocksRequired(numSmallBlocks * 4);
      numSmallBlockDepotBlocks = getBigBlocksRequired
        (numSmallBlocks * SMALL_BLOCK_SIZE);

      numRootEntryBlocks += getBigBlocksRequired
        (additionalPropertySets.size() * PROPERTY_STORAGE_BLOCK_SIZE);
    }


    int blocks  = getBigBlocksRequired(l);

    // First pad the data out so that it fits nicely into a whole number
    // of blocks
    if (l < SMALL_BLOCK_THRESHOLD)
    {
      requiredSize = SMALL_BLOCK_THRESHOLD;
    }
    else
    {
      requiredSize = blocks * BIG_BLOCK_SIZE;
    }
    
    out = os;


    // Do the calculations
    excelDataBlocks = requiredSize/BIG_BLOCK_SIZE;
    numBigBlockDepotBlocks = 1;

    int blockChainLength = (BIG_BLOCK_SIZE - BIG_BLOCK_DEPOT_BLOCKS_POS)/4;

    int startTotalBlocks = excelDataBlocks + 
      8 + // summary block
      8 + // document information
      additionalPropertyBlocks +
      numSmallBlockDepotBlocks +
      numSmallBlockDepotChainBlocks +
      numRootEntryBlocks;

    int totalBlocks = startTotalBlocks + numBigBlockDepotBlocks;

    // Calculate the number of BBD blocks needed to hold this info
    numBigBlockDepotBlocks = (int) Math.ceil( (double) totalBlocks / 
                                              (double) (BIG_BLOCK_SIZE/4));

    // Does this affect the total?
    totalBlocks = startTotalBlocks + numBigBlockDepotBlocks;

    // And recalculate
    numBigBlockDepotBlocks = (int) Math.ceil( (double) totalBlocks / 
                                              (double) (BIG_BLOCK_SIZE/4));

    // Does this affect the total?
    totalBlocks = startTotalBlocks + numBigBlockDepotBlocks;

    // See if the excel bbd chain can fit into the header block.
    // Remember to allow for the  end of chain indicator
    if (numBigBlockDepotBlocks > blockChainLength - 1 )
    {
      // Sod it - we need an extension block.  We have to go through
      // the whole tiresome calculation again
      extensionBlock = 0;

      // Compute the number of extension blocks
      int bbdBlocksLeft = numBigBlockDepotBlocks - blockChainLength + 1;

      numExtensionBlocks = (int) Math.ceil((double) bbdBlocksLeft /
                                           (double) (BIG_BLOCK_SIZE/4 - 1));

      // Modify the total number of blocks required and recalculate the
      // the number of bbd blocks
      totalBlocks = startTotalBlocks + 
                    numExtensionBlocks + 
                    numBigBlockDepotBlocks;
      numBigBlockDepotBlocks = (int) Math.ceil( (double) totalBlocks / 
                                                (double) (BIG_BLOCK_SIZE/4));

      // The final total
      totalBlocks = startTotalBlocks + 
                    numExtensionBlocks + 
                    numBigBlockDepotBlocks;
    }
    else
    {
      extensionBlock = -2;
      numExtensionBlocks = 0;
    }

    // Set the excel data start block to be after the header (and
    // its extensions)
    excelDataStartBlock = numExtensionBlocks;

    // Set the start block of the small block depot
    sbdStartBlock = -2;
    if (additionalPropertySets != null && numSmallBlockDepotBlocks != 0)
    {
      sbdStartBlock = excelDataStartBlock + 
                      excelDataBlocks + 
                      additionalPropertyBlocks +
                      16;
    }

    // Set the sbd chain start block to be after the excel data and the
    // small block depot
    sbdStartBlockChain = -2;
    
    if (sbdStartBlock != -2)
    {
      sbdStartBlockChain = sbdStartBlock + numSmallBlockDepotBlocks;
    }
    
    // Set the bbd start block to be after all the excel data
    if (sbdStartBlockChain != -2)
    {
      bbdStartBlock = sbdStartBlockChain +
                      numSmallBlockDepotChainBlocks;
    }
    else
    {
      bbdStartBlock = excelDataStartBlock +
                      excelDataBlocks + 
                      additionalPropertyBlocks +
                      16;
    }

    // Set the root start block to be after all the big block depot blocks
    rootStartBlock = bbdStartBlock +
                     numBigBlockDepotBlocks;


    if (totalBlocks != rootStartBlock + numRootEntryBlocks)
    {
      logger.warn("Root start block and total blocks are inconsistent " + 
                  " generated file may be corrupt");
      logger.warn("RootStartBlock " + rootStartBlock + " totalBlocks " + totalBlocks);
    }
  }

  /**
   * Reads the additional property sets from the read in compound file
   *
   * @param readCompoundFile the file read in
   * @exception CopyAdditionalPropertySetsException
   * @exception IOException
   */
  private void readAdditionalPropertySets
    (jxl.read.biff.CompoundFile readCompoundFile) 
    throws CopyAdditionalPropertySetsException, IOException
  {
    if (readCompoundFile == null)
    {
      return;
    }

    additionalPropertySets = new ArrayList();
    standardPropertySets = new HashMap();
    int blocksRequired = 0;

    int numPropertySets = readCompoundFile.getNumberOfPropertySets();
    
    for (int i = 0 ; i < numPropertySets ; i++)
    {
      PropertyStorage ps = readCompoundFile.getPropertySet(i);

      boolean standard = false;

      if (ps.name.equalsIgnoreCase(ROOT_ENTRY_NAME))
      {
        standard = true;
        ReadPropertyStorage rps = new ReadPropertyStorage(ps, null, i);
        standardPropertySets.put(ROOT_ENTRY_NAME, rps);
      }

      // See if it is a standard property set
      for (int j = 0 ; j < STANDARD_PROPERTY_SETS.length && !standard ; j++)
      {
        if (ps.name.equalsIgnoreCase(STANDARD_PROPERTY_SETS[j]))
        {
          // See if it comes directly off the root entry
          PropertyStorage ps2 = readCompoundFile.findPropertyStorage(ps.name);
          Assert.verify(ps2 != null);

          if (ps2 == ps)
          {
            standard = true;
            ReadPropertyStorage rps = new ReadPropertyStorage(ps, null, i);
            standardPropertySets.put(STANDARD_PROPERTY_SETS[j], rps);
          }
        }
      }

      if (!standard)
      {
        try
        {
          byte[] data = null;
          if (ps.size > 0 )
          {
            data = readCompoundFile.getStream(i);
          }
          else
          {
            data = new byte[0];
          }
          ReadPropertyStorage rps = new ReadPropertyStorage(ps, data, i);
          additionalPropertySets.add(rps);

          if (data.length > SMALL_BLOCK_THRESHOLD)
          {
            int blocks = getBigBlocksRequired(data.length);
            blocksRequired += blocks;
          }
          else
          {
            int blocks = getSmallBlocksRequired(data.length);
            numSmallBlocks += blocks;
          }
        }
        catch (BiffException e)
        {
          logger.error(e);
          throw new CopyAdditionalPropertySetsException();
        }
      }
    }

    additionalPropertyBlocks = blocksRequired;
  }

  /**
   * Writes out the excel file in OLE compound file format
   * 
   * @exception IOException 
   */
  public void write() throws IOException
  {
    writeHeader();
    writeExcelData();
    writeDocumentSummaryData();
    writeSummaryData();
    writeAdditionalPropertySets();
    writeSmallBlockDepot();
    writeSmallBlockDepotChain();
    writeBigBlockDepot();
    writePropertySets();
    
    // Don't flush or close the stream - this is handled by the enclosing File
    // object
  }

  /**
   * Writes out any additional property sets
   */
  private void writeAdditionalPropertySets() throws IOException
  {
    if (additionalPropertySets == null)
    {
      return;
    }

    for (Iterator i = additionalPropertySets.iterator(); i.hasNext() ;)
    {
      ReadPropertyStorage rps = (ReadPropertyStorage) i.next();
      byte[] data = rps.data;

      if (data.length > SMALL_BLOCK_THRESHOLD)
      {
        int numBlocks = getBigBlocksRequired(data.length);
        int requiredSize = numBlocks * BIG_BLOCK_SIZE;

        out.write(data, 0, data.length);
      
        byte[] padding = new byte[requiredSize - data.length];
        out.write(padding, 0, padding.length);
      }
    }
  }

  /**
   * Writes out the excel data, padding it out with empty bytes as
   * necessary
   * Also write out empty 
   * 
   * @exception IOException 
   */
  private void writeExcelData() throws IOException
  {
    excelData.writeData(out);

    byte[] padding = new byte[requiredSize - size];
    out.write(padding);
  }

  /**
   * Write out the document summary data.  This is just blank
   * 
   * @exception IOException 
   */
  private void writeDocumentSummaryData() throws IOException
  {
    byte[] padding = new byte[SMALL_BLOCK_THRESHOLD];

    // Write out the summary information
    out.write(padding);
  }

  /**
   * Write out the  summary data.  This is just blank
   * 
   * @exception IOException 
   */
  private void writeSummaryData() throws IOException
  {
    byte[] padding = new byte[SMALL_BLOCK_THRESHOLD];

    // Write out the summary information
    out.write(padding);
  }

  /**
   * Writes the compound file header
   * 
   * @exception IOException 
   */
  private void writeHeader() throws IOException
  {
    // Build up the header array
    byte[] headerBlock = new byte[BIG_BLOCK_SIZE];
    byte[] extensionBlockData = new byte[BIG_BLOCK_SIZE * numExtensionBlocks];

    // Copy in the identifier
    System.arraycopy(IDENTIFIER, 0, headerBlock, 0, IDENTIFIER.length);

    // Copy in some magic values - no idea what they mean
    headerBlock[0x18] = 0x3e;
    headerBlock[0x1a] = 0x3;
    headerBlock[0x1c] = (byte) 0xfe;
    headerBlock[0x1d] = (byte) 0xff;
    headerBlock[0x1e] = 0x9;
    headerBlock[0x20] = 0x6;
    headerBlock[0x39] = 0x10;

    // Set the number of BBD blocks
    IntegerHelper.getFourBytes(numBigBlockDepotBlocks, 
                               headerBlock, 
                               NUM_BIG_BLOCK_DEPOT_BLOCKS_POS);  

    // Set the small block depot chain 
    IntegerHelper.getFourBytes(sbdStartBlockChain,
                               headerBlock,
                               SMALL_BLOCK_DEPOT_BLOCK_POS);

    // Set the number of blocks in the small block depot chain 
    IntegerHelper.getFourBytes(numSmallBlockDepotChainBlocks,
                               headerBlock,
                               NUM_SMALL_BLOCK_DEPOT_BLOCKS_POS);

    // Set the extension block 
    IntegerHelper.getFourBytes(extensionBlock,
                               headerBlock,
                               EXTENSION_BLOCK_POS);

    // Set the number of extension blocks to be the number of BBD blocks - 1
    IntegerHelper.getFourBytes(numExtensionBlocks,
                               headerBlock, 
                               NUM_EXTENSION_BLOCK_POS);
    
    // Set the root start block
    IntegerHelper.getFourBytes(rootStartBlock,
                               headerBlock,
                               ROOT_START_BLOCK_POS);

    // Set the block numbers for the BBD.  Set the BBD running 
    // after the excel data and summary information
    int pos = BIG_BLOCK_DEPOT_BLOCKS_POS;

    // See how many blocks fit into the header
    int blocksToWrite = Math.min(numBigBlockDepotBlocks, 
                                 (BIG_BLOCK_SIZE - 
                                  BIG_BLOCK_DEPOT_BLOCKS_POS)/4);
    int blocksWritten = 0;

    for (int i = 0 ; i < blocksToWrite; i++)
    {
      IntegerHelper.getFourBytes(bbdStartBlock + i, 
                                 headerBlock,
                                 pos);
      pos += 4;
      blocksWritten++;
    }
    
    // Pad out the rest of the header with blanks
    for (int i = pos; i < BIG_BLOCK_SIZE; i++)
    {
      headerBlock[i] = (byte) 0xff;
    }

    out.write(headerBlock);

    // Write out the extension blocks
    pos = 0;

    for (int extBlock = 0; extBlock < numExtensionBlocks; extBlock++)
    {
      blocksToWrite = Math.min(numBigBlockDepotBlocks - blocksWritten, 
                               BIG_BLOCK_SIZE/4 -1);

      for(int j = 0 ; j < blocksToWrite; j++)
      {
        IntegerHelper.getFourBytes(bbdStartBlock + blocksWritten + j, 
                                   extensionBlockData,
                                   pos);
        pos += 4;
      }

      blocksWritten += blocksToWrite;

      // Indicate the next block, or the termination of the chain
      int nextBlock = (blocksWritten == numBigBlockDepotBlocks) ? 
                              -2 : extBlock+1 ;
      IntegerHelper.getFourBytes(nextBlock, extensionBlockData, pos);
      pos +=4;
    }

    if (numExtensionBlocks > 0)
    {
      // Pad out the rest of the extension block with blanks
      for (int i = pos; i < extensionBlockData.length; i++)
      {
        extensionBlockData[i] = (byte) 0xff;
      }

      out.write(extensionBlockData);
    }
  }

  /**
   * Checks that the data can fit into the current BBD block.  If not,
   * then it moves on to the next block
   * 
   * @exception IOException 
   */
  private void checkBbdPos() throws IOException
  {
    if (bbdPos >= BIG_BLOCK_SIZE)
    {
      // Write out the extension block.  This will simply be the next block
      out.write(bigBlockDepot);
      
      // Create a new block
      bigBlockDepot = new byte[BIG_BLOCK_SIZE];
      bbdPos = 0;
    }
  }

  /**
   * Writes out the big block chain
   *
   * @param startBlock the starting block of the big block chain
   * @param numBlocks the number of blocks in the chain
   * @exception IOException
   */
  private void writeBlockChain(int startBlock, int numBlocks) 
    throws IOException
  {
    int blocksToWrite = numBlocks - 1;
    int blockNumber   = startBlock + 1;
    
    while (blocksToWrite > 0)
    {
      int bbdBlocks = Math.min(blocksToWrite, (BIG_BLOCK_SIZE - bbdPos)/4);

      for (int i = 0 ; i < bbdBlocks; i++)
      {
        IntegerHelper.getFourBytes(blockNumber, bigBlockDepot, bbdPos);
        bbdPos +=4 ;
        blockNumber++;
      }
      
      blocksToWrite -= bbdBlocks;
      checkBbdPos();
    }

    // Write the end of the block chain 
    IntegerHelper.getFourBytes(-2, bigBlockDepot, bbdPos);
    bbdPos += 4;
    checkBbdPos();
  }

  /**
   * Writes the block chains for the additional property sets
   *
   * @exception IOException
   */
  private void writeAdditionalPropertySetBlockChains() throws IOException
  {
    if (additionalPropertySets == null)
    {
      return;
    }

    int blockNumber = excelDataStartBlock + excelDataBlocks + 16;
    for (Iterator i = additionalPropertySets.iterator(); i.hasNext() ; )
    {
      ReadPropertyStorage rps = (ReadPropertyStorage) i.next();
      if (rps.data.length > SMALL_BLOCK_THRESHOLD)
      {
        int numBlocks = getBigBlocksRequired(rps.data.length);

        writeBlockChain(blockNumber, numBlocks);
        blockNumber += numBlocks;
      }
    }
  }
  
  /**
   * Writes out the chains for the small block depot
   */
  private void writeSmallBlockDepotChain() throws IOException
  {
    if (sbdStartBlockChain == -2)
    {
      return;
    }

    byte[] smallBlockDepotChain = 
      new byte[numSmallBlockDepotChainBlocks * BIG_BLOCK_SIZE];

    int pos = 0;
    int sbdBlockNumber = 1;

    for (Iterator i = additionalPropertySets.iterator(); i.hasNext() ; )
    {
      ReadPropertyStorage rps = (ReadPropertyStorage) i.next();

      if (rps.data.length <= SMALL_BLOCK_THRESHOLD &&
          rps.data.length != 0)
      {
        int numSmallBlocks = getSmallBlocksRequired(rps.data.length);
        for (int j = 0 ; j < numSmallBlocks - 1 ; j++)
        {
          IntegerHelper.getFourBytes(sbdBlockNumber, 
                                     smallBlockDepotChain, 
                                     pos);
          pos += 4;
          sbdBlockNumber++;
        }

        // Write out the end of chain
        IntegerHelper.getFourBytes(-2, smallBlockDepotChain, pos);
        pos += 4;
        sbdBlockNumber++;
      }
    }

    out.write(smallBlockDepotChain);
  }

  /**
   * Writes out all the data in the small block depot
   *
   * @exception
   */
  private void writeSmallBlockDepot() throws IOException
  {
    if (additionalPropertySets == null)
    {
      return;
    }

    byte[] smallBlockDepot = 
      new byte[numSmallBlockDepotBlocks * BIG_BLOCK_SIZE];

    int pos = 0;

    for (Iterator i = additionalPropertySets.iterator() ; i.hasNext() ; )
    {
      ReadPropertyStorage rps = (ReadPropertyStorage) i.next();

      if (rps.data.length <= SMALL_BLOCK_THRESHOLD)
      {
        int smallBlocks = getSmallBlocksRequired(rps.data.length);
        int length = smallBlocks * SMALL_BLOCK_SIZE;
        System.arraycopy(rps.data, 0, smallBlockDepot, pos, rps.data.length);
        pos += length;
      }
    }

    out.write(smallBlockDepot);
  }

  /**
   * Writes out the Big Block Depot
   * 
   * @exception IOException 
   */
  private void writeBigBlockDepot() throws IOException
  {
    // This is after the excel data, the summary information, the
    // big block property sets and the small block depot
    bigBlockDepot = new byte[BIG_BLOCK_SIZE];
    bbdPos = 0;

    // Write out the extension blocks, indicating them as special blocks
    for (int i = 0 ; i < numExtensionBlocks; i++)
    {
      IntegerHelper.getFourBytes(-3, bigBlockDepot, bbdPos);
      bbdPos += 4;
      checkBbdPos();
    }

    writeBlockChain(excelDataStartBlock, excelDataBlocks);
    
    // The excel data has been written.  Now write out the rest of it

    // Write the block chain for the summary information
    int summaryInfoBlock = excelDataStartBlock + 
      excelDataBlocks + 
      additionalPropertyBlocks;

    for (int i = summaryInfoBlock; i < summaryInfoBlock + 7; i++)
    {
      IntegerHelper.getFourBytes(i + 1, bigBlockDepot, bbdPos);
      bbdPos +=4 ;
      checkBbdPos();
    } 

    // Write the end of the block chain for the summary info block
    IntegerHelper.getFourBytes(-2, bigBlockDepot, bbdPos);
    bbdPos += 4;
    checkBbdPos();

    // Write the block chain for the document summary information
    for (int i = summaryInfoBlock + 8; i < summaryInfoBlock + 15; i++)
    {
      IntegerHelper.getFourBytes(i + 1, bigBlockDepot, bbdPos);
      bbdPos +=4 ;
      checkBbdPos();
    } 

    // Write the end of the block chain for the document summary
    IntegerHelper.getFourBytes(-2, bigBlockDepot, bbdPos);
    bbdPos += 4;
    checkBbdPos();

    // Write out the block chain for the copied property sets, if present
    writeAdditionalPropertySetBlockChains();

    if (sbdStartBlock != -2)
    {
      // Write out the block chain for the small block depot
      writeBlockChain(sbdStartBlock, numSmallBlockDepotBlocks);

      // Write out the block chain for the small block depot chain
      writeBlockChain(sbdStartBlockChain, numSmallBlockDepotChainBlocks);
    }

    // The Big Block Depot immediately follows.  Denote these as a special 
    // block
    for (int i = 0; i < numBigBlockDepotBlocks; i++)
    {
      IntegerHelper.getFourBytes(-3, bigBlockDepot, bbdPos);
      bbdPos += 4;
      checkBbdPos();
    }

    // Write the root entry
    writeBlockChain(rootStartBlock, numRootEntryBlocks);

    // Pad out the remainder of the block
    if (bbdPos != 0)
    {
      for (int i = bbdPos; i < BIG_BLOCK_SIZE; i++)
      {
        bigBlockDepot[i] = (byte) 0xff;
      }
      out.write(bigBlockDepot);
    }
  }

  /**
   * Calculates the number of big blocks required to store data of the 
   * specified length
   *
   * @param length the length of the data
   * @return the number of big blocks required to store the data
   */
  private int getBigBlocksRequired(int length)
  {
    int blocks = length / BIG_BLOCK_SIZE;
    
    return (length % BIG_BLOCK_SIZE > 0 )? blocks + 1 : blocks;
  }

  /**
   * Calculates the number of small blocks required to store data of the 
   * specified length
   *
   * @param length the length of the data
   * @return the number of small blocks required to store the data
   */
  private int getSmallBlocksRequired(int length)
  {
    int blocks = length / SMALL_BLOCK_SIZE;
    
    return (length % SMALL_BLOCK_SIZE > 0 )? blocks + 1 : blocks;
  }

  /**
   * Writes out the property sets
   * 
   * @exception IOException 
   */
  private void writePropertySets() throws IOException
  {
    byte[] propertySetStorage = new byte[BIG_BLOCK_SIZE * numRootEntryBlocks];

    int pos = 0;
    int[] mappings = null;

    // Build up the mappings array
    if (additionalPropertySets != null)
    {
      mappings = new int[numPropertySets];
      
      // Map the standard ones to the first four
      for (int i = 0 ; i < STANDARD_PROPERTY_SETS.length ; i++)
      {
        ReadPropertyStorage rps = (ReadPropertyStorage) 
          standardPropertySets.get(STANDARD_PROPERTY_SETS[i]);

        if (rps != null)
        {
          mappings[rps.number] = i;
        }
        else
        {
          logger.warn("Standard property set " + STANDARD_PROPERTY_SETS[i] + 
                      " not present in source file");
        }
      }

      // Now go through the original ones
      int newMapping = STANDARD_PROPERTY_SETS.length;
      for (Iterator i = additionalPropertySets.iterator(); i.hasNext(); )
      {
        ReadPropertyStorage rps = (ReadPropertyStorage) i.next();
        mappings[rps.number] = newMapping;
        newMapping++;
      }
    }

    int child = 0;
    int previous = 0;
    int next = 0;

    // Compute the size of the root property set
    int size = 0;
    
    if (additionalPropertySets != null)
    {
      // Workbook
      size += getBigBlocksRequired(requiredSize) * BIG_BLOCK_SIZE;

      // The two information blocks
      size += getBigBlocksRequired(SMALL_BLOCK_THRESHOLD) * BIG_BLOCK_SIZE;
      size += getBigBlocksRequired(SMALL_BLOCK_THRESHOLD) * BIG_BLOCK_SIZE;

      // Additional property sets
      for (Iterator i = additionalPropertySets.iterator(); i.hasNext(); )
      {
        ReadPropertyStorage rps = (ReadPropertyStorage) i.next();
        if (rps.propertyStorage.type != 1)
        {
          if (rps.propertyStorage.size >= SMALL_BLOCK_THRESHOLD)
          {
            size += getBigBlocksRequired(rps.propertyStorage.size) * 
              BIG_BLOCK_SIZE;
          }
          else
          {
            size += getSmallBlocksRequired(rps.propertyStorage.size) * 
              SMALL_BLOCK_SIZE;
          }
        }
      }
    }

    // Set the root entry property set
    PropertyStorage ps = new PropertyStorage(ROOT_ENTRY_NAME);
    ps.setType(5);
    ps.setStartBlock(sbdStartBlock);
    ps.setSize(size);
    ps.setPrevious(-1);
    ps.setNext(-1);
    ps.setColour(0);

    child = 1;
    if (additionalPropertySets != null)
    {
      ReadPropertyStorage rps = (ReadPropertyStorage) 
                            standardPropertySets.get(ROOT_ENTRY_NAME);
      child = mappings[rps.propertyStorage.child];
    }
    ps.setChild(child);

    System.arraycopy(ps.data, 0, 
                     propertySetStorage, pos, 
                     PROPERTY_STORAGE_BLOCK_SIZE);
    pos += PROPERTY_STORAGE_BLOCK_SIZE;


    // Set the workbook property set
    ps = new PropertyStorage(WORKBOOK_NAME);
    ps.setType(2);
    ps.setStartBlock(excelDataStartBlock);
      // start the excel data after immediately after this block
    ps.setSize(requiredSize);
      // always use a big block stream - none of that messing around
      // with small blocks
    
    previous = 3;
    next = -1;

    if (additionalPropertySets != null)
    {
      ReadPropertyStorage rps = (ReadPropertyStorage) 
        standardPropertySets.get(WORKBOOK_NAME);
      previous = rps.propertyStorage.previous != -1 ? 
        mappings[rps.propertyStorage.previous] : -1;
      next = rps.propertyStorage.next != -1 ? 
        mappings[rps.propertyStorage.next] : -1 ;
    }

    ps.setPrevious(previous);
    ps.setNext(next);
    ps.setChild(-1);

    System.arraycopy(ps.data, 0, 
                     propertySetStorage, pos, 
                     PROPERTY_STORAGE_BLOCK_SIZE);
    pos += PROPERTY_STORAGE_BLOCK_SIZE;

    // Set the summary information
    ps = new PropertyStorage(SUMMARY_INFORMATION_NAME);
    ps.setType(2);
    ps.setStartBlock(excelDataStartBlock + excelDataBlocks);
    ps.setSize(SMALL_BLOCK_THRESHOLD);

    previous = 1;
    next = 3;

    if (additionalPropertySets != null)
    {
      ReadPropertyStorage rps = (ReadPropertyStorage) 
                            standardPropertySets.get(SUMMARY_INFORMATION_NAME);

      if (rps != null)
      {
        previous = rps.propertyStorage.previous != - 1 ?
          mappings[rps.propertyStorage.previous] : -1 ;
        next = rps.propertyStorage.next != - 1 ?
          mappings[rps.propertyStorage.next] : -1 ;
      }
    }

    ps.setPrevious(previous);
    ps.setNext(next);
    ps.setChild(-1);

    System.arraycopy(ps.data, 0, 
                     propertySetStorage, pos, 
                     PROPERTY_STORAGE_BLOCK_SIZE);
    pos += PROPERTY_STORAGE_BLOCK_SIZE;

    // Set the document summary information
    ps = new PropertyStorage(DOCUMENT_SUMMARY_INFORMATION_NAME);
    ps.setType(2);
    ps.setStartBlock(excelDataStartBlock + excelDataBlocks + 8);
    ps.setSize(SMALL_BLOCK_THRESHOLD);
    ps.setPrevious(-1);
    ps.setNext(-1);
    ps.setChild(-1);

    System.arraycopy(ps.data, 0, 
                     propertySetStorage, pos, 
                     PROPERTY_STORAGE_BLOCK_SIZE);
    pos += PROPERTY_STORAGE_BLOCK_SIZE;



    // Write out the additional property sets
    if (additionalPropertySets == null)
    {
      out.write(propertySetStorage);
      return;
    }
    
    int bigBlock = excelDataStartBlock + excelDataBlocks + 16;
    int smallBlock = 0;

    for (Iterator i = additionalPropertySets.iterator() ; i.hasNext(); )
    {
      ReadPropertyStorage rps = (ReadPropertyStorage) i.next();
 
      int block = rps.data.length > SMALL_BLOCK_THRESHOLD ? 
        bigBlock : smallBlock;

      ps = new PropertyStorage(rps.propertyStorage.name);
      ps.setType(rps.propertyStorage.type);
      ps.setStartBlock(block);
      ps.setSize(rps.propertyStorage.size);
      //      ps.setColour(rps.propertyStorage.colour);

      previous = rps.propertyStorage.previous != -1 ? 
        mappings[rps.propertyStorage.previous] : -1;
      next = rps.propertyStorage.next != -1 ? 
        mappings[rps.propertyStorage.next] : -1;
      child = rps.propertyStorage.child != -1 ? 
        mappings[rps.propertyStorage.child] : -1;

      ps.setPrevious(previous);
      ps.setNext(next);
      ps.setChild(child);

      System.arraycopy(ps.data, 0, 
                       propertySetStorage, pos, 
                       PROPERTY_STORAGE_BLOCK_SIZE);
      pos += PROPERTY_STORAGE_BLOCK_SIZE;

      if (rps.data.length > SMALL_BLOCK_THRESHOLD)
      {
        bigBlock += getBigBlocksRequired(rps.data.length);
      }
      else
      {
        smallBlock += getSmallBlocksRequired(rps.data.length);
      }
    }

    out.write(propertySetStorage);
  }
}
