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

package jxl.format;

/**
 * Enumeration type which contains the available excel paper sizes and their
 * codes
 */
public final class PaperSize
{
	private static final int LAST_PAPER_SIZE = 89;
	/**
	 * The excel encoding
	 */
	private int val;

	/**
	 * The paper sizes
	 */
	private static PaperSize[] paperSizes = new PaperSize[LAST_PAPER_SIZE + 1];

	/**
	 * Constructor
	 */
	private PaperSize(int v, boolean growArray)
	{
		val = v;

		if (v >= paperSizes.length && growArray)
		{
			// Grow the array and add this to it
			PaperSize[] newarray = new PaperSize[v + 1];
			System.arraycopy(paperSizes, 0, newarray, 0, paperSizes.length);
			paperSizes = newarray;
		}
		if (v < paperSizes.length)
		{
			paperSizes[v] = this;
		}
	}

	/**
	 * Constructor
	 */
	private PaperSize(int v)
	{
		this(v, true);
	}

	/**
	 * Accessor for the internal binary value association with this paper size
	 *
	 * @return the internal value
	 */
	public int getValue()
	{
		return val;
	}

	/**
	 * Gets the paper size for a specific value
	 *
	 * @param val the value
	 * @return the paper size
	 */
	public static PaperSize getPaperSize(int val)
	{
		PaperSize p = val > paperSizes.length - 1 ? null : paperSizes[val];
		return p == null ? new PaperSize(val, false) : p;
	}

	/** US Letter 8.5 x 11" */
	public static final PaperSize UNDEFINED = new PaperSize(0);

	/** US Letter 8.5 x 11" */
	public static final PaperSize LETTER = new PaperSize(1);

	/** Letter small 8.5" × 11" */
	public static final PaperSize LETTER_SMALL = new PaperSize(2);

	/** Tabloid 11" x 17" */
	public static final PaperSize TABLOID = new PaperSize(3);

	/** Leger 17" x 11" */
	public static final PaperSize LEDGER = new PaperSize(4);

	/** US Legal 8.5" x 14" */
	public static final PaperSize LEGAL = new PaperSize(5);

	/** Statement 5.5" x 8.5" */
	public static final PaperSize STATEMENT = new PaperSize(6);

	/** Executive 7.25" x 10.5" */
	public static final PaperSize EXECUTIVE = new PaperSize(7);

	/** A3 297mm x 420mm */
	public static final PaperSize A3 = new PaperSize(8);

	/** A4 210mm x 297mm */
	public static final PaperSize A4 = new PaperSize(9);

	/** A4 Small 210mm x 297 mm */
	public static final PaperSize A4_SMALL = new PaperSize(10);

	/** A5 148mm x 210mm */
	public static final PaperSize A5 = new PaperSize(11);

	/** B4 (JIS) 257mm x 364mm */
	public static final PaperSize B4 = new PaperSize(12);

	/** B5 (JIS) 182mm x 257mm */
	public static final PaperSize B5 = new PaperSize(13);

	/** Folio 8.5" x 13" */
	public static final PaperSize FOLIO = new PaperSize(14);

	/** Quarto 215mm x 275mm */
	public static final PaperSize QUARTO = new PaperSize(15);

	/** 10" x 14" */
	public static final PaperSize SIZE_10x14 = new PaperSize(16);

	/** 11" x 17" */
	public static final PaperSize SIZE_10x17 = new PaperSize(17);

	/** NOTE 8.5" x 11" */
	public static final PaperSize NOTE = new PaperSize(18);

	/** Envelope #9 3 7/8" x 8 7/8" */
	public static final PaperSize ENVELOPE_9 = new PaperSize(19);

	/** Envelope #10 4 1/8" x 9.5" */
	public static final PaperSize ENVELOPE_10 = new PaperSize(20);

	/** Envelope #11 4.5" x 10 3/8" */
	public static final PaperSize ENVELOPE_11 = new PaperSize(21);

	/** Envelope #12 4.75" x 11" */
	public static final PaperSize ENVELOPE_12 = new PaperSize(22);

	/** Envelope #14 5" x 11.5" */
	public static final PaperSize ENVELOPE_14 = new PaperSize(23);

	/** C 17" x 22" */
	public static final PaperSize C = new PaperSize(24);

	/** D 22" x 34" */
	public static final PaperSize D = new PaperSize(25);

	/** E 34" x 44" */
	public static final PaperSize E = new PaperSize(26);

	/** Envelope DL 110mm × 220mm */
	public static final PaperSize ENVELOPE_DL = new PaperSize(27);

	/** Envelope C5 162mm × 229mm */
	public static final PaperSize ENVELOPE_C5 = new PaperSize(28);

	/** Envelope C3 324mm × 458mm */
	public static final PaperSize ENVELOPE_C3 = new PaperSize(29);

	/** Envelope C4 229mm × 324mm */
	public static final PaperSize ENVELOPE_C4 = new PaperSize(30);

	/** Envelope C6 114mm × 162mm */
	public static final PaperSize ENVELOPE_C6 = new PaperSize(31);

	/** Envelope C6/C5 114mm × 229mm */
	public static final PaperSize ENVELOPE_C6_C5 = new PaperSize(32);

	/** B4 (ISO) 250mm × 353mm */
	public static final PaperSize B4_ISO = new PaperSize(33);

	/** B5 (ISO) 176mm × 250mm */
	public static final PaperSize B5_ISO = new PaperSize(34);

	/** B6 (ISO) 125mm × 176mm */
	public static final PaperSize B6_ISO = new PaperSize(35);

	/** Envelope Italy 110mm × 230mm */
	public static final PaperSize ENVELOPE_ITALY = new PaperSize(36);

	/** Envelope Monarch 3 7/8" × 7.5" */
	public static final PaperSize ENVELOPE_MONARCH = new PaperSize(37);

	/** 6.75 Envelope 3 5/8" × 6.5" */
	public static final PaperSize ENVELOPE_6_75 = new PaperSize(38);

	/** US Standard Fanfold 14 7/8" × 11" */
	public static final PaperSize US_FANFOLD = new PaperSize(39);

	/** German Std. Fanfold 8.5" × 12" */
	public static final PaperSize GERMAN_FANFOLD = new PaperSize(40);

	/** German Legal Fanfold 8.5" × 13" */
	public static final PaperSize GERMAN_LEGAL_FANFOLD = new PaperSize(41);

	/** B4 (ISO) 250mm × 353mm */
	public static final PaperSize B4_ISO_2 = new PaperSize(42);

	/** Japanese Postcard 100mm × 148mm */
	public static final PaperSize JAPANESE_POSTCARD = new PaperSize(43);

	/** 9×11 9" × 11" */
	public static final PaperSize SIZE_9x11 = new PaperSize(44);

	/** 10×11 10" × 11" */
	public static final PaperSize SIZE_10x11 = new PaperSize(45);

	/** 15×11 15" × 11" */
	public static final PaperSize SIZE_15x11 = new PaperSize(46);

	/** Envelope Invite 220mm × 220mm */
	public static final PaperSize ENVELOPE_INVITE = new PaperSize(47);

	/* 48 & 49 Undefined */

	/** Letter Extra 9.5" × 12" */
	public static final PaperSize LETTER_EXTRA = new PaperSize(50);

	/** Legal Extra 9.5" × 15" */
	public static final PaperSize LEGAL_EXTRA = new PaperSize(51);

	/** Tabloid Extra 11 11/16" × 18" */
	public static final PaperSize TABLOID_EXTRA = new PaperSize(52);

	/** A4 Extra 235mm × 322mm */
	public static final PaperSize A4_EXTRA = new PaperSize(53);

	/** Letter Transverse 8.5" × 11" */
	public static final PaperSize LETTER_TRANSVERSE = new PaperSize(54);

	/** A4 Transverse 210mm × 297mm */
	public static final PaperSize A4_TRANSVERSE = new PaperSize(55);

	/** Letter Extra Transv. 9.5" × 12" */
	public static final PaperSize LETTER_EXTRA_TRANSVERSE = new PaperSize(56);

	/** Super A/A4 227mm × 356mm */
	public static final PaperSize SUPER_A_A4 = new PaperSize(57);

	/** Super B/A3 305mm × 487mm */
	public static final PaperSize SUPER_B_A3 = new PaperSize(58);

	/** Letter Plus 8.5" x 12 11/16" */
	public static final PaperSize LETTER_PLUS = new PaperSize(59);

	/** A4 Plus 210mm × 330mm */
	public static final PaperSize A4_PLUS = new PaperSize(60);

	/** A5 Transverse 148mm × 210mm */
	public static final PaperSize A5_TRANSVERSE = new PaperSize(61);

	/** B5 (JIS) Transverse 182mm × 257mm */
	public static final PaperSize B5_TRANSVERSE = new PaperSize(62);

	/** A3 Extra 322mm × 445mm */
	public static final PaperSize A3_EXTRA = new PaperSize(63);

	/** A5 Extra 174mm × 235mm */
	public static final PaperSize A5_EXTRA = new PaperSize(64);

	/** B5 (ISO) Extra 201mm × 276mm */
	public static final PaperSize B5_EXTRA = new PaperSize(65);

	/** A2 420mm × 594mm */
	public static final PaperSize A2 = new PaperSize(66);

	/** A3 Transverse 297mm × 420mm */
	public static final PaperSize A3_TRANSVERSE = new PaperSize(67);

	/** A3 Extra Transverse 322mm × 445mm */
	public static final PaperSize A3_EXTRA_TRANSVERSE = new PaperSize(68);

	/** Dbl. Japanese Postcard 200mm × 148mm */
	public static final PaperSize DOUBLE_JAPANESE_POSTCARD = new PaperSize(69);

	/** A6 105mm × 148mm */
	public static final PaperSize A6 = new PaperSize(70);

	/* 71 - 74 undefined */

	/** Letter Rotated 11" × 8.5" */
	public static final PaperSize LETTER_ROTATED = new PaperSize(75);

	/** A3 Rotated 420mm × 297mm */
	public static final PaperSize A3_ROTATED = new PaperSize(76);

	/** A4 Rotated 297mm × 210mm */
	public static final PaperSize A4_ROTATED = new PaperSize(77);

	/** A5 Rotated 210mm × 148mm */
	public static final PaperSize A5_ROTATED = new PaperSize(78);

	/** B4 (JIS) Rotated 364mm × 257mm */
	public static final PaperSize B4_ROTATED = new PaperSize(79);

	/** B5 (JIS) Rotated 257mm × 182mm */
	public static final PaperSize B5_ROTATED = new PaperSize(80);

	/** Japanese Postcard Rot. 148mm × 100mm */
	public static final PaperSize JAPANESE_POSTCARD_ROTATED = new PaperSize(81);

	/** Dbl. Jap. Postcard Rot. 148mm × 200mm */
	public static final PaperSize DOUBLE_JAPANESE_POSTCARD_ROTATED = new PaperSize(82);

	/** A6 Rotated 148mm × 105mm */
	public static final PaperSize A6_ROTATED = new PaperSize(83);

	/* 84 - 87 undefined */

	/** B6 (JIS) 128mm × 182mm */
	public static final PaperSize B6 = new PaperSize(88);

	/** B6 (JIS) Rotated 182mm × 128mm */
	public static final PaperSize B6_ROTATED = new PaperSize(89);
}
