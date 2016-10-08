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

package jxl.biff.formula;

import jxl.common.Logger;

import jxl.WorkbookSettings;

/**
 * An enumeration detailing the Excel function codes
 */
final class Function
{
  /**
   * The logger
   */
  private static Logger logger = Logger.getLogger(Function.class);

  /**
   * The code which applies to this function
   */
  private final int code;

  /**
   * The property name of this function
   */
  private final String name;

  /**
   * The number of args this function expects
   */
  private final int numArgs;


  /**
   * All available functions.  This attribute is package protected in order
   * to enable the FunctionNames to initialize
   */
  private static Function[] functions = new Function[0];

  /**
   * Constructor
   * Sets the token value and adds this token to the array of all token
   *
   * @param v the biff code for the token
   * @param s the string
   * @param a the number of arguments
   */
  private Function(int v, String s, int a)
  {
    code = v;
    name = s;
    numArgs = a;

    // Grow the array
    Function[] newarray = new Function[functions.length + 1];
    System.arraycopy(functions, 0, newarray, 0, functions.length);
    newarray[functions.length] = this;
    functions = newarray;
  }

  /**
   * Standard hash code method
   *
   * @return the hash code
   */
  public int hashCode()
  {
    return code;
  }

  /**
   * Gets the function code - used when generating token data
   *
   * @return the code
   */
  int getCode()
  {
    return code;
  }

  /**
   * Gets the property name. Used by the FunctionNames object when initializing
   * the locale specific names
   *
   * @return the property name for this function
   */
  String getPropertyName()
  {
    return name;
  }

  /**
   * Gets the function name
   * @param ws the workbook settings
   * @return the function name
   */
  String getName(WorkbookSettings ws)
  {
    FunctionNames fn = ws.getFunctionNames();
    return fn.getName(this);
  }

  /**
   * Gets the number of arguments for this function
   *
   * @return the number of arguments
   */
  int getNumArgs()
  {
    return numArgs;
  }

  /**
   * Gets the type object from its integer value
   *
   * @param v the function value
   * @return the function
   */
  public static Function getFunction(int v)
  {
    Function f = null;

    for (int i = 0; i < functions.length; i++)
    {
      if (functions[i].code == v)
      {
        f = functions[i];
        break;
      }
    }

    return f != null ? f : UNKNOWN;
  }

  /**
   * Gets the type object from its string value.  Used when parsing strings
   *
   * @param v the function name
   * @param ws the workbook settings
   * @return the function
   */
  public static Function getFunction(String v, WorkbookSettings ws)
  {
    FunctionNames fn = ws.getFunctionNames();
    Function f = fn.getFunction(v);
    return f != null ? f : UNKNOWN;
  }

  /**
   * Accessor for all the functions, used by the internationalization
   * work around
   *
   * @return all the functions
   */
  static Function[] getFunctions()
  {
    return functions;
  }

  // The functions

  public static final Function COUNT =
    new Function(0x0, "count", 0xff);
  public static final Function ATTRIBUTE = new Function(0x1, "", 0xff);
  public static final Function ISNA =
    new Function(0x2, "isna", 1);
  public static final Function ISERROR =
    new Function(0x3, "iserror", 1);
  public static final Function SUM =
    new Function(0x4, "sum", 0xff);
  public static final Function AVERAGE =
    new Function(0x5, "average", 0xff);
  public static final Function MIN =
    new Function(0x6, "min", 0xff);
  public static final Function MAX =
    new Function(0x7, "max", 0xff);
  public static final Function ROW =
    new Function(0x8, "row", 0xff);
  public static final Function COLUMN =
    new Function(0x9, "column", 0xff);
  public static final Function NA =
    new Function(0xa, "na", 0);
  public static final Function NPV =
    new Function(0xb, "npv", 0xff);
  public static final Function STDEV =
    new Function(0xc, "stdev", 0xff);
  public static final Function DOLLAR =
    new Function(0xd, "dollar", 2);
  public static final Function FIXED =
    new Function(0xe, "fixed", 0xff);
  public static final Function SIN =
    new Function(0xf, "sin", 1);
  public static final Function COS =
    new Function(0x10, "cos", 1);
  public static final Function TAN =
    new Function(0x11, "tan", 1);
  public static final Function ATAN =
    new Function(0x12, "atan", 1);
  public static final Function PI =
    new Function(0x13, "pi", 0);
  public static final Function SQRT =
    new Function(0x14, "sqrt", 1);
  public static final Function EXP =
    new Function(0x15, "exp", 1);
  public static final Function LN =
    new Function(0x16, "ln", 1);
  public static final Function LOG10 =
    new Function(0x17, "log10", 1);
  public static final Function ABS =
    new Function(0x18, "abs", 1);
  public static final Function INT =
    new Function(0x19, "int", 1);
  public static final Function SIGN =
    new Function(0x1a, "sign", 1);
  public static final Function ROUND =
    new Function(0x1b, "round", 2);
  public static final Function LOOKUP =
    new Function(0x1c, "lookup", 2);
  public static final Function INDEX =
    new Function(0x1d, "index", 3);
  public static final Function REPT =  new Function(0x1e, "rept", 2);
  public static final Function MID =
    new Function(0x1f, "mid", 3);
  public static final Function LEN =
    new Function(0x20, "len", 1);
  public static final Function VALUE =
    new Function(0x21, "value", 1);
  public static final Function TRUE =
    new Function(0x22, "true", 0);
  public static final Function FALSE =
    new Function(0x23, "false", 0);
  public static final Function AND =
    new Function(0x24, "and", 0xff);
  public static final Function OR =
    new Function(0x25, "or", 0xff);
  public static final Function NOT =
    new Function(0x26, "not", 1);
  public static final Function MOD =
    new Function(0x27, "mod", 2);
  public static final Function DCOUNT =
    new Function(0x28, "dcount", 3);
  public static final Function DSUM =
    new Function(0x29, "dsum", 3);
  public static final Function DAVERAGE =
    new Function(0x2a, "daverage", 3);
  public static final Function DMIN =
    new Function(0x2b, "dmin", 3);
  public static final Function DMAX =
    new Function(0x2c, "dmax", 3);
  public static final Function DSTDEV =
    new Function(0x2d, "dstdev", 3);
  public static final Function VAR =
    new Function(0x2e, "var", 0xff);
  public static final Function DVAR =
    new Function(0x2f, "dvar", 3);
  public static final Function TEXT =
    new Function(0x30, "text", 2);
  public static final Function LINEST =
    new Function(0x31, "linest", 0xff);
  public static final Function TREND =
    new Function(0x32, "trend", 0xff);
  public static final Function LOGEST =
    new Function(0x33, "logest", 0xff);
  public static final Function GROWTH =
    new Function(0x34, "growth", 0xff);
  //public static final Function GOTO =  new Function(0x35, "GOTO",);
  //public static final Function HALT =  new Function(0x36, "HALT",);
  public static final Function PV =
    new Function(0x38, "pv", 0xff);
  public static final Function FV =
    new Function(0x39, "fv", 0xff);
  public static final Function NPER =
    new Function(0x3a, "nper", 0xff);
  public static final Function PMT =
    new Function(0x3b, "pmt", 0xff);
  public static final Function RATE =
    new Function(0x3c, "rate", 0xff);
  //public static final Function MIRR =  new Function(0x3d, "MIRR",);
  //public static final Function IRR =  new Function(0x3e, "IRR",);
  public static final Function RAND =
    new Function(0x3f, "rand", 0);
  public static final Function MATCH =
    new Function(0x40, "match", 3);
  public static final Function DATE =
    new Function(0x41, "date", 3);
  public static final Function TIME =
    new Function(0x42, "time", 3);
  public static final Function DAY =
    new Function(0x43, "day", 1);
  public static final Function MONTH =
    new Function(0x44, "month", 1);
  public static final Function YEAR =
    new Function(0x45, "year", 1);
  public static final Function WEEKDAY =
    new Function(0x46, "weekday", 2);
  public static final Function HOUR =
    new Function(0x47, "hour", 1);
  public static final Function MINUTE =
    new Function(0x48, "minute", 1);
  public static final Function SECOND =
    new Function(0x49, "second", 1);
  public static final Function NOW =
    new Function(0x4a, "now", 0);
  public static final Function AREAS =
    new Function(0x4b, "areas", 0xff);
  public static final Function ROWS =
    new Function(0x4c, "rows", 1);
  public static final Function COLUMNS =
    new Function(0x4d, "columns", 0xff);
  public static final Function OFFSET =
    new Function(0x4e, "offset", 0xff);
  //public static final Function ABSREF =  new Function(0x4f, "ABSREF",);
  //public static final Function RELREF =  new Function(0x50, "RELREF",);
  //public static final Function ARGUMENT =  new Function(0x51,"ARGUMENT",);
  public static final Function SEARCH =  new Function(0x52, "search", 0xff);
  public static final Function TRANSPOSE =
    new Function(0x53, "transpose", 0xff);
  public static final Function ERROR =
    new Function(0x54, "error", 1);
  //public static final Function STEP =  new Function(0x55, "STEP",);
  public static final Function TYPE =
    new Function(0x56, "type", 1);
  //public static final Function ECHO =  new Function(0x57, "ECHO",);
  //public static final Function SETNAME =  new Function(0x58, "SETNAME",);
  //public static final Function CALLER =  new Function(0x59, "CALLER",);
  //public static final Function DEREF =  new Function(0x5a, "DEREF",);
  //public static final Function WINDOWS =  new Function(0x5b, "WINDOWS",);
  //public static final Function SERIES =  new Function(0x5c, "SERIES",);
  //public static final Function DOCUMENTS =  new Function(0x5d,"DOCUMENTS",);
  //public static final Function ACTIVECELL =  new Function(0x5e,"ACTIVECELL",);
  //public static final Function SELECTION =  new Function(0x5f,"SELECTION",);
  //public static final Function RESULT =  new Function(0x60, "RESULT",);
  public static final Function ATAN2  =
    new Function(0x61, "atan2", 1);
  public static final Function ASIN =
    new Function(0x62, "asin", 1);
  public static final Function ACOS =
    new Function(0x63, "acos", 1);
  public static final Function CHOOSE =
    new Function(0x64, "choose", 0xff);
  public static final Function HLOOKUP =
    new Function(0x65, "hlookup", 0xff);
  public static final Function VLOOKUP =
    new Function(0x66, "vlookup", 0xff);
  //public static final Function LINKS =  new Function(0x67, "LINKS",);
  //public static final Function INPUT =  new Function(0x68, "INPUT",);
  public static final Function ISREF =
    new Function(0x69, "isref", 1);
  //public static final Function GETFORMULA =  new Function(0x6a,"GETFORMULA",);
  //public static final Function GETNAME =  new Function(0x6b, "GETNAME",);
  //public static final Function SETVALUE =  new Function(0x6c,"SETVALUE",);
  public static final Function LOG =
    new Function(0x6d, "log", 0xff);
  //public static final Function EXEC =  new Function(0x6e, "EXEC",);
  public static final Function CHAR =
    new Function(0x6f, "char", 1);
  public static final Function LOWER =
    new Function(0x70, "lower", 1);
  public static final Function UPPER =
    new Function(0x71, "upper", 1);
  public static final Function PROPER =
    new Function(0x72, "proper", 1);
  public static final Function LEFT =
    new Function(0x73, "left", 0xff);
  public static final Function RIGHT =
    new Function(0x74, "right", 0xff);
  public static final Function EXACT =
    new Function(0x75, "exact", 2);
  public static final Function TRIM =
    new Function(0x76, "trim", 1);
  public static final Function REPLACE =
    new Function(0x77, "replace", 4);
  public static final Function SUBSTITUTE =
    new Function(0x78, "substitute", 0xff);
  public static final Function CODE =
    new Function(0x79, "code", 1);
  //public static final Function NAMES =  new Function(0x7a, "NAMES",);
  //public static final Function DIRECTORY =  new Function(0x7b,"DIRECTORY",);
  public static final Function FIND =
    new Function(0x7c, "find", 0xff);
  public static final Function CELL =
    new Function(0x7d, "cell", 2);
  public static final Function ISERR =
    new Function(0x7e, "iserr", 1);
  public static final Function ISTEXT =
    new Function(0x7f, "istext", 1);
  public static final Function ISNUMBER =
    new Function(0x80, "isnumber", 1);
  public static final Function ISBLANK =
    new Function(0x81, "isblank", 1);
  public static final Function T =
    new Function(0x82, "t", 1);
  public static final Function N =
    new Function(0x83, "n", 1);
  //public static final Function FOPEN =  new Function(0x84, "FOPEN",);
  //public static final Function FCLOSE =  new Function(0x85, "FCLOSE",);
  //public static final Function FSIZE =  new Function(0x86, "FSIZE",);
  //public static final Function FREADLN =  new Function(0x87, "FREADLN",);
  //public static final Function FREAD =  new Function(0x88, "FREAD",);
  //public static final Function FWRITELN =  new Function(0x89,"FWRITELN",);
  //public static final Function FWRITE =  new Function(0x8a, "FWRITE",);
  //public static final Function FPOS =  new Function(0x8b, "FPOS",);
  public static final Function DATEVALUE =
    new Function(0x8c, "datevalue", 1);
  public static final Function TIMEVALUE =
    new Function(0x8d, "timevalue", 1);
  public static final Function SLN =
    new Function(0x8e, "sln", 3);
  public static final Function SYD =
    new Function(0x8f, "syd", 3);
  public static final Function DDB =
    new Function(0x90, "ddb", 0xff);
  //public static final Function GETDEF =  new Function(0x91, "GETDEF",);
  //public static final Function REFTEXT =  new Function(0x92, "REFTEXT",);
  //public static final Function TEXTREF =  new Function(0x93, "TEXTREF",);
  public static final Function INDIRECT =
    new Function(0x94, "indirect", 0xff);
  //public static final Function REGISTER =  new Function(0x95,"REGISTER",);
  //public static final Function CALL =  new Function(0x96, "CALL",);
  //public static final Function ADDBAR =  new Function(0x97, "ADDBAR",);
  //public static final Function ADDMENU =  new Function(0x98, "ADDMENU",);
  //public static final Function ADDCOMMAND =
  // new Function(0x99,"ADDCOMMAND",);
  //public static final Function ENABLECOMMAND =
  // new Function(0x9a,"ENABLECOMMAND",);
  //public static final Function CHECKCOMMAND =
  // new Function(0x9b,"CHECKCOMMAND",);
  //public static final Function RENAMECOMMAND =
  // new Function(0x9c,"RENAMECOMMAND",);
  //public static final Function SHOWBAR =  new Function(0x9d, "SHOWBAR",);
  //public static final Function DELETEMENU =
  //  new Function(0x9e,"DELETEMENU",);
  //public static final Function DELETECOMMAND =
  //  new Function(0x9f,"DELETECOMMAND",);
  //public static final Function GETCHARTITEM =
  //  new Function(0xa0,"GETCHARTITEM",);
  //public static final Function DIALOGBOX =  new Function(0xa1,"DIALOGBOX",);
  public static final Function CLEAN =
    new Function(0xa2, "clean", 1);
  public static final Function MDETERM =
    new Function(0xa3, "mdeterm", 0xff);
  public static final Function MINVERSE =
    new Function(0xa4, "minverse", 0xff);
  public static final Function MMULT =
    new Function(0xa5, "mmult", 0xff);
  //public static final Function FILES =  new Function(0xa6, "FILES",

  public static final Function IPMT =
    new Function(0xa7, "ipmt", 0xff);
  public static final Function PPMT =
    new Function(0xa8, "ppmt", 0xff);
  public static final Function COUNTA =
    new Function(0xa9, "counta", 0xff);
  public static final Function PRODUCT =
    new Function(0xb7, "product", 0xff);
  public static final Function FACT =
    new Function(0xb8, "fact", 1);
  //public static final Function GETCELL =  new Function(0xb9, "GETCELL",);
  //public static final Function GETWORKSPACE =
  //  new Function(0xba,"GETWORKSPACE",);
  //public static final Function GETWINDOW =  new Function(0xbb,"GETWINDOW",);
  //public static final Function GETDOCUMENT =
  //  new Function(0xbc,"GETDOCUMENT",);
  public static final Function DPRODUCT =
    new Function(0xbd, "dproduct", 3);
  public static final Function ISNONTEXT =
    new Function(0xbe, "isnontext", 1);
  //public static final Function GETNOTE =  new Function(0xbf, "GETNOTE",);
  //public static final Function NOTE =  new Function(0xc0, "NOTE",);
  public static final Function STDEVP =
    new Function(0xc1, "stdevp", 0xff);
  public static final Function VARP =
    new Function(0xc2, "varp", 0xff);
  public static final Function DSTDEVP =
    new Function(0xc3, "dstdevp", 0xff);
  public static final Function DVARP =
    new Function(0xc4, "dvarp", 0xff);
  public static final Function TRUNC =
    new Function(0xc5, "trunc", 0xff);
  public static final Function ISLOGICAL =
    new Function(0xc6, "islogical", 1);
  public static final Function DCOUNTA =
    new Function(0xc7, "dcounta", 0xff);
  public static final Function FINDB =
    new Function(0xcd, "findb", 0xff);
  public static final Function SEARCHB =
    new Function(0xce, "searchb", 3);
  public static final Function REPLACEB =
    new Function(0xcf, "replaceb", 4);
  public static final Function LEFTB =
    new Function(0xd0, "leftb", 0xff);
  public static final Function RIGHTB =
    new Function(0xd1, "rightb", 0xff);
  public static final Function MIDB =
    new Function(0xd2, "midb", 3);
  public static final Function LENB =
    new Function(0xd3, "lenb", 1);
  public static final Function ROUNDUP =
    new Function(0xd4, "roundup", 2);
  public static final Function ROUNDDOWN =
    new Function(0xd5, "rounddown", 2);
  public static final Function RANK =
    new Function(0xd8, "rank", 0xff);
  public static final Function ADDRESS =
    new Function(0xdb, "address", 0xff);
  public static final Function AYS360 =
    new Function(0xdc, "days360", 0xff);
  public static final Function ODAY =
    new Function(0xdd, "today", 0);
  public static final Function VDB =
    new Function(0xde, "vdb", 0xff);
  public static final Function MEDIAN =
    new Function(0xe3, "median", 0xff);
  public static final Function SUMPRODUCT =
    new Function(0xe4, "sumproduct", 0xff);
  public static final Function SINH =
    new Function(0xe5, "sinh", 1);
  public static final Function COSH =
    new Function(0xe6, "cosh", 1);
  public static final Function TANH =
    new Function(0xe7, "tanh", 1);
  public static final Function ASINH =
    new Function(0xe8, "asinh", 1);
  public static final Function ACOSH =
    new Function(0xe9, "acosh", 1);
  public static final Function ATANH =
    new Function(0xea, "atanh", 1);
  public static final Function INFO =
    new Function(0xf4, "info", 1);
  public static final Function AVEDEV =
    new Function(0x10d, "avedev", 0XFF);
  public static final Function BETADIST =
    new Function(0x10e, "betadist", 0XFF);
  public static final Function GAMMALN =
    new Function(0x10f, "gammaln", 1);
  public static final Function BETAINV =
    new Function(0x110, "betainv", 0XFF);
  public static final Function BINOMDIST =
    new Function(0x111, "binomdist", 4);
  public static final Function CHIDIST =
    new Function(0x112, "chidist", 2);
  public static final Function CHIINV =
    new Function(0x113, "chiinv", 2);
  public static final Function COMBIN =
    new Function(0x114, "combin", 2);
  public static final Function CONFIDENCE =
    new Function(0x115, "confidence", 3);
  public static final Function CRITBINOM =
    new Function(0x116, "critbinom", 3);
  public static final Function EVEN =
    new Function(0x117, "even", 1);
  public static final Function EXPONDIST =
    new Function(0x118, "expondist", 3);
  public static final Function FDIST =
    new Function(0x119, "fdist", 3);
  public static final Function FINV =
    new Function(0x11a, "finv", 3);
  public static final Function FISHER =
    new Function(0x11b, "fisher", 1);
  public static final Function FISHERINV =
    new Function(0x11c, "fisherinv", 1);
  public static final Function FLOOR =
    new Function(0x11d, "floor", 2);
  public static final Function GAMMADIST =
    new Function(0x11e, "gammadist", 4);
  public static final Function GAMMAINV =
    new Function(0x11f, "gammainv", 3);
  public static final Function CEILING =
    new Function(0x120, "ceiling", 2);
  public static final Function HYPGEOMDIST =
    new Function(0x121, "hypgeomdist", 4);
  public static final Function LOGNORMDIST =
    new Function(0x122, "lognormdist", 3);
  public static final Function LOGINV =
    new Function(0x123, "loginv", 3);
  public static final Function NEGBINOMDIST =
    new Function(0x124, "negbinomdist", 3);
  public static final Function NORMDIST =
    new Function(0x125, "normdist", 4);
  public static final Function NORMSDIST =
    new Function(0x126, "normsdist", 1);
  public static final Function NORMINV =
    new Function(0x127, "norminv", 3);
  public static final Function NORMSINV =
    new Function(0x128, "normsinv", 1);
  public static final Function STANDARDIZE =
    new Function(0x129, "standardize", 3);
  public static final Function ODD =
    new Function(0x12a, "odd", 1);
  public static final Function PERMUT =
    new Function(0x12b, "permut", 2);
  public static final Function POISSON =
    new Function(0x12c, "poisson", 3);
  public static final Function TDIST =
    new Function(0x12d, "tdist", 3);
  public static final Function WEIBULL =
    new Function(0x12e, "weibull", 4);
  public static final Function SUMXMY2 =
    new Function(303, "sumxmy2", 0xff);
  public static final Function SUMX2MY2 =
    new Function(304, "sumx2my2", 0xff);
  public static final Function SUMX2PY2 =
    new Function(305, "sumx2py2", 0xff);
  public static final Function CHITEST =
    new Function(0x132, "chitest", 0xff);
  public static final Function CORREL =
    new Function(0x133, "correl", 0xff);
  public static final Function COVAR =
    new Function(0x134, "covar", 0xff);
  public static final Function FORECAST =
    new Function(0x135, "forecast", 0xff);
  public static final Function FTEST =
    new Function(0x136, "ftest", 0xff);
  public static final Function INTERCEPT =
    new Function(0x137, "intercept", 0xff);
  public static final Function PEARSON =
    new Function(0x138, "pearson", 0xff);
  public static final Function RSQ =
    new Function(0x139, "rsq", 0xff);
  public static final Function STEYX =
    new Function(0x13a, "steyx", 0xff);
  public static final Function SLOPE =
    new Function(0x13b, "slope", 2);
  public static final Function TTEST =
    new Function(0x13c, "ttest", 0xff);
  public static final Function PROB =
    new Function(0x13d, "prob", 0xff);
  public static final Function DEVSQ =
    new Function(0x13e, "devsq", 0xff);
  public static final Function GEOMEAN =
    new Function(0x13f, "geomean", 0xff);
  public static final Function HARMEAN =
    new Function(0x140, "harmean", 0xff);
  public static final Function SUMSQ =
    new Function(0x141, "sumsq", 0xff);
  public static final Function KURT =
    new Function(0x142, "kurt", 0xff);
  public static final Function SKEW =
    new Function(0x143, "skew", 0xff);
  public static final Function ZTEST =
    new Function(0x144, "ztest", 0xff);
  public static final Function LARGE =
    new Function(0x145, "large", 0xff);
  public static final Function SMALL =
    new Function(0x146, "small", 0xff);
  public static final Function QUARTILE =
    new Function(0x147, "quartile", 0xff);
  public static final Function PERCENTILE =
    new Function(0x148, "percentile", 0xff);
  public static final Function PERCENTRANK =
    new Function(0x149, "percentrank", 0xff);
  public static final Function MODE =
    new Function(0x14a, "mode", 0xff);
  public static final Function TRIMMEAN =
    new Function(0x14b, "trimmean", 0xff);
  public static final Function TINV =
    new Function(0x14c, "tinv", 2);
  public static final Function CONCATENATE =
    new Function(0x150, "concatenate", 0xff);
  public static final Function POWER =
    new Function(0x151, "power", 2);
  public static final Function RADIANS  =
    new Function(0x156, "radians", 1);
  public static final Function DEGREES  =
    new Function(0x157, "degrees", 1);
  public static final Function SUBTOTAL =
    new Function(0x158, "subtotal", 0xff);
  public static final Function SUMIF  =
    new Function(0x159, "sumif", 0xff);
  public static final Function COUNTIF  =
    new Function(0x15a, "countif", 2);
  public static final Function COUNTBLANK =
    new Function(0x15b, "countblank", 1);
  public static final Function HYPERLINK =
    new Function(0x167, "hyperlink", 2);
  public static final Function AVERAGEA =
    new Function(0x169, "averagea", 0xff);
  public static final Function MAXA =
    new Function(0x16a, "maxa", 0xff);
  public static final Function MINA =
    new Function(0x16b, "mina", 0xff);
  public static final Function STDEVPA =
    new Function(0x16c, "stdevpa", 0xff);
  public static final Function VARPA =
    new Function(0x16d, "varpa", 0xff);
  public static final Function STDEVA =
    new Function(0x16e, "stdeva", 0xff);
  public static final Function VARA =
    new Function(0x16f, "vara", 0xff);

  // If token.  This is not an excel assigned number, but one made up
  // in order that the if command may be recognized
  public static final Function IF =
    new Function(0xfffe, "if", 0xff);

  // Unknown token
  public static final Function UNKNOWN = new Function(0xffff, "", 0);
}
